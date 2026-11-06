import sys
import os
import io
import time
import socket
import threading
import subprocess
import webbrowser
import urllib.request

# Ensure sys.stdout and sys.stderr are valid objects in windowed/no-console mode
class SafeWriter:
    def write(self, s):
        pass
    def flush(self):
        pass
    def isatty(self):
        return False

if sys.stdout is None:
    sys.stdout = SafeWriter()
elif not hasattr(sys.stdout, "isatty"):
    sys.stdout.isatty = lambda: False

if sys.stderr is None:
    sys.stderr = SafeWriter()
elif not hasattr(sys.stderr, "isatty"):
    sys.stderr.isatty = lambda: False

if sys.stdin is None:
    sys.stdin = io.StringIO()

if sys.platform == 'win32':
    try:
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def find_free_port(start_port=8000, max_port=8099):
    """Find the first available free port to avoid Errno 10048."""
    for port in range(start_port, max_port + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('127.0.0.1', port))
                return port
            except OSError:
                continue
    return 8000

def find_app_browser():
    """Look for Edge or Chrome executable for dedicated app-mode window."""
    candidate_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Microsoft\Edge\Application\msedge.exe"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Google\Chrome\Application\chrome.exe"),
    ]
    for p in candidate_paths:
        if p and os.path.exists(p):
            return p
    return None

def launch_desktop_window(url):
    """Launch dedicated frameless desktop window in app mode with isolated profile."""
    browser_exe = find_app_browser()
    if browser_exe:
        profile_dir = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "fti_desktop_profile")
        try:
            os.makedirs(profile_dir, exist_ok=True)
        except Exception:
            pass

        cmd = [
            browser_exe,
            f"--app={url}",
            f"--user-data-dir={profile_dir}",
            "--window-size=1420,900",
            "--window-position=center",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-extensions"
        ]
        try:
            proc = subprocess.Popen(cmd)
            return proc
        except Exception:
            pass

    # Fallback to standard browser
    try:
        webbrowser.open(url)
    except Exception:
        pass
    return None

def main():
    from logger_util import get_system_logger, PRIMARY_LOG_FILE
    logger = get_system_logger()

    port = find_free_port(start_port=8000, max_port=8099)
    url = f"http://127.0.0.1:{port}"

    from app import app, set_server_instance, get_last_heartbeat, get_unloading_time
    import uvicorn

    config = uvicorn.Config(
        app,
        host="127.0.0.1",
        port=port,
        log_config=None,
        access_log=False
    )
    server = uvicorn.Server(config)
    set_server_instance(server)

    # Start Uvicorn in background thread
    server_thread = threading.Thread(target=server.run, daemon=True)
    server_thread.start()

    # Wait for HTTP server to become responsive
    is_ready = False
    for _ in range(80):
        time.sleep(0.1)
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/heartbeat", timeout=0.5) as r:
                if r.status == 200:
                    is_ready = True
                    break
        except Exception:
            continue

    if is_ready:
        logger.info("=================================================================", component="MAIN")
        logger.info(f"   FTI IT Asset Inventory Compare - MVP Desktop Running", component="MAIN")
        logger.info(f"   URL: {url}", component="MAIN")
        logger.info(f"   Log File: {PRIMARY_LOG_FILE}", component="MAIN")
        logger.info("=================================================================", component="MAIN")
    else:
        logger.error(f"Failed to start backend server on port {port}", component="MAIN")

    # Launch desktop window
    browser_proc = launch_desktop_window(url)

    # Monitor lifecycle (Clean shutdown on window close or exit button)
    try:
        startup_grace_period = 45.0  # seconds grace period for initial app launch
        startup_time = time.time()
        while not server.should_exit:
            time.sleep(1)
            now = time.time()

            # 1. Check if client explicitly sent unload beacon and didn't reconnect within 5 seconds
            unload_t = get_unloading_time()
            if unload_t is not None and (now - unload_t) > 5.0:
                logger.info("Desktop window closed (unload beacon confirmed). Shutting down server...", component="MAIN")
                server.should_exit = True
                break

            # 2. Check heartbeat watchdog timeout (safety monitor if browser crashed or killed)
            if (now - startup_time) > startup_grace_period:
                last_hb = get_last_heartbeat()
                if (now - last_hb) > 30.0:
                    logger.info("No heartbeat received from UI for >30s. Shutting down server...", component="MAIN")
                    server.should_exit = True
                    break

    except (KeyboardInterrupt, SystemExit):
        logger.info("Shutdown signal received. Exiting...", component="MAIN")
        server.should_exit = True
    finally:
        server.should_exit = True
        if browser_proc is not None and browser_proc.poll() is None:
            try:
                browser_proc.terminate()
            except Exception:
                pass

if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, SystemExit):
        sys.exit(0)
    except Exception as e:
        import traceback
        err_msg = traceback.format_exc()
        try:
            from logger_util import get_system_logger
            logger = get_system_logger()
            logger.error(f"Fatal Startup Crash: {err_msg}", component="CRASH")
        except Exception:
            pass
        log_path = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "fti_fatal_error.log")
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] FATAL CRASH:\n{err_msg}\n")
        except Exception:
            pass
        sys.exit(1)
