import argparse
import glob
import os
import signal
import subprocess
import sys
import time
from pathlib import Path


FACE_AUTH_FAILURE_EXIT_CODE = 86


def _start_controller(script_path: Path, label: str, extra_args=None) -> subprocess.Popen:
    env = os.environ.copy()
    env.setdefault("PYTHONUNBUFFERED", "1")

    if extra_args is None:
        extra_args = []

    creationflags = 0
    if os.name == "nt":
        creationflags = subprocess.CREATE_NEW_PROCESS_GROUP

    process = subprocess.Popen(
        [sys.executable, str(script_path), *extra_args],
        cwd=str(script_path.parent),
        env=env,
        creationflags=creationflags,
    )
    print(f"[MAIN] Started {label} (pid={process.pid})")
    return process


def _normalize_username(raw_value: str) -> str:
    return raw_value.strip().lower().replace(" ", "_")


def _is_user_registered(known_faces_dir: Path, username: str) -> bool:
    patterns = [
        str(known_faces_dir / f"{username}.jpg"),
        str(known_faces_dir / f"{username}.jpeg"),
        str(known_faces_dir / f"{username}.png"),
        str(known_faces_dir / f"{username}.bmp"),
        str(known_faces_dir / f"{username}_*.jpg"),
        str(known_faces_dir / f"{username}_*.jpeg"),
        str(known_faces_dir / f"{username}_*.png"),
        str(known_faces_dir / f"{username}_*.bmp"),
    ]
    for pattern in patterns:
        if glob.glob(pattern):
            return True
    return False


def _register_user_if_needed(base_dir: Path, username: str) -> bool:
    known_faces_dir = base_dir / "known_faces"
    known_faces_dir.mkdir(parents=True, exist_ok=True)

    if _is_user_registered(known_faces_dir, username):
        print(f"[MAIN] Face data already exists for '{username}'. Skipping registration.")
        return True

    register_script = base_dir / "face_auth_register.py"
    if not register_script.exists():
        print(f"[MAIN] Missing registration script: {register_script}")
        return False

    print(f"[MAIN] No face data found for '{username}'. Starting registration...")
    result = subprocess.run(
        [sys.executable, str(register_script), username, "--count", "15"],
        cwd=str(base_dir),
    )

    if result.returncode != 0:
        print(f"[MAIN] Registration failed with code {result.returncode}")
        return False

    if not _is_user_registered(known_faces_dir, username):
        print("[MAIN] Registration finished, but no images were found. Aborting.")
        return False

    print(f"[MAIN] Registration complete for '{username}'.")
    return True


def _terminate_process(process: subprocess.Popen, label: str):
    if process.poll() is not None:
        return

    print(f"[MAIN] Stopping {label}...")
    try:
        if os.name == "nt":
            process.send_signal(signal.CTRL_BREAK_EVENT)
        else:
            process.terminate()
    except Exception:
        process.terminate()

    try:
        process.wait(timeout=4)
    except subprocess.TimeoutExpired:
        print(f"[MAIN] Force-killing {label}")
        process.kill()


def main():
    parser = argparse.ArgumentParser(description="Run registration check + voice + gesture controllers")
    parser.add_argument("--username", default=None, help="Username to verify/register face data")
    parser.add_argument(
        "-fd",
        "--face-disabled",
        action="store_true",
        help="Run gesture controller without face authentication",
    )
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    gesture_script = base_dir / "gesture_Controller_debug.py"
    voice_script = base_dir / "voice_Controller_omega.py"

    if not gesture_script.exists():
        raise FileNotFoundError(f"Missing: {gesture_script}")
    if not voice_script.exists():
        raise FileNotFoundError(f"Missing: {voice_script}")

    raw_username = args.username
    if not raw_username:
        raw_username = input("Enter username for face auth: ")

    username = _normalize_username(raw_username)
    if not username:
        print("[MAIN] Username cannot be empty.")
        return

    if args.face_disabled:
        print("[MAIN] FACE AUTH disabled by flag (-fd).")
    else:
        if not _register_user_if_needed(base_dir, username):
            return

    print("[MAIN] Launching gesture and voice controllers...")
    print("[MAIN] Press Ctrl+C in this terminal to stop both.")

    voice_proc = _start_controller(voice_script, "voice controller")
    # Small stagger helps camera and audio init avoid startup collisions.
    time.sleep(1.2)
    gesture_extra_args = ["-fd"] if args.face_disabled else []
    gesture_proc = _start_controller(gesture_script, "gesture controller", gesture_extra_args)

    should_restart_without_face_auth = False

    try:
        while True:
            gesture_code = gesture_proc.poll()
            voice_code = voice_proc.poll()

            if gesture_code is not None:
                print(f"[MAIN] Gesture controller exited with code {gesture_code}")
                if (
                    gesture_code == FACE_AUTH_FAILURE_EXIT_CODE
                    and not args.face_disabled
                ):
                    print("[MAIN] Face authentication failed. Switching to -fd mode.")
                    should_restart_without_face_auth = True
                break
            if voice_code is not None:
                print(f"[MAIN] Voice controller exited with code {voice_code}")
                break

            time.sleep(0.6)

    except KeyboardInterrupt:
        print("\n[MAIN] Keyboard interrupt received.")

    finally:
        _terminate_process(gesture_proc, "gesture controller")
        _terminate_process(voice_proc, "voice controller")
        print("[MAIN] Shutdown complete.")

    if should_restart_without_face_auth:
        relaunch_cmd = [sys.executable, str(Path(__file__).resolve()), "-fd", "--username", username]
        print(f"[MAIN] Relaunching without face auth: {' '.join(relaunch_cmd)}")
        os.execv(sys.executable, relaunch_cmd)


if __name__ == "__main__":
    main()
