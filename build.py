import subprocess
import sys
import shutil
from pathlib import Path
import traceback


ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
MAIN = SRC / "Main.py"


requirements =  "requirements.txt"




try:
    subprocess.check_call([
        sys.executable,
        "-m",
        "pip",
        "install",
        "-r",
        str(requirements)
    ])


    for folder in ("build", "dist"):
        path = Path(folder)
        if path.exists():
            shutil.rmtree(path)
            print(f"Deleted: {path}")
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--name", "RobotStructuralPlanners",
        "--paths", str(SRC),
        str(MAIN),
    ]

    subprocess.run(cmd, check=True)

    print(f"\nBuild complete: {ROOT / 'dist' / 'MyApp.exe'}")

    for folder in ("build",):
        path = Path(folder)
        if path.exists():
            shutil.rmtree(path)
            print(f"Deleted: {path}")
except Exception:
    traceback.print_exc()

for file in Path(".").glob("*.spec"):
    file.unlink()
    print(f"Deleted: {file}")
name = input("Press Enter to close .... ")
