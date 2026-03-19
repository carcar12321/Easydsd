"""
build_exe.py  -  easydsd EXE 빌더
bat 파일이 이 스크립트를 호출합니다.
직접 실행도 가능:  python build_exe.py
"""
import subprocess, sys, os

print("=" * 55)
print("  easydsd EXE 빌더")
print("=" * 55)

# 1. 필요 라이브러리 설치
print("\n[1/3] 라이브러리 설치 중...")
pkgs = ["flask", "openpyxl", "pyinstaller"]
subprocess.check_call(
    [sys.executable, "-m", "pip", "install"] + pkgs + ["--quiet", "--upgrade"]
)
print("      설치 완료!")

# 2. PyInstaller 실행
print("\n[2/3] EXE 빌드 중... (1~3분 소요)")
cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--noconsole",
    "--name", "easydsd",
    "--hidden-import", "openpyxl",
    "--hidden-import", "openpyxl.styles",
    "--hidden-import", "openpyxl.utils",
    "--hidden-import", "flask",
    "--hidden-import", "werkzeug",
    "--hidden-import", "werkzeug.serving",
    "--hidden-import", "werkzeug.routing",
    "--hidden-import", "werkzeug.exceptions",
    "--hidden-import", "jinja2",
    "--hidden-import", "click",
    "--collect-all", "flask",
    "--collect-all", "werkzeug",
    "--collect-all", "jinja2",
    "--clean",
    "dart_gui.py"
]
result = subprocess.run(cmd)

if result.returncode != 0:
    print("\n[오류] 빌드 실패!")
    input("엔터를 누르면 창이 닫힙니다...")
    sys.exit(1)

# 3. 완료
exe_path = os.path.join("dist", "easydsd.exe")
print("\n" + "=" * 55)
print("  빌드 완료!")
print(f"  생성된 파일: {exe_path}")
print()
print("  easydsd.exe 하나만 배포하시면 됩니다.")
print("  Python이 없는 PC에서도 바로 실행됩니다.")
print("=" * 55)

# dist 폴더 열기 (Windows)
if os.path.exists("dist"):
    try:
        os.startfile("dist")
    except Exception:
        pass

input("\n엔터를 누르면 창이 닫힙니다...")
