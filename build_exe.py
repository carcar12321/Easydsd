"""
build_exe.py  -  easydsd v0.2 EXE 빌더
google-generativeai, grpc 포함 완전 패키징
"""
import subprocess, sys, os

print("=" * 55)
print("  easydsd v0.2 EXE 빌더")
print("  (Gemini AI 기능 포함)")
print("=" * 55)

# ── 1. 필요 라이브러리 설치 ────────────────────────────────
print("\n[1/3] 라이브러리 설치 중...")
pkgs = [
    "flask",
    "openpyxl",
    "pyinstaller",
    "google-generativeai",
    "grpcio",
    "grpcio-status",
]
subprocess.check_call(
    [sys.executable, "-m", "pip", "install"] + pkgs + ["--quiet", "--upgrade"]
)
print("      설치 완료!")

# ── 2. PyInstaller 빌드 ───────────────────────────────────
print("\n[2/3] EXE 빌드 중... (3~10분 소요)")
cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--noconsole",
    "--name", "easydsd",

    # Flask / Werkzeug
    "--hidden-import", "flask",
    "--hidden-import", "werkzeug",
    "--hidden-import", "werkzeug.serving",
    "--hidden-import", "werkzeug.routing",
    "--hidden-import", "werkzeug.exceptions",
    "--collect-all", "flask",
    "--collect-all", "werkzeug",
    "--collect-all", "jinja2",

    # openpyxl
    "--hidden-import", "openpyxl",
    "--hidden-import", "openpyxl.styles",
    "--hidden-import", "openpyxl.utils",
    "--collect-all", "openpyxl",

    # Gemini AI (핵심)
    "--hidden-import", "google.generativeai",
    "--hidden-import", "google.api_core",
    "--hidden-import", "google.auth",
    "--hidden-import", "google.protobuf",
    "--collect-all", "google",
    "--collect-all", "google.generativeai",
    "--collect-all", "grpc",
    "--collect-all", "grpcio",

    # 기타
    "--hidden-import", "click",
    "--hidden-import", "jinja2",

    "--clean",
    "dart_gui.py"
]

result = subprocess.run(cmd)

if result.returncode != 0:
    print("\n[오류] 빌드 실패!")
    input("\n엔터를 누르면 창이 닫힙니다...")
    sys.exit(1)

# ── 3. 완료 ───────────────────────────────────────────────
exe_path = os.path.join("dist", "easydsd.exe")
print("\n" + "=" * 55)
print("  빌드 완료!")
if os.path.exists(exe_path):
    size_mb = os.path.getsize(exe_path) / 1024 / 1024
    print(f"  파일: {exe_path}  ({size_mb:.0f} MB)")
print()
print("  easydsd.exe 하나만 배포하시면 됩니다.")
print("=" * 55)

if os.path.exists("dist"):
    try: os.startfile("dist")
    except: pass

input("\n엔터를 누르면 창이 닫힙니다...")
