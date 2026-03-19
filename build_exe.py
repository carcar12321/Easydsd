"""
build_exe.py  -  easydsd v0.6 EXE 빌더
google-generativeai 완전 패키징 버전
"""
import subprocess, sys, os

print("=" * 58)
print("  easydsd v0.6 EXE 빌더 (Gemini AI 완전 포함)")
print("=" * 58)

# ── 1. 필요 라이브러리 설치 ────────────────────────────────────
print("\n[1/3] 라이브러리 설치 중...")
pkgs = [
    "flask",
    "openpyxl",
    "pyinstaller",
    # Gemini AI 관련 (순서 중요)
    "google-generativeai",
    "google-api-core",
    "grpcio",
    "grpcio-status",
    "pydantic",
]
subprocess.check_call(
    [sys.executable, "-m", "pip", "install"] + pkgs + ["--upgrade", "-q"]
)
print("      완료!")

# ── 2. PyInstaller 빌드 ───────────────────────────────────────
print("\n[2/3] EXE 빌드 중... (5~15분 소요)")

cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--noconsole",
    "--name", "easydsd",

    # ── Flask / Jinja2 ──────────────────────────────────────
    "--hidden-import", "flask",
    "--hidden-import", "werkzeug",
    "--hidden-import", "werkzeug.serving",
    "--hidden-import", "werkzeug.routing",
    "--hidden-import", "werkzeug.exceptions",
    "--hidden-import", "jinja2",
    "--hidden-import", "click",
    "--collect-all",   "flask",
    "--collect-all",   "werkzeug",
    "--collect-all",   "jinja2",

    # ── openpyxl ─────────────────────────────────────────────
    "--hidden-import", "openpyxl",
    "--hidden-import", "openpyxl.styles",
    "--hidden-import", "openpyxl.utils",
    "--collect-all",   "openpyxl",

    # ── google-generativeai 핵심 ──────────────────────────────
    "--hidden-import", "google.generativeai",
    "--hidden-import", "google.ai.generativelanguage",
    "--hidden-import", "google.api_core",
    "--hidden-import", "google.auth",
    "--hidden-import", "google.auth.transport",
    "--hidden-import", "google.protobuf",
    "--hidden-import", "pydantic",
    "--collect-all",   "google",
    "--collect-all",   "google.generativeai",
    "--collect-all",   "pydantic",

    # ── gRPC (google-generativeai 의존성) ─────────────────────
    "--hidden-import", "grpc",
    "--hidden-import", "grpc._channel",
    "--hidden-import", "grpc._interceptor",
    "--collect-all",   "grpc",
    "--collect-all",   "grpcio",

    # ── 메타데이터 복사 (importlib.metadata 사용 패키지 필수) ───
    "--copy-metadata", "google-generativeai",
    "--copy-metadata", "google-api-core",
    "--copy-metadata", "grpcio",
    "--copy-metadata", "pydantic",

    "--clean",
    "dart_gui.py",
]

result = subprocess.run(cmd)

if result.returncode != 0:
    print("\n[오류] 빌드 실패!")
    print("오류 내용을 캡처해서 개발자에게 문의하세요.")
    input("\n엔터를 누르면 창이 닫힙니다...")
    sys.exit(1)

# ── 3. 완료 ───────────────────────────────────────────────────
exe_path = os.path.join("dist", "easydsd.exe")
print("\n" + "=" * 58)
print("  빌드 완료!")
if os.path.exists(exe_path):
    size_mb = os.path.getsize(exe_path) / 1024 / 1024
    print(f"  파일: {exe_path}  ({size_mb:.0f} MB)")
print()
print("  easydsd.exe 하나만 배포하시면 됩니다.")
print("  Python 없는 PC에서도 바로 실행됩니다.")
print("=" * 58)

if os.path.exists("dist"):
    try: os.startfile("dist")
    except: pass

input("\n엔터를 누르면 창이 닫힙니다...")
