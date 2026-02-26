import subprocess
import sys
from pathlib import Path

MODE = sys.argv[1] if len(sys.argv) > 1 else None

# Режим 1: DOCX → MD
if MODE == "to_md":
    INPUT_DIR = Path("input")
    OUTPUT_DIR = Path("output")
    files = list(INPUT_DIR.glob("*.docx"))

    if not files:
        print("❌ Нет файлов .docx в папке input/")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"📄 Найдено файлов: {len(files)}\n")
    success, failed = [], []

    for docx in files:
        out = OUTPUT_DIR / f"{docx.stem}.md"
        result = subprocess.run(
            ["pandoc", str(docx), "-o", str(out),
             "--wrap=none", "-t", "markdown_strict+pipe_tables+fenced_code_blocks"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            success.append(docx.name)
            print(f"  ✅ {docx.name}")
        else:
            failed.append((docx.name, result.stderr.strip()))
            print(f"  ❌ {docx.name}: {result.stderr.strip()}")

    print(f"\n✅ Успешно: {len(success)}  |  ❌ Ошибок: {len(failed)}")

# Режим 2: MD → DOCX
elif MODE == "to_docx":
    INPUT_DIR = Path("input_md")
    OUTPUT_DIR = Path("output_docx")
    files = list(INPUT_DIR.glob("*.md"))

    if not files:
        print("❌ Нет файлов .md в папке input_md/")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"📄 Найдено файлов: {len(files)}\n")
    success, failed = [], []

    for md in files:
        out = OUTPUT_DIR / f"{md.stem}.docx"
        result = subprocess.run(
            ["pandoc", str(md), "-o", str(out), "-t", "docx"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            success.append(md.name)
            print(f"  ✅ {md.name}")
        else:
            failed.append((md.name, result.stderr.strip()))
            print(f"  ❌ {md.name}: {result.stderr.strip()}")

    print(f"\n✅ Успешно: {len(success)}  |  ❌ Ошибок: {len(failed)}")

# Режим 3: PDF → DOCX (новое!)
elif MODE == "pdf_to_docx":
    try:
        from pdf2docx import Converter
    except ImportError:
        print("❌ Библиотека pdf2docx не установлена!")
        print("Установи: pip install pdf2docx")
        sys.exit(1)

    INPUT_DIR = Path("input_pdf")
    OUTPUT_DIR = Path("output_docx")
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    files = list(INPUT_DIR.glob("*.pdf"))

    if not files:
        print("❌ Нет файлов .pdf в папке input_pdf/")
        sys.exit(1)

    print(f"📄 Найдено файлов: {len(files)}\n")
    success, failed = [], []

    for pdf in files:
        out = OUTPUT_DIR / f"{pdf.stem}.docx"
        try:
            cv = Converter(str(pdf))
            cv.convert(str(out))
            cv.close()
            success.append(pdf.name)
            print(f"  ✅ {pdf.name}")
        except Exception as e:
            failed.append((pdf.name, str(e)))
            print(f"  ❌ {pdf.name}: {e}")

    print(f"\n✅ Успешно: {len(success)}  |  ❌ Ошибок: {len(failed)}")

else:
    print("Использование:")
    print("  python convert.py to_md          # DOCX → MD")
    print("  python convert.py to_docx        # MD → DOCX")
    print("  python convert.py pdf_to_docx    # PDF → DOCX")
    sys.exit(1)
