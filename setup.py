import subprocess
import sys
from pathlib import Path

# --- 1. Создаём папки ---
print("📁 Создаём рабочие папки...")
for folder in ["input", "input_md", "input_pdf", "output", "output_docx"]:
    Path(folder).mkdir(exist_ok=True)
    print(f"   ✅ {folder}/")

# --- 2. Устанавливаем Python-зависимости ---
print("\n📦 Устанавливаем Python-зависимости...")
result = subprocess.run(
    [sys.executable, "-m", "pip", "install", "-r", "requirements.txt"],
    capture_output=False
)
if result.returncode != 0:
    print("❌ Ошибка при установке зависимостей")
    sys.exit(1)
print("✅ Зависимости установлены")

# --- 3. Устанавливаем Pandoc через pypandoc ---
print("\n⚙️  Проверяем и устанавливаем Pandoc...")
try:
    import pypandoc
    try:
        version = pypandoc.get_pandoc_version()
        print(f"✅ Pandoc уже установлен: v{version}")
    except OSError:
        print("   Pandoc не найден — устанавливаем...")
        pypandoc.download_pandoc()
        print("✅ Pandoc установлен")
except Exception as e:
    print(f"❌ Не удалось установить Pandoc: {e}")
    print("   Установи вручную: https://pandoc.org/installing.html")
    sys.exit(1)

print("\n🚀 Готово! Проект настроен.")
print("\nДальше:")
print("  python convert.py to_md          # DOCX → MD")
print("  python convert.py to_docx        # MD → DOCX")
print("  python convert.py pdf_to_docx    # PDF → DOCX")
