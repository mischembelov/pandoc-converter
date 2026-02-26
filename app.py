import os
import sys
import threading
import subprocess
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD

# --- Настройки ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

OUTPUT_DIR_MD = Path("output")
OUTPUT_DIR_DOCX = Path("output_docx")
OUTPUT_DIR_MD.mkdir(exist_ok=True)
OUTPUT_DIR_DOCX.mkdir(exist_ok=True)


def get_output_dir(mode):
    return OUTPUT_DIR_MD if mode == "to_md" else OUTPUT_DIR_DOCX


def convert_file(filepath, mode):
    f = Path(filepath)
    if mode == "to_md":
        out = OUTPUT_DIR_MD / f"{f.stem}.md"
        result = subprocess.run(
            ["pandoc", str(f), "-o", str(out),
             "--wrap=none", "-t", "markdown_strict+pipe_tables+fenced_code_blocks"],
            capture_output=True, text=True
        )
        return result.returncode == 0, result.stderr.strip()

    elif mode == "to_docx":
        out = OUTPUT_DIR_DOCX / f"{f.stem}.docx"
        result = subprocess.run(
            ["pandoc", str(f), "-o", str(out), "-t", "docx"],
            capture_output=True, text=True
        )
        return result.returncode == 0, result.stderr.strip()

    elif mode == "pdf_to_docx":
        try:
            from pdf2docx import Converter
            out = OUTPUT_DIR_DOCX / f"{f.stem}.docx"
            cv = Converter(str(f))
            cv.convert(str(out))
            cv.close()
            return True, ""
        except Exception as e:
            return False, str(e)


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pandoc Converter")
        self.geometry("700x680")
        self.resizable(False, False)
        self.configure(bg="#1c1c1e")

        self.files = []
        self.mode = ctk.StringVar(value="to_md")

        self._build_ui()

    def _build_ui(self):
        # Заголовок
        ctk.CTkLabel(self, text="📄 Pandoc Converter",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(20, 4))
        ctk.CTkLabel(self, text="Пакетная конвертация документов",
                     font=ctk.CTkFont(size=13), text_color="gray").pack(pady=(0, 16))

        # Зона Drag & Drop
        self.drop_frame = ctk.CTkFrame(self, width=640, height=130,
                                       fg_color="#2c2c2e", corner_radius=12,
                                       border_width=2, border_color="#3a3a3c")
        self.drop_frame.pack(padx=30, pady=(0, 12))
        self.drop_frame.pack_propagate(False)

        self.drop_label = ctk.CTkLabel(
            self.drop_frame,
            text="⬇  Перетащите файлы сюда",
            font=ctk.CTkFont(size=15),
            text_color="#8e8e93"
        )
        self.drop_label.pack(expand=True)

        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self._on_drop)

        # Кнопка выбора файлов
        ctk.CTkButton(self, text="📂  Выбрать файлы",
                      width=200, height=36,
                      command=self._choose_files).pack(pady=(0, 16))

        # Выбор режима
        mode_frame = ctk.CTkFrame(self, fg_color="#2c2c2e", corner_radius=10)
        mode_frame.pack(padx=30, fill="x", pady=(0, 16))

        ctk.CTkLabel(mode_frame, text="Режим конвертации:",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=16, pady=(12, 6))

        modes = [
            ("DOCX  →  Markdown", "to_md"),
            ("Markdown  →  DOCX", "to_docx"),
            ("PDF  →  DOCX", "pdf_to_docx"),
        ]
        for label, value in modes:
            ctk.CTkRadioButton(mode_frame, text=label, variable=self.mode,
                               value=value, font=ctk.CTkFont(size=13),
                               command=self._update_hint).pack(anchor="w", padx=24, pady=4)

        self.hint_label = ctk.CTkLabel(mode_frame, text=self._get_hint(),
                                       font=ctk.CTkFont(size=11), text_color="#636366")
        self.hint_label.pack(anchor="w", padx=24, pady=(2, 12))

        # Кнопка конвертации
        self.convert_btn = ctk.CTkButton(
            self, text="🚀  Конвертировать",
            width=220, height=42,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._start_conversion
        )
        self.convert_btn.pack(pady=(0, 12))

        # Прогресс-бар
        self.progress = ctk.CTkProgressBar(self, width=640)
        self.progress.set(0)
        self.progress.pack(padx=30, pady=(0, 8))

        self.progress_label = ctk.CTkLabel(self, text="",
                                           font=ctk.CTkFont(size=12), text_color="gray")
        self.progress_label.pack()

        # Лог
        self.log_box = ctk.CTkTextbox(self, width=640, height=160,
                                      font=ctk.CTkFont(family="Courier", size=12),
                                      fg_color="#1c1c1e", border_width=1,
                                      border_color="#3a3a3c")
        self.log_box.pack(padx=30, pady=(10, 10))
        self.log_box.configure(state="disabled")

        # Кнопка открытия папки
        self.open_btn = ctk.CTkButton(
            self, text="📁  Открыть папку с результатом",
            width=260, height=34,
            fg_color="#3a3a3c", hover_color="#48484a",
            command=self._open_output
        )
        self.open_btn.pack(pady=(0, 20))

    def _get_hint(self):
        hints = {
            "to_md": "Поддерживаемые файлы: .docx",
            "to_docx": "Поддерживаемые файлы: .md",
            "pdf_to_docx": "Поддерживаемые файлы: .pdf  |  Только текстовые PDF",
        }
        return hints[self.mode.get()]

    def _update_hint(self):
        self.hint_label.configure(text=self._get_hint())

    def _on_drop(self, event):
        raw = event.data
        # Парсим пути (могут быть в фигурных скобках если с пробелами)
        paths = []
        if "{" in raw:
            import re
            paths = re.findall(r"\{([^}]+)\}", raw)
            rest = re.sub(r"\{[^}]+\}", "", raw).split()
            paths += rest
        else:
            paths = raw.split()

        added = 0
        for p in paths:
            p = p.strip()
            if p and Path(p).is_file() and p not in self.files:
                self.files.append(p)
                added += 1

        self._update_drop_label()

    def _choose_files(self):
        mode = self.mode.get()
        ext_map = {
            "to_md": [("Word документы", "*.docx")],
            "to_docx": [("Markdown файлы", "*.md")],
            "pdf_to_docx": [("PDF файлы", "*.pdf")],
        }
        chosen = filedialog.askopenfilenames(filetypes=ext_map[mode])
        for f in chosen:
            if f not in self.files:
                self.files.append(f)
        self._update_drop_label()

    def _update_drop_label(self):
        count = len(self.files)
        if count == 0:
            self.drop_label.configure(text="⬇  Перетащите файлы сюда")
        else:
            self.drop_label.configure(
                text=f"✅  Добавлено файлов: {count}  |  Нажмите ✕ чтобы очистить",
                text_color="#32d74b"
            )
            self.drop_frame.bind("<Button-3>", lambda e: self._clear_files())

    def _clear_files(self):
        self.files = []
        self.drop_label.configure(text="⬇  Перетащите файлы сюда", text_color="#8e8e93")

    def _log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _start_conversion(self):
        if not self.files:
            self._log("❌ Нет файлов для конвертации. Добавьте файлы.")
            return
        self.convert_btn.configure(state="disabled")
        threading.Thread(target=self._run_conversion, daemon=True).start()

    def _run_conversion(self):
        self._clear_log()
        mode = self.mode.get()
        total = len(self.files)
        success = 0
        failed = 0

        self.progress.set(0)
        self._log(f"📄 Файлов к конвертации: {total}\n")

        for i, filepath in enumerate(self.files):
            name = Path(filepath).name
            self.progress_label.configure(text=f"Обработка {i+1} из {total}: {name}")

            ok, err = convert_file(filepath, mode)

            if ok:
                self._log(f"  ✅ {name}")
                success += 1
            else:
                self._log(f"  ❌ {name}: {err or 'неизвестная ошибка'}")
                failed += 1

            self.progress.set((i + 1) / total)

        self._log(f"\n{'─'*40}")
        self._log(f"✅ Успешно: {success}  |  ❌ Ошибок: {failed}")
        self.progress_label.configure(text=f"Готово: {success} из {total}")
        self.convert_btn.configure(state="normal")

    def _open_output(self):
        mode = self.mode.get()
        out_dir = get_output_dir(mode)
        if sys.platform == "win32":
            os.startfile(out_dir.resolve())
        elif sys.platform == "darwin":
            subprocess.run(["open", str(out_dir.resolve())])
        else:
            subprocess.run(["xdg-open", str(out_dir.resolve())])


if __name__ == "__main__":
    app = App()
    app.mainloop()