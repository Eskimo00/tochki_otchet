from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from openpyxl.utils.exceptions import InvalidFileException

from core import build_default_output, ensure_unique_path, process_file


def run_gui() -> None:
    """Запускает окно Tkinter для генерации отчёта."""
    root = tk.Tk()
    root.title("Точки-отчет")
    root.geometry("540x260")
    root.resizable(False, False)
    root.configure(bg="#f6f1d3")  # мягкий теплый фон

    source_var = tk.StringVar()
    dest_var = tk.StringVar()

    def select_source() -> None:
        """Выбор исходного Excel-файла."""
        path = filedialog.askopenfilename(
            title="Выберите исходный Excel",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            source_var.set(path)
            if not dest_var.get():
                dest_var.set(build_default_output(Path(path)))

    def select_dest() -> None:
        """Выбор пути сохранения итогового отчёта."""
        path = filedialog.asksaveasfilename(
            title="Куда сохранить итоговый файл",
            defaultextension=".xlsx",
            initialfile="отчет прохождения точек.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            dest_var.set(str(ensure_unique_path(Path(path))))

    def _show_error(message: str) -> None:
        messagebox.showerror("Ошибка", message)

    def generate() -> None:
        """Готовим и сохраняем итоговый отчет."""
        src = Path(source_var.get())
        dst = Path(dest_var.get()) if dest_var.get() else build_default_output(src)
        dst = ensure_unique_path(dst)
        dest_var.set(str(dst))

        if not src.exists():
            _show_error("Укажите исходный файл.")
            return
        try:
            saved_path = process_file(src, dst)
        except FileNotFoundError:
            _show_error("Исходный файл не найден.")
            return
        except PermissionError:
            _show_error("Нет доступа к файлу. Закройте его и попробуйте снова.")
            return
        except InvalidFileException as exc:
            _show_error(f"Некорректный Excel-файл: {exc}")
            return
        except Exception as exc:  # noqa: BLE001
            _show_error(
                f"Не удалось сформировать отчет ({type(exc).__name__}):\n{exc}"
            )
            return

        messagebox.showinfo("Готово", f"Отчет сохранен:\n{saved_path}")

    label_kwargs = {
        "bg": "#f6f1d3",
        "fg": "#333",
        "font": ("Segoe UI", 10, "bold"),
    }
    entry_kwargs = {"bg": "#fffaf0", "fg": "#333", "font": ("Segoe UI", 10)}
    button_style = {
        "bg": "#5cb85c",
        "fg": "white",
        "activebackground": "#4cae4c",
        "activeforeground": "white",
        "font": ("Segoe UI", 10, "bold"),
        "bd": 0,
        "highlightthickness": 0,
    }
    button_alt = {
        "bg": "#5bc0de",
        "fg": "white",
        "activebackground": "#31b0d5",
        "activeforeground": "white",
        "font": ("Segoe UI", 10, "bold"),
        "bd": 0,
        "highlightthickness": 0,
    }

    tk.Label(root, text="Исходный файл:", **label_kwargs).place(x=20, y=20)
    tk.Entry(root, textvariable=source_var, width=55, **entry_kwargs).place(
        x=20, y=45
    )
    tk.Button(
        root,
        text="Выбрать...",
        command=select_source,
        **button_alt,
    ).place(x=410, y=40)

    tk.Label(root, text="Итоговый файл:", **label_kwargs).place(x=20, y=85)
    tk.Entry(root, textvariable=dest_var, width=55, **entry_kwargs).place(
        x=20, y=110
    )
    tk.Button(
        root,
        text="Сохранить как...",
        command=select_dest,
        **button_alt,
    ).place(x=410, y=105)

    tk.Button(
        root,
        text="Сформировать",
        width=20,
        command=generate,
        **button_style,
    ).place(x=190, y=170)

    root.mainloop()
