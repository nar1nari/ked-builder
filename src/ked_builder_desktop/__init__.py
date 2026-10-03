import threading
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, ttk

from ked_builder import build_ked
from ked_builder.program_info import (
    get_program_code,
    get_program_name,
    get_program_name_kk,
    get_program_year,
)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("KED Builder")
        self.resizable(False, False)

        self.year_var = tk.StringVar()
        self.number_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.name_kk_var = tk.StringVar()

        self.file_vars = {
            "xlsx": tk.StringVar(),
            "ru": tk.StringVar(),
            "kz": tk.StringVar(),
            "pps": tk.StringVar(),
            "template": tk.StringVar(),
        }

        pad = {"padx": 6, "pady": 3}
        frame = ttk.Frame(self, padding=10)
        frame.grid(row=0, column=0)

        row = 0

        # --- Файлы ---
        files = [
            (
                "xlsx",
                "Учебный план (xlsx) *",
                [("Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")],
            ),
            (
                "ru",
                "Документ ОП на русском (docx) *",
                [("Word", "*.docx"), ("Все файлы", "*.*")],
            ),
            (
                "kz",
                "Документ ОП на казахском (docx) *",
                [("Word", "*.docx"), ("Все файлы", "*.*")],
            ),
            ("pps", "ППС (необязательно)", [("Все файлы", "*.*")]),
            (
                "template",
                "Шаблон (необязательно)",
                [("Word", "*.docx *.dotx"), ("Все файлы", "*.*")],
            ),
        ]
        for key, label, ftypes in files:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", **pad)
            ttk.Entry(frame, textvariable=self.file_vars[key], width=55).grid(
                row=row, column=1, **pad
            )
            ttk.Button(
                frame,
                text="Выбрать...",
                command=lambda k=key, f=ftypes: self.choose_file(k, f),
            ).grid(row=row, column=2, **pad)
            row += 1

        ttk.Button(
            frame, text="Заполнить поля из учебного плана", command=self.parse_plan
        ).grid(row=row, column=1, sticky="w", **pad)
        row += 1

        # --- Параметры ---
        fields = [
            ("Год *", self.year_var),
            ("Номер программы *", self.number_var),
            ("Название программы (RU) *", self.name_var),
            ("Название программы (KK) *", self.name_kk_var),
        ]
        for label, var in fields:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", **pad)
            ttk.Entry(frame, textvariable=var, width=55).grid(row=row, column=1, **pad)
            row += 1

        self.generate_btn = ttk.Button(
            frame, text="Сгенерировать", command=self.generate
        )
        self.generate_btn.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        self.status_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.status_var, foreground="gray").grid(
            row=row, column=0, columnspan=3, sticky="w", **pad
        )

    # ---------- helpers ----------
    def choose_file(self, key: str, filetypes) -> None:
        path = filedialog.askopenfilename(filetypes=filetypes)
        if not path:
            return
        self.file_vars[key].set(path)
        if key == "xlsx":
            self.parse_plan(silent=True)

    def parse_plan(self, silent: bool = False) -> None:
        path = self.file_vars["xlsx"].get().strip()
        if not path:
            if not silent:
                messagebox.showwarning("KED Builder", "Сначала выберите учебный план.")
            return
        try:
            year = get_program_year(path)
            number = get_program_code(path)
            name = get_program_name(path)
            name_kk = get_program_name_kk(path)
        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror(
                "KED Builder", f"Не удалось прочитать учебный план: {exc}"
            )
            return

        if year is not None:
            self.year_var.set(str(year))
        if number is not None:
            self.number_var.set(str(number))
        if name is not None:
            self.name_var.set(str(name))
        if name_kk is not None:
            self.name_kk_var.set(str(name_kk))
        self.status_var.set("Поля заполнены из учебного плана.")

    # ---------- generation ----------
    def generate(self) -> None:
        xlsx = self.file_vars["xlsx"].get().strip()
        ru = self.file_vars["ru"].get().strip()
        kz = self.file_vars["kz"].get().strip()
        pps = self.file_vars["pps"].get().strip() or None
        template = self.file_vars["template"].get().strip() or None

        if not (xlsx and ru and kz):
            messagebox.showwarning(
                "KED Builder", "Укажите учебный план, документ RU и документ KZ."
            )
            return

        try:
            year = int(self.year_var.get().strip())
        except ValueError:
            messagebox.showwarning("KED Builder", "Год должен быть целым числом.")
            return

        number = self.number_var.get().strip()
        name = self.name_var.get().strip()
        name_kk = self.name_kk_var.get().strip()
        if not (number and name and name_kk):
            messagebox.showwarning("KED Builder", "Заполните все поля программы.")
            return

        output = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=f"KED_{number}_{year}.docx",
            filetypes=[("Word", "*.docx")],
        )
        if not output:
            return

        self.generate_btn.state(["disabled"])
        self.status_var.set("Генерация...")

        def worker() -> None:
            try:
                build_ked(
                    year=year,
                    program_number=number,
                    program_name=name,
                    program_name_kk=name_kk,
                    xlsx_doc=xlsx,
                    ru_doc=ru,
                    kz_doc=kz,
                    output_doc=output,
                    template_doc=template,
                    pps_doc=pps,
                )
                from pathlib import Path

                if not Path(output).exists():
                    raise RuntimeError("build_ked не создал выходной файл")
                self.after(0, lambda: self._done(output, None))
            except Exception as exc:
                traceback.print_exc()
                self.after(0, lambda e=exc: self._done(output, e))

        threading.Thread(target=worker, daemon=True).start()

    def _done(self, output: str, error: Exception | None) -> None:
        self.generate_btn.state(["!disabled"])
        if error is None:
            self.status_var.set(f"Готово: {output}")
            messagebox.showinfo("KED Builder", f"Документ сохранён:\n{output}")
        else:
            self.status_var.set("Ошибка.")
            messagebox.showerror(
                "KED Builder", f"Не удалось сгенерировать документ: {error}"
            )


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
