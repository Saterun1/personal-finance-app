"""Графический интерфейс приложения «Учёт личных финансов» (Tkinter).

Запуск:  python finance_gui.py
"""

from __future__ import annotations

import os
import tkinter as tk
from datetime import date
from tkinter import messagebox, ttk

from finance_classes import FinanceManager, Transaction, TransactionType

APP_TITLE = "Учёт личных финансов"
DEFAULT_FILENAME = "transactions.xlsx"

BG = "#f5f6f8"
INCOME_COLOR = "#1a7f37"
EXPENSE_COLOR = "#c62828"


class FinanceApp:
    """Главное окно приложения."""

    def __init__(self, root: tk.Tk, filename: str = DEFAULT_FILENAME) -> None:
        self.root = root
        self.manager = FinanceManager(filename)
        self.row_index: dict[str, int] = {}  # iid строки -> индекс в manager.transactions
        self._sort_column: str | None = None
        self._sort_reverse: bool = False

        self.root.title(APP_TITLE)
        self.root.geometry("940x640")
        self.root.minsize(820, 560)
        self.root.configure(bg=BG)

        self._create_variables()
        self._create_style()
        self._build_menu()
        self._build_input_frame()
        self._build_info_frame()
        self._build_filter_frame()
        self._build_table()
        self._build_status_bar()

        self.refresh()

    # ------------------------------------------------------------------
    # Инициализация
    # ------------------------------------------------------------------
    def _create_variables(self) -> None:
        self.amount_var = tk.StringVar()
        self.category_var = tk.StringVar()
        self.date_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))
        self.description_var = tk.StringVar()
        self.category_hint_var = tk.StringVar(value="Тип операции: —")

        self.filter_type_var = tk.StringVar(value="Все")
        self.filter_category_var = tk.StringVar(value="Все")
        self.filter_month_var = tk.StringVar(value="Все")
        self.search_var = tk.StringVar()

        self.balance_var = tk.StringVar()
        self.income_var = tk.StringVar()
        self.expense_var = tk.StringVar()
        self.count_var = tk.StringVar()
        self.status_var = tk.StringVar()

        self.amount_var.trace_add("write", lambda *_: self._update_category_hint())
        self.category_var.trace_add("write", lambda *_: self._update_category_hint())
        self.search_var.trace_add("write", lambda *_: self.refresh_table())

    def _create_style(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Big.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("TButton", padding=(10, 5))

    def _build_menu(self) -> None:
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Сохранить в Excel", command=self.save_now)
        file_menu.add_command(label="Экспорт в новый файл...", command=self.export_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Очистить все операции", command=self.clear_all)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.destroy)
        menubar.add_cascade(label="Файл", menu=file_menu)

        category_menu = tk.Menu(menubar, tearoff=0)
        category_menu.add_command(label="Добавить категорию...", command=self.add_category_dialog)
        menubar.add_cascade(label="Категории", menu=category_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Аналитика по категориям", command=self.show_analytics)
        help_menu.add_command(label="О программе", command=self.show_about)
        menubar.add_cascade(label="Справка", menu=help_menu)

        self.root.config(menu=menubar)
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)

    # ------------------------------------------------------------------
    # Элементы интерфейса
    # ------------------------------------------------------------------
    def _build_input_frame(self) -> None:
        frame = ttk.LabelFrame(self.root, text="Добавить операцию", padding=10)
        frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        ttk.Label(frame, text="Сумма:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.amount_entry = ttk.Entry(frame, textvariable=self.amount_var, width=18)
        self.amount_entry.grid(row=0, column=1, sticky=tk.EW, padx=(8, 16), pady=3)

        ttk.Label(frame, text="Категория:").grid(row=0, column=2, sticky=tk.W, pady=3)
        self.category_combo = ttk.Combobox(
            frame, textvariable=self.category_var, state="readonly", width=22
        )
        self.category_combo.grid(row=0, column=3, sticky=tk.EW, padx=(8, 16), pady=3)
        self.category_combo.bind("<<ComboboxSelected>>", lambda _event: self._update_category_hint())

        ttk.Label(frame, textvariable=self.category_hint_var).grid(
            row=0, column=4, sticky=tk.W
        )

        ttk.Label(frame, text="Дата (ГГГГ-ММ-ДД):").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.date_entry = ttk.Entry(frame, textvariable=self.date_var, width=18)
        self.date_entry.grid(row=1, column=1, sticky=tk.EW, padx=(8, 16), pady=3)

        ttk.Label(frame, text="Описание:").grid(row=1, column=2, sticky=tk.W, pady=3)
        self.description_entry = ttk.Entry(frame, textvariable=self.description_var, width=22)
        self.description_entry.grid(row=1, column=3, sticky=tk.EW, padx=(8, 16), pady=3)
        self.description_entry.bind("<Return>", lambda _event: self.add_transaction())

        buttons = ttk.Frame(frame)
        buttons.grid(row=2, column=0, columnspan=5, sticky=tk.W, pady=(8, 0))
        ttk.Button(buttons, text="Добавить", command=self.add_transaction).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Удалить выбранную", command=self.delete_selected).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Показать аналитику", command=self.show_analytics).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Очистить форму", command=self.clear_inputs).pack(side=tk.LEFT)

        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(3, weight=1)

    def _build_info_frame(self) -> None:
        frame = ttk.LabelFrame(self.root, text="Финансовая информация", padding=10)
        frame.pack(fill=tk.X, padx=10, pady=5)

        cards = (
            ("Баланс", self.balance_var, "Big.TLabel", INCOME_COLOR),
            ("Доходы", self.income_var, "TLabel", INCOME_COLOR),
            ("Расходы", self.expense_var, "TLabel", EXPENSE_COLOR),
            ("Операций", self.count_var, "TLabel", "#333333"),
        )
        for index, (title, variable, style, color) in enumerate(cards):
            card = ttk.Frame(frame, style="Card.TFrame", padding=(16, 8))
            card.grid(row=0, column=index, sticky=tk.NSEW, padx=4)
            frame.columnconfigure(index, weight=1)
            ttk.Label(card, text=title, style="TLabel").pack(anchor=tk.W)
            label = ttk.Label(card, textvariable=variable, style=style, foreground=color)
            label.pack(anchor=tk.W)

    def _build_filter_frame(self) -> None:
        frame = ttk.LabelFrame(self.root, text="Фильтр и поиск", padding=8)
        frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(frame, text="Тип:").grid(row=0, column=0, sticky=tk.W)
        self.filter_type_combo = ttk.Combobox(
            frame,
            textvariable=self.filter_type_var,
            values=["Все", TransactionType.INCOME.value, TransactionType.EXPENSE.value],
            state="readonly",
            width=10,
        )
        self.filter_type_combo.grid(row=0, column=1, padx=(4, 14))
        self.filter_type_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_table())

        ttk.Label(frame, text="Категория:").grid(row=0, column=2, sticky=tk.W)
        self.filter_category_combo = ttk.Combobox(
            frame, textvariable=self.filter_category_var, state="readonly", width=18
        )
        self.filter_category_combo.grid(row=0, column=3, padx=(4, 14))
        self.filter_category_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_table())

        ttk.Label(frame, text="Месяц:").grid(row=0, column=4, sticky=tk.W)
        self.filter_month_combo = ttk.Combobox(
            frame,
            textvariable=self.filter_month_var,
            values=["Все"],
            state="readonly",
            width=10,
        )
        self.filter_month_combo.grid(row=0, column=5, padx=(4, 14))
        self.filter_month_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_table())

        ttk.Label(frame, text="Поиск:").grid(row=0, column=6, sticky=tk.W)
        search_entry = ttk.Entry(frame, textvariable=self.search_var, width=22)
        search_entry.grid(row=0, column=7, sticky=tk.W, padx=(4, 8))

        ttk.Button(frame, text="Сбросить", command=self.reset_filters).grid(row=0, column=8, padx=(8, 0))

    def _build_table(self) -> None:
        frame = ttk.LabelFrame(self.root, text="История операций", padding=10)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("date", "category", "type", "amount", "description")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        headings = {
            "date": ("Дата", 110),
            "category": ("Категория", 150),
            "type": ("Тип", 90),
            "amount": ("Сумма", 130),
            "description": ("Описание", 300),
        }
        for key, (title, width) in headings.items():
            self.tree.heading(key, text=title, command=lambda k=key: self.sort_by(k))
            anchor = tk.E if key == "amount" else tk.W
            self.tree.column(key, width=width, anchor=anchor, stretch=key == "description")

        self.tree.tag_configure("income", foreground=INCOME_COLOR)
        self.tree.tag_configure("expense", foreground=EXPENSE_COLOR)
        self.tree.tag_configure("odd", background="#fafafa")
        self.tree.bind("<Delete>", lambda _event: self.delete_selected())
        self.tree.bind("<Double-1>", self._edit_cell)

        yscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        xscroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        yscroll.grid(row=0, column=1, sticky=tk.NS)
        xscroll.grid(row=1, column=0, sticky=tk.EW)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _build_status_bar(self) -> None:
        bar = ttk.Frame(self.root, padding=(12, 4))
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Label(bar, textvariable=self.status_var).pack(side=tk.LEFT)
        ttk.Label(bar, text=f"Файл данных: {os.path.abspath(self.manager.filename)}").pack(
            side=tk.RIGHT
        )

    # ------------------------------------------------------------------
    # Обновление интерфейса
    # ------------------------------------------------------------------
    def _update_category_hint(self) -> None:
        category = self.manager.get_category(self.category_var.get())
        if category is None:
            self.category_hint_var.set("Тип операции: —")
        else:
            self.category_hint_var.set(f"Тип операции: {category.type.value}")

    def refresh(self) -> None:
        """Перерисовывает все данные (после загрузки файла и изменений)."""
        self.category_combo["values"] = [c.name for c in self.manager.categories]
        if not self.category_var.get() and self.manager.categories:
            self.category_var.set(self.manager.categories[0].name)
        self.filter_category_combo["values"] = ["Все"] + [
            c.name for c in self.manager.categories
        ]
        if self.filter_category_var.get() == "":
            self.filter_category_var.set("Все")
        months = sorted(
            {t.date.strftime("%Y-%m") for t in self.manager.transactions} | {date.today().strftime("%Y-%m")},
            reverse=True,
        )
        self.filter_month_combo["values"] = ["Все"] + months
        if self.filter_month_var.get() == "":
            self.filter_month_var.set("Все")
        self._update_category_hint()
        self.refresh_stats()
        self.refresh_table()

    def refresh_stats(self, transactions: list[Transaction] | None = None) -> None:
        source = self.manager.transactions if transactions is None else transactions
        balance = self.manager.get_balance(source)
        self.balance_var.set(f"{balance:,.2f} руб.".replace(",", " "))
        self.income_var.set(f"{self.manager.get_income_total(source):,.2f} руб.".replace(",", " "))
        self.expense_var.set(f"{self.manager.get_expense_total(source):,.2f} руб.".replace(",", " "))
        self.count_var.set(str(len(source)))

    def _filtered(self) -> list[Transaction]:
        type_ = None
        if self.filter_type_var.get() in (t.value for t in TransactionType):
            type_ = TransactionType.from_text(self.filter_type_var.get())
        category = self.filter_category_var.get()
        month = self.filter_month_var.get()
        return self.manager.get_transactions(
            type_=type_,
            category_name=None if category == "Все" else category,
            month=None if month == "Все" else month,
            search=self.search_var.get() or None,
        )

    def refresh_table(self) -> None:
        transactions = self._filtered()
        if self._sort_column:
            transactions.sort(key=lambda t: getattr(t, self._sort_column), reverse=self._sort_reverse)

        self.tree.delete(*self.tree.get_children())
        self.row_index.clear()
        positions = {id(t): i for i, t in enumerate(self.manager.transactions)}
        for number, transaction in enumerate(transactions):
            iid = f"row{number}"
            self.row_index[iid] = positions[id(transaction)]
            tag = "income" if transaction.is_income else "expense"
            if number % 2:
                tag = (tag, "odd")
            self.tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    transaction.date_str,
                    transaction.category.name,
                    transaction.category.type.value,
                    f"{transaction.amount:,.2f}".replace(",", " "),
                    transaction.description,
                ),
                tags=tag,
            )
        self.refresh_stats(transactions)
        self.status_var.set(
            f"Показано операций: {len(transactions)} из {len(self.manager.transactions)}"
        )

    # ------------------------------------------------------------------
    # Действия пользователя
    # ------------------------------------------------------------------
    def add_transaction(self) -> None:
        amount_text = self.amount_var.get().strip().replace(" ", "").replace(",", ".")
        if not amount_text or not self.category_var.get() or not self.date_var.get().strip():
            messagebox.showwarning("Ошибка", "Заполните обязательные поля: сумма, категория, дата.")
            return
        try:
            transaction = Transaction(
                amount=amount_text,
                category=self.manager.get_category(self.category_var.get()),
                date=self.date_var.get().strip(),
                description=self.description_var.get().strip(),
            )
        except ValueError as exc:
            messagebox.showerror("Ошибка", str(exc))
            return

        self.manager.add_transaction(transaction)
        self.clear_inputs()
        self.refresh()
        self.status_var.set("Операция добавлена и сохранена в Excel-файл")

    def delete_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Удаление", "Выберите операцию в таблице.")
            return
        index = self.row_index.get(selection[0])
        if index is None or not self.manager.delete_transaction(index):
            messagebox.showerror("Ошибка", "Не удалось удалить операцию.")
            return
        self.refresh()
        self.status_var.set("Операция удалена")

    def sort_by(self, column: str) -> None:
        """Сортировка таблицы по выбранной колонке (повторный клик меняет порядок)."""
        if self._sort_column == column:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_column, self._sort_reverse = column, False
        self.refresh_table()

    def clear_inputs(self) -> None:
        self.amount_var.set("")
        self.description_var.set("")
        self.date_var.set(date.today().strftime("%Y-%m-%d"))
        self.category_var.set(self.manager.categories[0].name if self.manager.categories else "")
        self.amount_entry.focus_set()

    def reset_filters(self) -> None:
        self.filter_type_var.set("Все")
        self.filter_category_var.set("Все")
        self.filter_month_var.set("Все")
        self.search_var.set("")
        self._sort_column, self._sort_reverse = None, False
        self.refresh_table()

    def save_now(self) -> None:
        self.manager.save_to_file()
        self.status_var.set(f"Сохранено в {self.manager.filename}")

    def export_dialog(self) -> None:
        from tkinter import filedialog

        path = filedialog.asksaveasfilename(
            title="Сохранить копию данных",
            defaultextension=".xlsx",
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        self.manager.save_to_file(path)
        self.status_var.set(f"Экспортировано в {path}")

    def clear_all(self) -> None:
        if not self.manager.transactions:
            messagebox.showinfo("Очистка", "Список операций уже пуст.")
            return
        if not messagebox.askyesno("Подтверждение", "Удалить все операции? Действие необратимо."):
            return
        self.manager.clear()
        self.refresh()
        self.status_var.set("Все операции удалены")

    def add_category_dialog(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("Новая категория")
        window.resizable(False, False)
        window.transient(self.root)

        ttk.Label(window, text="Название:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=(10, 4))
        name_var = tk.StringVar()
        name_entry = ttk.Entry(window, textvariable=name_var, width=24)
        name_entry.grid(row=0, column=1, padx=10, pady=(10, 4))

        ttk.Label(window, text="Тип:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=4)
        type_var = tk.StringVar(value=TransactionType.EXPENSE.value)
        ttk.Combobox(
            window,
            textvariable=type_var,
            values=[t.value for t in TransactionType],
            state="readonly",
            width=21,
        ).grid(row=1, column=1, padx=10, pady=4)

        def save() -> None:
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Ошибка", "Введите название категории.", parent=window)
                return
            category = self.manager.add_category(name, type_var.get())
            self.category_var.set(category.name)
            self.refresh()
            messagebox.showinfo("Готово", f"Категория «{category.name}» добавлена.", parent=window)
            window.destroy()

        buttons = ttk.Frame(window)
        buttons.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(buttons, text="Сохранить", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons, text="Отмена", command=window.destroy).pack(side=tk.LEFT, padx=5)

        window.bind("<Return>", lambda _event: save())
        window.bind("<Escape>", lambda _event: window.destroy())
        name_entry.focus_set()

    # ------------------------------------------------------------------
    # Аналитика
    # ------------------------------------------------------------------
    def show_analytics(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("Аналитика по категориям")
        window.geometry("640x520")
        window.transient(self.root)

        summary = self.manager.get_category_summary()
        balance = self.manager.get_balance()

        header = ttk.Frame(window, padding=10)
        header.pack(fill=tk.X)
        ttk.Label(
            header,
            text=f"Баланс: {balance:,.2f} руб.   •   Доходы: "
            f"{self.manager.get_income_total():,.2f} руб.   •   Расходы: "
            f"{self.manager.get_expense_total():,.2f} руб.".replace(",", " "),
            style="Big.TLabel",
        ).pack(anchor=tk.W)
        if not summary:
            ttk.Label(header, text="Нет данных для анализа.").pack(anchor=tk.W, pady=10)
            ttk.Button(header, text="Закрыть", command=window.destroy).pack(pady=10)
            return

        ttk.Label(window, text="Суммы по категориям", style="Title.TLabel").pack(anchor=tk.W, padx=10)
        self._build_summary_table(window, summary)

        ttk.Label(window, text="Динамика по месяцам", style="Title.TLabel").pack(
            anchor=tk.W, padx=10, pady=(12, 0)
        )
        self._build_months_table(window, self.manager.get_monthly_summary())

        ttk.Button(window, text="Закрыть", command=window.destroy).pack(pady=10)

    def _build_summary_table(self, parent: tk.Widget, summary: dict[str, float]) -> None:
        frame = ttk.Frame(parent, padding=(10, 4))
        frame.pack(fill=tk.BOTH, expand=False)

        columns = ("category", "type", "amount", "share")
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=min(len(summary) + 1, 8))
        for key, title, width, anchor in (
            ("category", "Категория", 200, tk.W),
            ("type", "Тип", 90, tk.W),
            ("amount", "Сумма", 130, tk.E),
            ("share", "Доля расходов", 110, tk.E),
        ):
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor=anchor)
        tree.tag_configure("income", foreground=INCOME_COLOR)
        tree.tag_configure("expense", foreground=EXPENSE_COLOR)

        total_expense = sum(abs(value) for value in summary.values() if value < 0) or 1
        for index, (name, value) in enumerate(summary.items()):
            is_income = self.manager.get_category(name) and self.manager.get_category(name).is_income()
            tree.insert(
                "",
                tk.END,
                values=(
                    name,
                    "Доход" if is_income else "Расход",
                    f"{value:,.2f}".replace(",", " "),
                    f"{abs(value) / total_expense * 100:.1f} %" if value < 0 else "—",
                ),
                tags="income" if is_income else "expense",
                iid=f"c{index}",
            )

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_months_table(self, parent: tk.Widget, months: dict[str, dict[str, float]]) -> None:
        frame = ttk.Frame(parent, padding=(10, 4))
        frame.pack(fill=tk.BOTH, expand=True)

        columns = ("month", "income", "expense", "balance")
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        for key, title, width in (
            ("month", "Месяц", 120),
            ("income", "Доходы", 140),
            ("expense", "Расходы", 140),
            ("balance", "Баланс", 140),
        ):
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor=tk.E if key != "month" else tk.W)
        tree.tag_configure("income", foreground=INCOME_COLOR)
        tree.tag_configure("expense", foreground=EXPENSE_COLOR)

        for index, (month, item) in enumerate(months.items()):
            tree.insert(
                "",
                tk.END,
                iid=f"m{index}",
                values=(
                    month,
                    f"{item['income']:,.2f}".replace(",", " "),
                    f"{item['expense']:,.2f}".replace(",", " "),
                    f"{item['balance']:,.2f}".replace(",", " "),
                ),
                tags="income" if item["balance"] >= 0 else "expense",
            )
        tree.pack(fill=tk.BOTH, expand=True)

    def show_about(self) -> None:
        messagebox.showinfo(
            "О программе",
            f"{APP_TITLE}\n\n"
            "Desktop-приложение на Python + Tkinter.\n"
            "Данные хранятся в Excel-книге (transactions.xlsx) рядом с программой.\n"
            "Клавиша Delete удаляет выбранную операцию.",
        )

    # ------------------------------------------------------------------
    def _edit_cell(self, event: tk.Event) -> None:  # pragma: no cover - удобство UI
        """Двойной клик по строке переносит сумму в форму для быстрой правки."""
        row = self.tree.identify_row(event.y)
        if not row:
            return
        index = self.row_index.get(row)
        if index is None:
            return
        transaction = self.manager.transactions[index]
        self.amount_var.set(f"{transaction.amount:.2f}")
        self.category_var.set(transaction.category.name)
        self.date_var.set(transaction.date_str)
        self.description_var.set(transaction.description)
        self.amount_entry.focus_set()
        self.status_var.set(
            "Данные операции перенесены в форму. Нажмите «Добавить», чтобы создать новую запись."
        )


def main() -> None:
    root = tk.Tk()
    FinanceApp(root, os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFAULT_FILENAME))
    root.mainloop()


if __name__ == "__main__":
    main()
