"""Модели данных и бизнес-логика приложения «Учёт личных финансов».

Данные хранятся в Excel-книге transactions.xlsx (лист «Операции»).
Модуль не зависит от графического интерфейса, поэтому его можно
использовать и тестировать отдельно от Tkinter.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date, datetime
from enum import Enum
from typing import Iterable

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError as exc:  # pragma: no cover - только при отсутствии пакета
    raise SystemExit(
        "Не установлена библиотека openpyxl (работа с Excel-файлами).\n"
        "Установите её командой:  pip install openpyxl"
    ) from exc

__all__ = [
    "TransactionType",
    "Category",
    "Transaction",
    "FinanceManager",
    "DEFAULT_CATEGORIES",
    "HEADER",
    "DEFAULT_FILENAME",
    "SHEET_NAME",
    "ANALYTICS_SHEET",
    "parse_date",
]


# --------------------------------------------------------------------------
# Перечисления и справочники
# --------------------------------------------------------------------------
class TransactionType(Enum):
    """Тип финансовой операции."""

    INCOME = "Доход"
    EXPENSE = "Расход"

    @classmethod
    def from_text(cls, text: str) -> "TransactionType":
        """Создаёт тип операции из строки (регистр не важен)."""
        normalized = (text or "").strip().lower()
        for item in cls:
            if normalized in (item.value.lower(), item.name.lower()):
                return item
        raise ValueError(f"Неизвестный тип операции: {text!r}")


#: Категории, которые создаются при первом запуске приложения.
DEFAULT_CATEGORIES: tuple[tuple[str, TransactionType], ...] = (
    ("Зарплата", TransactionType.INCOME),
    ("Инвестиции", TransactionType.INCOME),
    ("Подарки", TransactionType.INCOME),
    ("Продукты", TransactionType.EXPENSE),
    ("Транспорт", TransactionType.EXPENSE),
    ("Жильё", TransactionType.EXPENSE),
    ("Развлечения", TransactionType.EXPENSE),
)

#: Имя файла данных по умолчанию.
DEFAULT_FILENAME = "transactions.xlsx"

#: Имя листа с операциями.
SHEET_NAME = "Операции"

#: Имя листа с аналитикой (создаётся при сохранении).
ANALYTICS_SHEET = "Аналитика"

#: Заголовок таблицы операций.
HEADER = ("Amount", "Category", "Type", "Date", "Description")

DATE_FORMAT = "%Y-%m-%d"

#: Оформление: жирный заголовок на тёмном фоне.
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_FILL = PatternFill("solid", fgColor="2F5597")
_BOLD = Font(bold=True)
_COLUMN_WIDTHS = (12, 20, 10, 12, 42)


def parse_date(value: str | date | datetime) -> date:
    """Приводит строку вида ``ГГГГ-ММ-ДД`` к объекту :class:`datetime.date`."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return datetime.strptime(str(value).strip(), DATE_FORMAT).date()
    except ValueError as exc:  # pragma: no cover - сообщение для пользователя
        raise ValueError(
            f"Некорректная дата: {value!r}. Ожидается формат ГГГГ-ММ-ДД"
        ) from exc


# --------------------------------------------------------------------------
# Модели
# --------------------------------------------------------------------------
@dataclass
class Category:
    """Категория доходов или расходов."""

    name: str
    type: TransactionType

    def __post_init__(self) -> None:
        self.name = str(self.name).strip()
        if not self.name:
            raise ValueError("Название категории не может быть пустым")
        if not isinstance(self.type, TransactionType):
            self.type = TransactionType.from_text(str(self.type))

    def __str__(self) -> str:  # pragma: no cover - только для отладки
        return f"{self.name} ({self.type.value})"


@dataclass
class Transaction:
    """Финансовая операция."""

    amount: float
    category: Category
    date: date | str | None = None
    description: str = ""

    def __post_init__(self) -> None:
        try:
            amount = float(self.amount)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Некорректная сумма: {self.amount!r}") from exc
        if amount <= 0:
            raise ValueError("Сумма операции должна быть больше нуля")

        if not isinstance(self.category, Category):
            raise ValueError("Категория должна быть объектом Category")

        self.amount = round(amount, 2)
        self.date = parse_date(self.date if self.date is not None else date.today())
        self.description = str(self.description or "").strip()

    @property
    def date_str(self) -> str:
        """Дата в формате, используемом при сохранении в CSV."""
        return self.date.strftime(DATE_FORMAT)

    @property
    def is_income(self) -> bool:
        return self.category.type is TransactionType.INCOME

    @property
    def signed_amount(self) -> float:
        """Сумма со знаком: ``+`` для дохода, ``-`` для расхода."""
        return self.amount if self.is_income else -self.amount

    def to_row(self) -> list[object]:
        """Строка для листа Excel («Операции»)."""
        return [
            self.amount,
            self.category.name,
            self.category.type.value,
            self.date_str,
            self.description,
        ]


# --------------------------------------------------------------------------
# Менеджер
# --------------------------------------------------------------------------
class FinanceManager:
    """Управляет операциями, категориями и хранением данных в CSV."""

    def __init__(self, filename: str = DEFAULT_FILENAME) -> None:
        self.filename = filename
        self.transactions: list[Transaction] = []
        self.categories: list[Category] = [
            Category(name, type_) for name, type_ in DEFAULT_CATEGORIES
        ]
        self.load_from_file()

    # -- работа с категориями ---------------------------------------------
    def get_category(self, name: str) -> Category | None:
        """Возвращает категорию по названию либо ``None``."""
        target = (name or "").strip().lower()
        for category in self.categories:
            if category.name.lower() == target:
                return category
        return None

    def add_category(self, name: str, type_: TransactionType | str) -> Category:
        """Добавляет новую категорию, если её ещё нет."""
        existing = self.get_category(name)
        if existing:
            return existing
        category = Category(name, type_)
        self.categories.append(category)
        return category

    def ensure_category(self, name: str, type_: TransactionType | str) -> Category:
        """Возвращает категорию, создавая её при необходимости."""
        return self.add_category(name, type_)

    # -- работа с операциями ----------------------------------------------
    def add_transaction(
        self,
        transaction: Transaction | None = None,
        *,
        amount: float | None = None,
        category: Category | str | None = None,
        date_: str | date | None = None,
        description: str = "",
    ) -> Transaction:
        """Добавляет операцию и сохраняет файл.

        Можно передать готовый объект :class:`Transaction` либо отдельные
        параметры. В любом случае данные сразу записываются в CSV.
        """
        if transaction is None:
            if amount is None or category is None:
                raise ValueError("Нужно указать сумму и категорию")
            if isinstance(category, str):
                found = self.get_category(category)
                if found is None:
                    raise ValueError(f"Категория не найдена: {category!r}")
                category = found
            transaction = Transaction(
                amount=amount,
                category=category,
                date=date_ or date.today(),
                description=description,
            )
        self.transactions.append(transaction)
        self.save_to_file()
        return transaction

    def delete_transaction(self, index: int) -> bool:
        """Удаляет операцию по индексу. Возвращает ``True`` при успехе."""
        if 0 <= index < len(self.transactions):
            del self.transactions[index]
            self.save_to_file()
            return True
        return False

    def clear(self) -> None:
        """Удаляет все операции."""
        self.transactions.clear()
        self.save_to_file()

    def get_transactions(
        self,
        type_: TransactionType | None = None,
        category_name: str | None = None,
        month: str | None = None,
        search: str | None = None,
    ) -> list[Transaction]:
        """Возвращает отфильтрованный список операций (по дате — новые сверху)."""
        result: list[Transaction] = list(self.transactions)

        if type_ is not None:
            result = [t for t in result if t.category.type is type_]
        if category_name:
            result = [t for t in result if t.category.name == category_name]
        if month:
            result = [t for t in result if t.date.strftime("%Y-%m") == month]
        if search:
            needle = search.strip().lower()
            result = [
                t
                for t in result
                if needle in t.description.lower()
                or needle in t.category.name.lower()
            ]

        return sorted(result, key=lambda t: (t.date, t.category.name), reverse=True)

    # -- аналитика ---------------------------------------------------------
    def get_income_total(self, transactions: Iterable[Transaction] | None = None) -> float:
        source = self.transactions if transactions is None else transactions
        return round(sum(t.amount for t in source if t.is_income), 2)

    def get_expense_total(self, transactions: Iterable[Transaction] | None = None) -> float:
        source = self.transactions if transactions is None else transactions
        return round(sum(t.amount for t in source if not t.is_income), 2)

    def get_balance(self, transactions: Iterable[Transaction] | None = None) -> float:
        """Текущий баланс: доходы минус расходы."""
        source = self.transactions if transactions is None else transactions
        return round(sum(t.signed_amount for t in source), 2)

    def get_category_summary(
        self, transactions: Iterable[Transaction] | None = None
    ) -> dict[str, float]:
        """Сумма операций по каждой категории (расходы — со знаком минус)."""
        source = self.transactions if transactions is None else transactions
        summary: dict[str, float] = {}
        for transaction in source:
            name = transaction.category.name
            summary[name] = round(summary.get(name, 0.0) + transaction.signed_amount, 2)
        return dict(sorted(summary.items(), key=lambda item: -abs(item[1])))

    def get_monthly_summary(self) -> dict[str, dict[str, float]]:
        """Разбивка доходов, расходов и баланса по месяцам (``ГГГГ-ММ``)."""
        summary: dict[str, dict[str, float]] = {}
        for transaction in self.transactions:
            key = transaction.date.strftime("%Y-%m")
            item = summary.setdefault(key, {"income": 0.0, "expense": 0.0, "balance": 0.0})
            if transaction.is_income:
                item["income"] += transaction.amount
            else:
                item["expense"] += transaction.amount
        for item in summary.values():
            item["income"] = round(item["income"], 2)
            item["expense"] = round(item["expense"], 2)
            item["balance"] = round(item["income"] - item["expense"], 2)
        return dict(sorted(summary.items(), reverse=True))

    # -- хранение данных (Excel) ------------------------------------------
    def save_to_file(self, filename: str | None = None) -> None:
        """Сохраняет все операции в Excel-книгу (``.xlsx``).

        Первый лист «Операции» содержит сами данные, второй лист
        «Аналитика» — сводку по категориям и по месяцам.
        """
        path = filename or self.filename
        directory = os.path.dirname(os.path.abspath(path))
        if directory and not os.path.isdir(directory):
            os.makedirs(directory, exist_ok=True)

        workbook = openpyxl.Workbook()
        self._fill_transactions_sheet(workbook.active)
        self._fill_analytics_sheet(workbook.create_sheet(ANALYTICS_SHEET))
        workbook.save(path)

    def _fill_transactions_sheet(self, sheet: "openpyxl.worksheet.worksheet.Worksheet") -> None:
        """Заполняет лист «Операции» с оформлением заголовка."""
        sheet.title = SHEET_NAME
        sheet.append(list(HEADER))
        for cell in sheet[1]:
            cell.font = _HEADER_FONT
            cell.fill = _HEADER_FILL
            cell.alignment = Alignment(horizontal="center")

        for transaction in self.transactions:
            sheet.append(transaction.to_row())

        for index, width in enumerate(_COLUMN_WIDTHS, start=1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1):
            row[0].number_format = "0.00"

        sheet.freeze_panes = "A2"
        if self.transactions:
            sheet.auto_filter.ref = f"A1:E{len(self.transactions) + 1}"

    def _fill_analytics_sheet(
        self, sheet: "openpyxl.worksheet.worksheet.Worksheet"
    ) -> None:
        """Заполняет лист «Аналитика»: суммы по категориям и по месяцам."""
        sheet.append(["Аналитика по категориям"])
        sheet["A1"].font = Font(bold=True, size=13)

        sheet.append(["Категория", "Тип", "Доходы", "Расходы", "Итог"])
        for cell in sheet[2]:
            cell.font = _HEADER_FONT
            cell.fill = _HEADER_FILL
            cell.alignment = Alignment(horizontal="center")

        summary = self.get_category_summary()
        for name, total in summary.items():
            category = self.get_category(name)
            income = self.get_income_total(
                t for t in self.transactions if t.category.name == name
            )
            expense = self.get_expense_total(
                t for t in self.transactions if t.category.name == name
            )
            sheet.append(
                [
                    name,
                    category.type.value if category else "",
                    income or None,
                    expense or None,
                    total,
                ]
            )

        sheet.append([])
        sheet.append(["Итого", "", self.get_income_total(), self.get_expense_total(), self.get_balance()])
        for cell in sheet[sheet.max_row]:
            cell.font = _BOLD

        sheet.append([])
        sheet.append(["Динамика по месяцам"])
        sheet[sheet.max_row][0].font = Font(bold=True, size=13)
        sheet.append(["Месяц", "Доходы", "Расходы", "Баланс"])
        for cell in sheet[sheet.max_row]:
            cell.font = _HEADER_FONT
            cell.fill = _HEADER_FILL
            cell.alignment = Alignment(horizontal="center")
        for month, item in self.get_monthly_summary().items():
            sheet.append([month, item["income"], item["expense"], item["balance"]])

        for column, width in zip("ABCDE", (20, 12, 14, 14, 14)):
            sheet.column_dimensions[column].width = width

    def load_from_file(self, filename: str | None = None) -> int:
        """Загружает операции из Excel-книги. Возвращает число записей."""
        path = filename or self.filename
        if not os.path.exists(path):
            self.transactions = []
            return 0

        try:
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except Exception:
            # Файл повреждён или не является книгой Excel — начинаем с нуля.
            self.transactions = []
            return 0

        try:
            sheet = (
                workbook[SHEET_NAME]
                if SHEET_NAME in workbook.sheetnames
                else workbook.worksheets[0]
            )
            rows = sheet.iter_rows(values_only=True)
        except (KeyError, IndexError):
            workbook.close()
            self.transactions = []
            return 0

        loaded: list[Transaction] = []
        for index, row in enumerate(rows):
            if not row or all(cell is None or str(cell).strip() == "" for cell in row):
                continue
            values = list(row) + [None] * (len(HEADER) - len(row))
            if index == 0 and str(values[0] or "").strip().lower() == HEADER[0].lower():
                continue  # строка заголовка
            if values[1] is None:
                continue
            try:
                category = self.ensure_category(
                    str(values[1]).strip(),
                    values[2] if values[2] is not None else TransactionType.EXPENSE.value,
                )
                transaction = Transaction(
                    amount=values[0],
                    category=category,
                    date=values[3] if values[3] is not None else date.today(),
                    description=values[4] or "",
                )
            except (ValueError, TypeError, AttributeError):
                # Некорректная строка не должна ломать загрузку файла.
                continue
            loaded.append(transaction)

        workbook.close()
        self.transactions = loaded
        return len(loaded)
