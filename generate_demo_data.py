"""Генератор демонстрационных данных для transactions.xlsx.

Запуск:  python generate_demo_data.py
Скрипт заполняет Excel-книгу несколькими месяцами операций, чтобы
приложение сразу открылось с непустой таблицей и аналитикой.
"""

from __future__ import annotations

import random
from calendar import monthrange
from datetime import date

from finance_classes import DEFAULT_FILENAME, FinanceManager, Transaction

FILENAME = DEFAULT_FILENAME

# (категория, базовая сумма, сколько раз в месяц, описание)
INCOME_PLAN = (
    ("Зарплата", 95_000, 1, "Зарплата за месяц"),
    ("Инвестиции", 7_500, 1, "Дивиденды и проценты"),
    ("Подарки", 3_000, 1, "Подарок"),
)

EXPENSE_PLAN = (
    ("Продукты", 2_500, 2, "Продукты"),
    ("Продукты", 4_200, 1, "Продукты, крупная закупка"),
    ("Транспорт", 1_800, 1, "Проездной и бензин"),
    ("Жильё", 42_000, 1, "Аренда квартиры"),
    ("Развлечения", 2_600, 1, "Кино, концерты, подписки"),
)


def month_start(offset: int) -> date:
    """Первый день месяца, отстоящего на ``offset`` месяцев назад."""
    anchor = date.today().replace(day=1)
    total = anchor.year * 12 + (anchor.month - 1) - offset
    return date(total // 12, total % 12 + 1, 1)


def generate(months: int = 4, seed: int | None = 7) -> int:
    """Заполняет Excel-книгу несколькими месяцами правдоподобных операций."""
    random.seed(seed)
    manager = FinanceManager(FILENAME)
    manager.clear()
    today = date.today()

    count = 0
    for offset in range(months):
        anchor = month_start(offset)
        days_in_month = monthrange(anchor.year, anchor.month)[1]
        for name, base, times, description in INCOME_PLAN + EXPENSE_PLAN:
            for _ in range(times):
                day = random.randint(1, min(days_in_month, 28))
                if offset == 0 and day > today.day:
                    continue
                transaction_date = anchor.replace(day=day)
                amount = round(base * random.uniform(0.85, 1.2), 2)
                manager.add_transaction(
                    Transaction(
                        amount=amount,
                        category=manager.get_category(name),
                        date=transaction_date,
                        description=description,
                    )
                )
                count += 1
    return count


if __name__ == "__main__":
    total = generate()
    print(f"Создано операций: {total}. Файл: {FILENAME}")
