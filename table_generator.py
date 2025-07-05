data = [
    {'a': 1, 'b': 2},
    {'a': 3, 'c': 4},
    {'b': 5, 'c': 6, 'd': 7}
]

# Собираем все возможные ключи
all_keys = set()
for d in data:
    all_keys.update(d.keys())

# Сортируем ключи (по желанию, можно убрать sorted)
all_keys = sorted(all_keys)

# Заголовки: первый — "Key", остальные — индексы словарей
headers = ['Key'] + [str(i) for i in range(len(data))]

# Формируем строки
rows = []
for key in all_keys:
    row = [key]
    for d in data:
        row.append(d.get(key, 'n/a'))
    rows.append(row)

# Печать в виде таблицы (без библиотек)
from tabulate import tabulate
print(tabulate(rows, headers=headers, tablefmt='grid'))
