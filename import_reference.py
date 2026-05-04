import os
import re
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'timberlogic.settings')  # Замените на имя вашего проекта
django.setup()

from lumber_track.models import ProductType, ProductName, UnitDimension
import openpyxl

# Путь к файлу
file_path = 'names.xlsx'

# Проверка существования файла
if not os.path.exists(file_path):
    print(f'Ошибка: файл {file_path} не найден!')
    print(f'Текущая папка: {os.getcwd()}')
    exit(1)

# Получаем или создаем тип "Штучный"
product_type, _ = ProductType.objects.get_or_create(name="Штучный")
print(f'Тип продукции: {product_type.name}')

# Загрузка Excel файла
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Статистика
stats = {
    'total_rows': 0,
    'names_created': 0,
    'names_existed': 0,
    'dimensions_created': 0,
    'dimensions_existed': 0,
    'errors': 0
}

print(f'\nНачат импорт справочников из файла {file_path}')
print(f'Лист: {sheet.title}\n')
print('=' * 60)

# Регулярное выражение для парсинга размеров
size_pattern = re.compile(r'(\d+)\s*[\*xх]\s*(\d+)\s*[\*xх]\s*(\d+)', re.IGNORECASE)

# Для отслеживания дубликатов в рамках импорта
processed_names = set()
processed_dimensions = set()

# Чтение данных (предполагаем, что первая строка - заголовки)
for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
    if not row or len(row) < 3:
        continue

    product_name_str = row[0]  # колонка A - наименование
    unit_str = row[1]  # колонка B - единица измерения (не используем)
    size_str = row[2]  # колонка C - размеры через *

    # Пропускаем пустые строки
    if not product_name_str or not size_str:
        continue

    stats['total_rows'] += 1

    # Очистка
    product_name_str = str(product_name_str).strip()
    size_str = str(size_str).strip()

    # 1. Парсинг размеров
    match = size_pattern.search(size_str)
    if not match:
        print(f'❌ Строка {row_idx}: не удалось распарсить размеры "{size_str}"')
        stats['errors'] += 1
        continue

    length = int(match.group(1))
    width = int(match.group(2))
    height = int(match.group(3))

    # 2. Добавляем размер в справочник UnitDimension
    dim_key = (length, width, height)
    if dim_key not in processed_dimensions:
        dim_obj, created = UnitDimension.objects.get_or_create(
            length=length,
            width=width,
            height=height
        )
        if created:
            stats['dimensions_created'] += 1
            print(f'✅ Строка {row_idx}: добавлен новый размер {length}x{width}x{height}')
        else:
            stats['dimensions_existed'] += 1
            print(f'   Строка {row_idx}: размер {length}x{width}x{height} уже есть в справочнике')
        processed_dimensions.add(dim_key)

    # 3. Добавляем наименование в справочник ProductName
    if product_name_str not in processed_names:
        name_obj, created = ProductName.objects.get_or_create(
            name=product_name_str,
            defaults={'product_type': product_type}
        )
        if created:
            stats['names_created'] += 1
            print(f'📦 Строка {row_idx}: добавлено новое наименование "{product_name_str}"')
        else:
            stats['names_existed'] += 1
            # Проверяем, что тип правильный
            if name_obj.product_type != product_type:
                print(
                    f'⚠️ Строка {row_idx}: "{product_name_str}" уже существует с типом "{name_obj.product_type}", а должен быть "{product_type}"')
        processed_names.add(product_name_str)
    else:
        print(f'   Строка {row_idx}: наименование "{product_name_str}" уже обработано')

# Итоги
print('\n' + '=' * 60)
print('ИМПОРТ СПРАВОЧНИКОВ ЗАВЕРШЕН')
print(f'📊 Обработано строк: {stats["total_rows"]}')
print(f'📦 Наименований создано: {stats["names_created"]}, уже существовало: {stats["names_existed"]}')
print(f'📏 Размеров создано: {stats["dimensions_created"]}, уже существовало: {stats["dimensions_existed"]}')
print(f'❌ Ошибок: {stats["errors"]}')
print('=' * 60)

# Показать что добавилось
print('\n📋 Новые наименования в справочнике:')
for name in ProductName.objects.filter(product_type=product_type).order_by('-id')[:10]:
    print(f'  - {name.name}')

print('\n📐 Новые размеры в справочнике (первые 10):')
for dim in UnitDimension.objects.all().order_by('-id')[:10]:
    print(f'  - {dim.length}x{dim.width}x{dim.height} мм')