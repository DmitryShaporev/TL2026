import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'timberlogic.settings')  # ЗАМЕНИТЕ your_project на имя вашего проекта
django.setup()

from lumber_track.models import Dimension
import openpyxl

file_path = 'sizes.xlsx'

# Проверка существования файла
if not os.path.exists(file_path):
    print(f'Ошибка: файл {file_path} не найден!')
    print(f'Текущая папка: {os.getcwd()}')
    exit(1)

# Загрузка файла
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Статистика
created = 0
skipped = 0
errors = 0

print(f'Начат импорт из файла {file_path}')
print(f'Лист: {sheet.title}\n')

# Чтение данных
for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
    if not row or len(row) < 3:
        continue

    thickness = row[0]  # колонка A
    width = row[1]  # колонка B
    length = row[2]  # колонка C

    # Пропускаем пустые значения
    if thickness is None or width is None or length is None:
        continue

    # Пробуем пропустить строку с заголовками (если в первой строке текст)
    if row_idx == 1:
        if isinstance(thickness, str) or isinstance(width, str) or isinstance(length, str):
            print(f'Пропущена строка {row_idx} (возможно заголовки): {thickness}, {width}, {length}\n')
            continue

    try:
        # Конвертация в целые числа (поддерживает числа с плавающей точкой)
        thickness = int(float(thickness))
        width = int(float(width))
        length = int(float(length))
    except (ValueError, TypeError) as e:
        print(f'Ошибка в строке {row_idx}: {thickness}, {width}, {length} - {e}')
        errors += 1
        continue

    # Создание уникальной записи
    obj, created_flag = Dimension.objects.get_or_create(
        thickness=thickness,
        width=width,
        length=length
    )

    if created_flag:
        created += 1
        print(f'+ Строка {row_idx}: {thickness}x{width}x{length}')
    else:
        skipped += 1
        print(f'  Строка {row_idx}: {thickness}x{width}x{length} (уже существует)')

# Итоги
print('\n' + '=' * 50)
print('ИМПОРТ ЗАВЕРШЕН')
print(f'✅ Добавлено новых: {created}')
print(f'⏭️  Пропущено дубликатов: {skipped}')
print(f'❌ Ошибок: {errors}')
print(f'📊 Всего в таблице размеров: {Dimension.objects.count()}')
print('=' * 50)