# lumber_track/views.py
import json
import json
from django.db import models
from django.db.models import Q
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
from .models import (
    ProductType, WoodSpecies, QualityGrade, ProductName,
    UnitDimension, LumberDimension, ProductItem
)
from django.db import IntegrityError
from django.http import JsonResponse
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.views.decorators.csrf import csrf_exempt

from .models import ProductType, WoodSpecies, QualityGrade, ProductName, UnitDimension, LumberDimension, ProductItem

# lumber_track/views.py - добавьте в начало файла

from django.db import connection


# lumber_track/views.py

def get_available_stocks_with_details(location_id=None, exclude_document_id=None):
    """Возвращает позиции с остатками и деталями для селекта расхода"""

    query = """
        WITH stock_calc AS (
            SELECT 
                ROW_NUMBER() OVER (ORDER BY pn.name, ws.name, qg.code) as id,
                di.product_name_id,
                pn.name as product_name,
                di.species_id,
                ws.name as species,
                di.grade_id,
                qg.code as grade,
                di.lumber_dim_id,
                ld.thickness,
                ld.width,
                ld.length,
                di.unit_dim_id,
                ud.length as unit_length,
                ud.width as unit_width,
                ud.height as unit_height,
                COALESCE(SUM(CASE WHEN d.doc_type = 1 THEN di.quantity ELSE 0 END), 0) +
                COALESCE(SUM(CASE WHEN d.doc_type = 2 THEN di.quantity ELSE 0 END), 0) -
                COALESCE(SUM(CASE WHEN d.doc_type = 3 THEN di.quantity ELSE 0 END), 0) as balance
            FROM lumber_track_documentitem di
            JOIN lumber_track_document d ON di.document_id = d.id
            JOIN lumber_track_productname pn ON di.product_name_id = pn.id
            JOIN lumber_track_woodspecies ws ON di.species_id = ws.id
            JOIN lumber_track_qualitygrade qg ON di.grade_id = qg.id
            LEFT JOIN lumber_track_lumberdimension ld ON di.lumber_dim_id = ld.id
            LEFT JOIN lumber_track_unitdimension ud ON di.unit_dim_id = ud.id
            WHERE d.doc_date <= date('now')
    """

    if location_id:
        query += f" AND d.location_id = {location_id}"

    # Исключаем текущий документ из расчёта
    if exclude_document_id:
        query += f" AND d.id != {exclude_document_id}"

    query += """
            GROUP BY 
                di.product_name_id, pn.name,
                di.species_id, ws.name,
                di.grade_id, qg.code,
                di.lumber_dim_id, ld.thickness, ld.width, ld.length,
                di.unit_dim_id, ud.length, ud.width, ud.height
            HAVING balance > 0
        )
        SELECT 
            id,
            product_name_id,
            product_name,
            species_id,
            species,
            grade_id,
            grade,
            lumber_dim_id,
            thickness,
            width,
            length,
            unit_dim_id,
            unit_length,
            unit_width,
            unit_height,
            balance,
            CASE 
                WHEN lumber_dim_id IS NOT NULL THEN 
                    thickness || '-' || width || '-' || length || ' мм'
                ELSE 
                    unit_length || '-' || unit_width || '-' || unit_height || ' мм'
            END as dimension_display,
            CASE 
                WHEN lumber_dim_id IS NOT NULL THEN 'lumber'
                ELSE 'unit'
            END as dimension_type,
            COALESCE(lumber_dim_id, unit_dim_id) as dimension_id
        FROM stock_calc
        ORDER BY product_name, species, grade, dimension_display
    """

    with connection.cursor() as cursor:
        cursor.execute(query)
        columns = [col[0] for col in cursor.description]
        results = [dict(zip(columns, row)) for row in cursor.fetchall()]

    return results
# ========== ГЛАВНАЯ И СПРАВОЧНИКИ ==========
def home_view(request):
    return render(request, 'lumber_track/home.html')

def directories_view(request):
    return render(request, 'lumber_track/directories.html')


# ========== ТИПЫ ПРОДУКЦИИ ==========
def producttype_list(request):
    items = ProductType.objects.all().order_by('name')
    context = {
        'title': 'Типы продукции',
        'items': items,
        'create_url': '/directories/producttype/create/',
        'delete_url': 'lumber_track:producttype_delete',
    }
    return render(request, 'lumber_track/directory_table.html', context)

def producttype_create(request):
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        if name:
            ProductType.objects.create(name=name)

    return redirect('lumber_track:producttype_list')

def producttype_edit(request, pk):
    item = get_object_or_404(ProductType, pk=pk)
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        if name:
            item.name = name
            item.save()

    return redirect('lumber_track:producttype_list')

def producttype_delete(request, pk):
    item = get_object_or_404(ProductType, pk=pk)

    item.delete()

    return redirect('lumber_track:producttype_list')


# ========== ПОРОДЫ ДРЕВЕСИНЫ ==========
def woodspecies_list(request):
    items = WoodSpecies.objects.all().order_by('name')
    context = {
        'title': 'Породы древесины',
        'items': items,
        'create_url': '/directories/woodspecies/create/',
        'delete_url': 'lumber_track:woodspecies_delete',
    }
    return render(request, 'lumber_track/directory_table.html', context)

def woodspecies_create(request):
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        if name:
            WoodSpecies.objects.create(name=name)

    return redirect('lumber_track:woodspecies_list')

def woodspecies_edit(request, pk):
    item = get_object_or_404(WoodSpecies, pk=pk)
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        if name:
            item.name = name
            item.save()

    return redirect('lumber_track:woodspecies_list')

def woodspecies_delete(request, pk):
    item = get_object_or_404(WoodSpecies, pk=pk)

    item.delete()

    return redirect('lumber_track:woodspecies_list')


# ========== КАТЕГОРИИ КАЧЕСТВА ==========
def qualitygrade_list(request):
    items = QualityGrade.objects.all().order_by('code')
    context = {
        'title': 'Категории качества',
        'items': items,
        'create_url': '/directories/qualitygrade/create/',
        'delete_url': 'lumber_track:qualitygrade_delete',
    }
    return render(request, 'lumber_track/directory_table.html', context)

def qualitygrade_create(request):
    if request.method == 'POST':
        code = request.POST.get('name', '').strip().upper()  # Категории хранятся в поле code
        if code:
            QualityGrade.objects.create(code=code)

    return redirect('lumber_track:qualitygrade_list')

def qualitygrade_edit(request, pk):
    item = get_object_or_404(QualityGrade, pk=pk)
    if request.method == 'POST':
        code = request.POST.get('name', '').strip().upper()
        if code:
            item.code = code
            item.save()

    return redirect('lumber_track:qualitygrade_list')

def qualitygrade_delete(request, pk):
    item = get_object_or_404(QualityGrade, pk=pk)

    item.delete()

    return redirect('lumber_track:qualitygrade_list')


# ========== API для получения данных (редактирование) ==========
def producttype_data(request, pk):
    item = get_object_or_404(ProductType, pk=pk)
    return JsonResponse({'name': item.name})

def woodspecies_data(request, pk):
    item = get_object_or_404(WoodSpecies, pk=pk)
    return JsonResponse({'name': item.name})

def qualitygrade_data(request, pk):
    item = get_object_or_404(QualityGrade, pk=pk)
    return JsonResponse({'name': item.code})  # у QualityGrade поле code


# ========== НАИМЕНОВАНИЯ ИЗДЕЛИЙ ==========
def productname_list(request):
    items = ProductName.objects.all().select_related('product_type').order_by('name')
    product_types = ProductType.objects.all().order_by('name')
    context = {
        'title': 'Наименования изделий',
        'items': items,
        'product_types': product_types,
        'create_url': '/directories/productname/create/',
        'delete_url': 'lumber_track:productname_delete',
    }
    return render(request, 'lumber_track/directory_productname.html', context)

def productname_create(request):
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        product_type_id = request.POST.get('product_type')
        if name and product_type_id:
            product_type = get_object_or_404(ProductType, pk=product_type_id)
            ProductName.objects.create(name=name, product_type=product_type)
    return redirect('lumber_track:productname_list')

def productname_edit(request, pk):
    item = get_object_or_404(ProductName, pk=pk)
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        product_type_id = request.POST.get('product_type')
        if name and product_type_id:
            item.name = name
            item.product_type_id = product_type_id
            item.save()
    return redirect('lumber_track:productname_list')

def productname_delete(request, pk):
    item = get_object_or_404(ProductName, pk=pk)
    item.delete()
    return redirect('lumber_track:productname_list')

def productname_data(request, pk):
    item = get_object_or_404(ProductName, pk=pk)
    return JsonResponse({
        'name': item.name,
        'product_type_id': item.product_type_id
    })


# ========== РАЗМЕРЫ ШТУЧНЫХ ИЗДЕЛИЙ ==========
def unitdimension_list(request):
    items = UnitDimension.objects.all().order_by('length', 'width', 'height')
    context = {
        'title': 'Размеры изделий',
        'items': items,
        'create_url': '/directories/unitdimension/create/',
        'delete_url': 'lumber_track:unitdimension_delete',
    }
    return render(request, 'lumber_track/directory_unitdimension.html', context)


def unitdimension_create(request):
    if request.method == 'POST':
        length = request.POST.get('length', '').strip()
        width = request.POST.get('width', '').strip()
        height = request.POST.get('height', '').strip()

        if length and width and height:
            try:
                UnitDimension.objects.create(
                    length=int(length),
                    width=int(width),
                    height=int(height)
                )
            except (ValueError, IntegrityError):
                pass  # игнорируем ошибки (например, дубликаты)
    return redirect('lumber_track:unitdimension_list')


def unitdimension_edit(request, pk):
    item = get_object_or_404(UnitDimension, pk=pk)
    if request.method == 'POST':
        length = request.POST.get('length', '').strip()
        width = request.POST.get('width', '').strip()
        height = request.POST.get('height', '').strip()

        if length and width and height:
            try:
                item.length = int(length)
                item.width = int(width)
                item.height = int(height)
                item.save()
            except ValueError:
                pass
    return redirect('lumber_track:unitdimension_list')


def unitdimension_delete(request, pk):
    item = get_object_or_404(UnitDimension, pk=pk)
    item.delete()
    return redirect('lumber_track:unitdimension_list')


def unitdimension_data(request, pk):
    item = get_object_or_404(UnitDimension, pk=pk)
    return JsonResponse({
        'length': item.length,
        'width': item.width,
        'height': item.height,
        'display': str(item)
    })


# ========== РАЗМЕРЫ ПОГОНАЖА ==========
def lumberdimension_list(request):
    items = LumberDimension.objects.all().order_by('thickness', 'width', 'length')
    context = {
        'title': 'Размеры погонажа',
        'items': items,
        'create_url': '/directories/lumberdimension/create/',
        'delete_url': 'lumber_track:lumberdimension_delete',
    }
    return render(request, 'lumber_track/directory_lumberdimension.html', context)


def lumberdimension_create(request):
    if request.method == 'POST':
        thickness = request.POST.get('thickness', '').strip()
        width = request.POST.get('width', '').strip()
        length = request.POST.get('length', '').strip()

        if thickness and width and length:
            try:
                LumberDimension.objects.create(
                    thickness=int(thickness),
                    width=int(width),
                    length=int(length)
                )
            except (ValueError, IntegrityError):
                pass  # игнорируем ошибки (например, дубликаты)
    return redirect('lumber_track:lumberdimension_list')


def lumberdimension_edit(request, pk):
    item = get_object_or_404(LumberDimension, pk=pk)
    if request.method == 'POST':
        thickness = request.POST.get('thickness', '').strip()
        width = request.POST.get('width', '').strip()
        length = request.POST.get('length', '').strip()

        if thickness and width and length:
            try:
                item.thickness = int(thickness)
                item.width = int(width)
                item.length = int(length)
                item.save()
            except ValueError:
                pass
    return redirect('lumber_track:lumberdimension_list')


def lumberdimension_delete(request, pk):
    item = get_object_or_404(LumberDimension, pk=pk)
    item.delete()
    return redirect('lumber_track:lumberdimension_list')


def lumberdimension_data(request, pk):
    item = get_object_or_404(LumberDimension, pk=pk)
    return JsonResponse({
        'thickness': item.thickness,
        'width': item.width,
        'length': item.length,
        'display': str(item),
        'volume': round(item.volume_m3, 6),
        'area': round(item.area_m2, 3)
    })


# ========== СПРАВОЧНИК ИЗДЕЛИЙ (основной) ==========
def productitem_list(request):
    items = ProductItem.objects.all().select_related(
        'product_name', 'species', 'grade', 'lumber_dim', 'unit_dim'
    ).order_by('-created_at')

    context = {
        'title': 'Справочник изделий',
        'items': items,
        'product_names': ProductName.objects.all().select_related('product_type'),
        'species_list': WoodSpecies.objects.all().order_by('name'),
        'grades_list': QualityGrade.objects.all().order_by('code'),
        'lumber_dims': LumberDimension.objects.all().order_by('thickness', 'width', 'length'),
        'unit_dims': UnitDimension.objects.all().order_by('length', 'width', 'height'),
        'delete_url': 'lumber_track:productitem_delete',
    }
    return render(request, 'lumber_track/directory_productitem.html', context)


def productitem_create(request):
    if request.method == 'POST':
        product_name_id = request.POST.get('product_name')
        species_id = request.POST.get('species')
        grade_id = request.POST.get('grade')
        lumber_dim_id = request.POST.get('lumber_dim')
        unit_dim_id = request.POST.get('unit_dim')

        if product_name_id and species_id and grade_id:
            ProductItem.objects.create(
                product_name_id=product_name_id,
                species_id=species_id,
                grade_id=grade_id,
                lumber_dim_id=lumber_dim_id or None,
                unit_dim_id=unit_dim_id or None,
                is_active=True
            )
    return redirect('lumber_track:productitem_list')


def productitem_delete(request, pk):
    item = get_object_or_404(ProductItem, pk=pk)
    item.delete()
    return redirect('lumber_track:productitem_list')


# ========== API для Select2 (быстрое добавление) ==========

@csrf_exempt
@require_http_methods(["POST"])
def api_add_productname(request):
    """Быстрое добавление наименования изделия"""
    try:
        data = json.loads(request.body)
        name = data.get('name', '').strip()
        product_type_id = data.get('product_type_id')

        if not name:
            return JsonResponse({'error': 'Название не может быть пустым'}, status=400)

        # Определяем тип продукции по умолчанию (если не выбран)
        if not product_type_id:
            # Пробуем найти тип "Погонаж" или создаем
            product_type, _ = ProductType.objects.get_or_create(name="Погонаж")
            product_type_id = product_type.id

        obj, created = ProductName.objects.get_or_create(
            name=name,
            defaults={'product_type_id': product_type_id}
        )

        return JsonResponse({
            'id': obj.id,
            'text': f"{obj.name} ({obj.product_type.name})",
            'created': created
        })
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=400)


@csrf_exempt
@require_http_methods(["POST"])
def api_add_woodspecies(request):
    """Быстрое добавление породы древесины"""
    try:
        data = json.loads(request.body)
        name = data.get('name', '').strip()

        if not name:
            return JsonResponse({'error': 'Название не может быть пустым'}, status=400)

        obj, created = WoodSpecies.objects.get_or_create(name=name)
        return JsonResponse({'id': obj.id, 'text': obj.name, 'created': created})
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=400)


@csrf_exempt
@require_http_methods(["POST"])
def api_add_qualitygrade(request):
    """Быстрое добавление категории качества"""
    try:
        data = json.loads(request.body)
        code = data.get('name', '').strip().upper()

        if not code:
            return JsonResponse({'error': 'Код не может быть пустым'}, status=400)

        obj, created = QualityGrade.objects.get_or_create(code=code)
        return JsonResponse({'id': obj.id, 'text': obj.code, 'created': created})
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=400)


@csrf_exempt
@require_http_methods(["POST"])
def api_add_lumberdimension(request):
    """Быстрое добавление размера погонажа"""
    try:
        data = json.loads(request.body)
        thickness = int(data.get('thickness', 0))
        width = int(data.get('width', 0))
        length = int(data.get('length', 0))

        if not all([thickness, width, length]):
            return JsonResponse({'error': 'Все размеры должны быть указаны'}, status=400)

        obj, created = LumberDimension.objects.get_or_create(
            thickness=thickness,
            width=width,
            length=length
        )

        return JsonResponse({
            'id': obj.id,
            'text': f"{thickness}-{width}-{length} мм ({obj.volume_m3:.6f} м³)",
            'created': created
        })
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=400)


@csrf_exempt
@require_http_methods(["POST"])
def api_add_unitdimension(request):
    """Быстрое добавление размера штучного изделия"""
    try:
        data = json.loads(request.body)
        length = int(data.get('length', 0))
        width = int(data.get('width', 0))
        height = int(data.get('height', 0))

        if not all([length, width, height]):
            return JsonResponse({'error': 'Все размеры должны быть указаны'}, status=400)

        obj, created = UnitDimension.objects.get_or_create(
            length=length,
            width=width,
            height=height
        )

        return JsonResponse({
            'id': obj.id,
            'text': f"{length}-{width}-{height} мм",
            'created': created
        })
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=400)


# ========== API для поиска (Select2) ==========

def api_search_productname(request):
    """Поиск наименований для Select2"""
    term = request.GET.get('term', '')
    items = ProductName.objects.filter(name__icontains=term).select_related('product_type')[:20]
    results = [{'id': item.id, 'text': f"{item.name} ({item.product_type.name})"} for item in items]
    return JsonResponse({'results': results})


def api_search_woodspecies(request):
    """Поиск пород для Select2"""
    term = request.GET.get('term', '')
    items = WoodSpecies.objects.filter(name__icontains=term)[:20]
    results = [{'id': item.id, 'text': item.name} for item in items]
    return JsonResponse({'results': results})


def api_search_qualitygrade(request):
    """Поиск категорий для Select2"""
    term = request.GET.get('term', '')
    items = QualityGrade.objects.filter(code__icontains=term)[:20]
    results = [{'id': item.id, 'text': item.code} for item in items]
    return JsonResponse({'results': results})


def api_search_lumberdim(request):
    """Поиск размеров погонажа для Select2"""
    term = request.GET.get('term', '')
    items = LumberDimension.objects.filter(
        models.Q(thickness__icontains=term) |
        models.Q(width__icontains=term) |
        models.Q(length__icontains=term)
    )[:20]
    results = [{'id': item.id, 'text': f"{item.thickness}-{item.width}-{item.length} мм ({item.volume_m3:.6f} м³)"} for
               item in items]
    return JsonResponse({'results': results})


def api_search_unitdim(request):
    """Поиск размеров штучных для Select2"""
    term = request.GET.get('term', '')
    items = UnitDimension.objects.filter(
        models.Q(length__icontains=term) |
        models.Q(width__icontains=term) |
        models.Q(height__icontains=term)
    )[:20]
    results = [{'id': item.id, 'text': f"{item.length}-{item.width}-{item.height} мм"} for item in items]
    return JsonResponse({'results': results})


def documents_page(request):
    """Страница с карточками документов"""
    return render(request, 'lumber_track/documents.html')


# lumber_track/views.py - добавьте функции

from .models import Document, DocumentItem
from django.db.models import Q


def document_journal(request, doc_type):
    """Универсальный журнал документов"""
    doc_type_name = dict(Document.DOCUMENT_TYPES).get(doc_type, 'Документы')
    documents = Document.objects.filter(doc_type=doc_type).prefetch_related('items').order_by('-doc_date',
                                                                                              '-created_at')

    # Для каждой записи добавляем флаг has_items
    for doc in documents:
        doc.has_items = doc.items.exists()

    context = {
        'title': doc_type_name,
        'documents': documents,
        'doc_type': doc_type,
        'delete_url': 'lumber_track:document_delete',
    }
    return render(request, 'lumber_track/document_journal.html', context)


def document_delete(request, pk):
    """Удаление документа только если нет позиций"""
    doc = get_object_or_404(Document, pk=pk)

    # Сохраняем тип для перенаправления
    doc_type = doc.doc_type
    doc_number = doc.doc_number

    # Проверяем, есть ли позиции
    if doc.items.exists():
        messages.error(request,
                       f'❌ Невозможно удалить документ "{doc_number}", так как он содержит позиции. Сначала удалите все позиции.')
    else:
        doc.delete()
        messages.success(request, f'✅ Документ "{doc_number}" успешно удален.')

    # Перенаправляем в соответствующий журнал
    if doc_type == 1:
        return redirect('lumber_track:document_initial_journal')
    elif doc_type == 2:
        return redirect('lumber_track:document_income_journal')
    else:
        return redirect('lumber_track:document_outcome_journal')


# lumber_track/views.py - функция document_create

# lumber_track/views.py
# lumber_track/views.py

# lumber_track/views.py

# lumber_track/views.py

def document_create(request, doc_type):
    """Создание документа"""
    doc_type_name = dict(Document.DOCUMENT_TYPES).get(doc_type, 'Документ')

    if request.method == 'POST':
        doc_number = request.POST.get('doc_number', '').strip()

        # Проверяем уникальность комбинации тип + номер
        if Document.objects.filter(doc_type=doc_type, doc_number=doc_number).exists():
            messages.error(request, f'Документ с номером "{doc_number}" для данного типа документа уже существует!')
            context = {
                'doc_type': doc_type,
                'doc_type_name': doc_type_name,
                'product_names': ProductName.objects.all().select_related('product_type'),
                'species_list': WoodSpecies.objects.all().order_by('name'),
                'grades_list': QualityGrade.objects.all().order_by('code'),
                'lumber_dims': LumberDimension.objects.all().order_by('thickness', 'width', 'length'),
                'unit_dims': UnitDimension.objects.all().order_by('length', 'width', 'height'),
                'locations': StorageLocation.objects.all().order_by('name'),
            }
            # Для расхода нужен другой шаблон при ошибке
            if doc_type == 3:
                available_stocks = get_available_stocks_with_details()
                context['available_stocks'] = available_stocks
                return render(request, 'lumber_track/document_outcome_form.html', context)
            return render(request, 'lumber_track/document_form.html', context)

        # Определяем location_id ДО создания документа
        location_id = request.POST.get('location')

        # Для начальных остатков
        if not location_id and doc_type == 1:
            location_id = 2  # ID склада готовой продукции

        # Для расхода - автоматически подставляем склад (откуда списываем)
        if not location_id and doc_type == 3:
            location_id = 2  # ID склада готовой продукции

        # Создаем документ
        doc = Document.objects.create(
            doc_type=doc_type,
            doc_number=doc_number,
            doc_date=request.POST.get('doc_date'),
            note=request.POST.get('note', ''),
            location_id=location_id
        )

        # Для расхода - обрабатываем позиции с остатками
        if doc_type == 3:  # Расход
            stock_items = request.POST.getlist('stock_item[]')
            quantities = request.POST.getlist('quantity[]')
            product_names = request.POST.getlist('product_name[]')
            species_list = request.POST.getlist('species[]')
            grades_list = request.POST.getlist('grade[]')
            dimension_ids = request.POST.getlist('dimension_id[]')
            dimension_types = request.POST.getlist('dimension_type[]')

            # Сохраняем "куда перемещается" в примечание
            to_location_id = request.POST.get('to_location')
            if to_location_id:
                to_location = StorageLocation.objects.get(id=to_location_id)
                note_text = request.POST.get('note', '')
                new_note = f"{note_text}\nПеремещено в: {to_location.name}" if note_text else f"Перемещено в: {to_location.name}"
                doc.note = new_note
                doc.save()

            for i in range(len(stock_items)):
                if not stock_items[i]:
                    continue

                quantity = int(quantities[i]) if i < len(quantities) and quantities[i] else 0
                if quantity == 0:
                    continue

                product_name_id = product_names[i] if i < len(product_names) else None
                species_id = species_list[i] if i < len(species_list) else None
                grade_id = grades_list[i] if i < len(grades_list) else None
                dimension_id = dimension_ids[i] if i < len(dimension_ids) else None
                dimension_type = dimension_types[i] if i < len(dimension_types) else None

                lumber_dim_id = None
                unit_dim_id = None

                if dimension_type == 'lumber':
                    lumber_dim_id = dimension_id
                elif dimension_type == 'unit':
                    unit_dim_id = dimension_id

                DocumentItem.objects.create(
                    document=doc,
                    product_name_id=product_name_id,
                    species_id=species_id,
                    grade_id=grade_id,
                    lumber_dim_id=lumber_dim_id,
                    unit_dim_id=unit_dim_id,
                    quantity=quantity
                )
        else:
            # Для начальных остатков и прихода - обычная обработка
            product_names = request.POST.getlist('product_name[]')
            species_list = request.POST.getlist('species[]')
            grades_list = request.POST.getlist('grade[]')
            dimension_ids = request.POST.getlist('dimension_id[]')
            dimension_types = request.POST.getlist('dimension_type[]')
            quantities = request.POST.getlist('quantity[]')

            for i in range(len(product_names)):
                if not product_names[i] or not species_list[i] or not grades_list[i]:
                    continue

                quantity = int(quantities[i]) if i < len(quantities) and quantities[i] else 0
                if quantity == 0:
                    continue

                lumber_dim_id = None
                unit_dim_id = None

                if i < len(dimension_ids) and dimension_ids[i]:
                    if i < len(dimension_types) and dimension_types[i] == 'lumber':
                        lumber_dim_id = dimension_ids[i]
                    elif i < len(dimension_types) and dimension_types[i] == 'unit':
                        unit_dim_id = dimension_ids[i]

                DocumentItem.objects.create(
                    document=doc,
                    product_name_id=product_names[i],
                    species_id=species_list[i],
                    grade_id=grades_list[i],
                    lumber_dim_id=lumber_dim_id,
                    unit_dim_id=unit_dim_id,
                    quantity=quantity
                )

        # Перенаправляем в журнал
        redirect_urls = {
            1: 'lumber_track:document_initial_journal',
            2: 'lumber_track:document_income_journal',
            3: 'lumber_track:document_outcome_journal',
        }
        return redirect(redirect_urls.get(doc_type, 'lumber_track:document_initial_journal'))

    # ========== GET ЗАПРОС ==========
    context = {
        'doc_type': doc_type,
        'doc_type_name': doc_type_name,
        'product_names': ProductName.objects.all().select_related('product_type'),
        'species_list': WoodSpecies.objects.all().order_by('name'),
        'grades_list': QualityGrade.objects.all().order_by('code'),
        'lumber_dims': LumberDimension.objects.all().order_by('thickness', 'width', 'length'),
        'unit_dims': UnitDimension.objects.all().order_by('length', 'width', 'height'),
        'locations': StorageLocation.objects.all().order_by('name'),
    }

    # Выбираем шаблон в зависимости от типа документа
    if doc_type == 3:  # Расход
        available_stocks = get_available_stocks_with_details()
        context['available_stocks'] = available_stocks
        return render(request, 'lumber_track/document_outcome_form.html', context)
    else:
        return render(request, 'lumber_track/document_form.html', context)
# API для получения данных о размере (для расчета объема и площади)
def api_get_lumberdim_data(request, pk):
    """Получение данных размера погонажа для расчета"""
    dim = get_object_or_404(LumberDimension, pk=pk)
    return JsonResponse({
        'volume': dim.volume_m3,
        'area': dim.area_m2
    })


# API для быстрого добавления новых справочников
@csrf_exempt
def api_add_dimension(request):
    """Быстрое добавление размера погонажа"""
    if request.method == 'POST':
        data = json.loads(request.body)
        thickness = int(data.get('thickness'))
        width = int(data.get('width'))
        length = int(data.get('length'))

        obj, created = LumberDimension.objects.get_or_create(
            thickness=thickness,
            width=width,
            length=length
        )
        return JsonResponse({
            'id': obj.id,
            'volume': obj.volume_m3,
            'area': obj.area_m2
        })
    return JsonResponse({'error': 'Ошибка'}, status=400)


@csrf_exempt
def api_add_productname(request):
    if request.method == 'POST':
        data = json.loads(request.body)
        name = data.get('name', '').strip()
        # Для простоты берем первый тип продукции
        product_type = ProductType.objects.first()
        if name and product_type:
            obj, created = ProductName.objects.get_or_create(
                name=name,
                defaults={'product_type': product_type}
            )
            return JsonResponse({'id': obj.id, 'text': obj.name})
    return JsonResponse({'error': 'Ошибка'}, status=400)


# lumber_track/views.py - добавьте

from .models import StorageLocation


# ========== МЕСТА ХРАНЕНИЯ ==========
# lumber_track/views.py

def storagelocation_list(request):
    items = StorageLocation.objects.all().order_by('name')
    context = {
        'title': 'Места хранения',
        'items': items,
        'create_url': '/directories/storagelocation/create/',
        'delete_url': 'lumber_track:storagelocation_delete',
    }
    return render(request, 'lumber_track/directory_table.html', context)

def storagelocation_create(request):
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        if name:
            StorageLocation.objects.create(name=name)
    return redirect('lumber_track:storagelocation_list')

def storagelocation_edit(request, pk):
    item = get_object_or_404(StorageLocation, pk=pk)
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        if name:
            item.name = name
            item.save()
    return redirect('lumber_track:storagelocation_list')

def storagelocation_delete(request, pk):
    item = get_object_or_404(StorageLocation, pk=pk)
    item.delete()
    return redirect('lumber_track:storagelocation_list')

def storagelocation_data(request, pk):
    item = get_object_or_404(StorageLocation, pk=pk)
    return JsonResponse({'name': item.name})


# lumber_track/views.py

# lumber_track/views.py

def document_edit(request, pk):
    """Редактирование документа"""
    doc = get_object_or_404(Document, pk=pk)
    doc_type = doc.doc_type
    doc_type_name = dict(Document.DOCUMENT_TYPES).get(doc_type, 'Документ')

    # Для расхода - извлекаем "куда перемещается" из примечания
    to_location_id = None
    to_location_name = None
    original_note = doc.note

    if doc_type == 3 and doc.note:
        # Ищем "Перемещено в:" в примечании
        import re
        match = re.search(r'Перемещено в: (.+)', doc.note)
        if match:
            to_location_name = match.group(1).strip()
            # Убираем эту строку из примечания для отображения
            original_note = re.sub(r'\n?Перемещено в: .+', '', doc.note).strip()

    if request.method == 'POST':
        # Обновляем шапку документа
        doc.doc_number = request.POST.get('doc_number')
        doc.doc_date = request.POST.get('doc_date')

        # Для расхода - сохраняем "куда перемещается"
        if doc_type == 3:
            to_location_id = request.POST.get('to_location')
            to_location = StorageLocation.objects.get(id=to_location_id)
            note_text = request.POST.get('note', '')
            new_note = f"{note_text}\nПеремещено в: {to_location.name}" if note_text else f"Перемещено в: {to_location.name}"
            doc.note = new_note
        else:
            doc.note = request.POST.get('note', '')

        doc.save()

        # Удаляем старые позиции
        doc.items.all().delete()

        if doc_type == 3:  # Расход
            # Сохраняем позиции из формы расхода
            stock_items = request.POST.getlist('stock_item[]')
            quantities = request.POST.getlist('quantity[]')
            product_names = request.POST.getlist('product_name[]')
            species_list = request.POST.getlist('species[]')
            grades_list = request.POST.getlist('grade[]')
            dimension_ids = request.POST.getlist('dimension_id[]')
            dimension_types = request.POST.getlist('dimension_type[]')

            for i in range(len(stock_items)):
                if not stock_items[i]:
                    continue

                quantity = int(quantities[i]) if i < len(quantities) and quantities[i] else 0
                if quantity == 0:
                    continue

                product_name_id = product_names[i] if i < len(product_names) else None
                species_id = species_list[i] if i < len(species_list) else None
                grade_id = grades_list[i] if i < len(grades_list) else None
                dimension_id = dimension_ids[i] if i < len(dimension_ids) else None
                dimension_type = dimension_types[i] if i < len(dimension_types) else None

                lumber_dim_id = None
                unit_dim_id = None

                if dimension_type == 'lumber':
                    lumber_dim_id = dimension_id
                elif dimension_type == 'unit':
                    unit_dim_id = dimension_id

                DocumentItem.objects.create(
                    document=doc,
                    product_name_id=product_name_id,
                    species_id=species_id,
                    grade_id=grade_id,
                    lumber_dim_id=lumber_dim_id,
                    unit_dim_id=unit_dim_id,
                    quantity=quantity
                )
        else:
            # Для начальных остатков и прихода
            product_names = request.POST.getlist('product_name[]')
            species_list = request.POST.getlist('species[]')
            grades_list = request.POST.getlist('grade[]')
            dimension_ids = request.POST.getlist('dimension_id[]')
            dimension_types = request.POST.getlist('dimension_type[]')
            quantities = request.POST.getlist('quantity[]')

            for i in range(len(product_names)):
                if not product_names[i] or not species_list[i] or not grades_list[i]:
                    continue

                quantity = int(quantities[i]) if i < len(quantities) and quantities[i] else 0
                if quantity == 0:
                    continue

                lumber_dim_id = None
                unit_dim_id = None

                if i < len(dimension_ids) and dimension_ids[i]:
                    if i < len(dimension_types) and dimension_types[i] == 'lumber':
                        lumber_dim_id = dimension_ids[i]
                    elif i < len(dimension_types) and dimension_types[i] == 'unit':
                        unit_dim_id = dimension_ids[i]

                DocumentItem.objects.create(
                    document=doc,
                    product_name_id=product_names[i],
                    species_id=species_list[i],
                    grade_id=grades_list[i],
                    lumber_dim_id=lumber_dim_id,
                    unit_dim_id=unit_dim_id,
                    quantity=quantity
                )

        # Перенаправляем в журнал
        redirect_urls = {
            1: 'lumber_track:document_initial_journal',
            2: 'lumber_track:document_income_journal',
            3: 'lumber_track:document_outcome_journal',
        }
        return redirect(redirect_urls.get(doc_type, 'lumber_track:document_initial_journal'))

    # GET запрос
    context = {
        'doc': doc,
        'doc_type': doc_type,
        'doc_type_name': doc_type_name,
        'original_note': original_note,
        'to_location_name': to_location_name,
        'product_names': ProductName.objects.all().select_related('product_type'),
        'species_list': WoodSpecies.objects.all().order_by('name'),
        'grades_list': QualityGrade.objects.all().order_by('code'),
        'lumber_dims': LumberDimension.objects.all().order_by('thickness', 'width', 'length'),
        'unit_dims': UnitDimension.objects.all().order_by('length', 'width', 'height'),
        'locations': StorageLocation.objects.all().order_by('name'),
        'items': doc.items.all().select_related('product_name', 'species', 'grade', 'lumber_dim', 'unit_dim'),
    }

    # Для расхода - добавляем остатки (исключая текущий документ)
    if doc_type == 3:
        available_stocks = get_available_stocks_with_details(
            location_id=2,  # склад
            exclude_document_id=doc.id  # <-- ЭТО ГЛАВНОЕ: исключаем текущий документ
        )
        context['available_stocks'] = available_stocks
        return render(request, 'lumber_track/document_outcome_edit.html', context)
    else:
        return render(request, 'lumber_track/document_edit.html', context)