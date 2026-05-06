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