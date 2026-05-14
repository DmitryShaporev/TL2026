# lumbertrack/admin.py

from django.contrib import admin
from .models import (
    ProductType, WoodSpecies, QualityGrade,
    ProductName, LumberDimension, UnitDimension, ProductItem, DocumentItem, Document
)


@admin.register(ProductType)
class ProductTypeAdmin(admin.ModelAdmin):
    list_display = ['name']
    search_fields = ['name']


@admin.register(WoodSpecies)
class WoodSpeciesAdmin(admin.ModelAdmin):
    list_display = ['name']
    search_fields = ['name']


@admin.register(QualityGrade)
class QualityGradeAdmin(admin.ModelAdmin):
    list_display = ['code', 'description']
    search_fields = ['code']


@admin.register(ProductName)
class ProductNameAdmin(admin.ModelAdmin):
    list_display = ['name', 'product_type']
    list_filter = ['product_type']
    search_fields = ['name']


@admin.register(LumberDimension)
class LumberDimensionAdmin(admin.ModelAdmin):
    list_display = ['thickness', 'width', 'length', 'volume_m3', 'area_m2']
    list_filter = ['thickness', 'width']
    search_fields = ['thickness', 'width', 'length']


@admin.register(UnitDimension)
class UnitDimensionAdmin(admin.ModelAdmin):
    list_display = ['length', 'width', 'height']
    list_filter = ['length', 'width']
    search_fields = ['length', 'width', 'height']


@admin.register(ProductItem)
class ProductItemAdmin(admin.ModelAdmin):
    list_display = ['full_name', 'product_name', 'species', 'grade', 'is_active']
    list_filter = ['product_name__product_type', 'species', 'grade', 'is_active']
    search_fields = ['product_name__name', 'species__name', 'grade__code']
    list_editable = ['is_active']

    fieldsets = (
        ('Основная информация', {
            'fields': ('product_name', 'species', 'grade', 'is_active')
        }),
        ('Размеры', {
            'fields': ('lumber_dim', 'unit_dim'),
            'description': 'Заполните размеры в зависимости от типа продукции'
        }),
    )


class DocumentItemInline(admin.TabularInline):
    """Позиции документа внутри документа (табличная часть)"""
    model = DocumentItem
    extra = 1
    fields = ['product_name', 'species', 'grade', 'lumber_dim', 'unit_dim', 'quantity']

    # Настройки для отображения как select
    autocomplete_fields = ['product_name', 'species', 'grade']

    # Для размеров используем обычные select с поиском
    raw_id_fields = []  # убираем raw_id_fields

    def formfield_for_foreignkey(self, db_field, request, **kwargs):
        """Настраиваем отображение выбора размеров"""
        if db_field.name == 'lumber_dim':
            # Сортируем размеры для удобства
            kwargs['queryset'] = LumberDimension.objects.all().order_by('thickness', 'width', 'length')
            # Добавляем пустой вариант
            kwargs['empty_label'] = '---------'
        elif db_field.name == 'unit_dim':
            kwargs['queryset'] = UnitDimension.objects.all().order_by('length', 'width', 'height')
            kwargs['empty_label'] = '---------'
        return super().formfield_for_foreignkey(db_field, request, **kwargs)


@admin.register(Document)
class DocumentAdmin(admin.ModelAdmin):
    list_display = ['doc_number', 'get_doc_type_display', 'doc_date', 'location', 'created_at']
    list_filter = ['doc_type', 'doc_date', 'location']
    search_fields = ['doc_number', 'note']
    date_hierarchy = 'doc_date'
    inlines = [DocumentItemInline]
    fieldsets = (
        ('Основная информация', {
            'fields': ('doc_type', 'doc_number', 'doc_date', 'location')
        }),
        ('Дополнительно', {
            'fields': ('note',),
            'classes': ('collapse',)
        }),
    )

@admin.register(DocumentItem)
class DocumentItemAdmin(admin.ModelAdmin):
    list_display = ['document', 'product_name', 'species', 'grade', 'dimension_display', 'quantity']
    list_filter = ['document__doc_type', 'product_name', 'species', 'grade']
    search_fields = ['document__doc_number', 'product_name__name']

    def formfield_for_foreignkey(self, db_field, request, **kwargs):
        """Настраиваем отображение выбора размеров"""
        if db_field.name == 'lumber_dim':
            kwargs['queryset'] = LumberDimension.objects.all().order_by('thickness', 'width', 'length')
            kwargs['empty_label'] = '---------'
        elif db_field.name == 'unit_dim':
            kwargs['queryset'] = UnitDimension.objects.all().order_by('length', 'width', 'height')
            kwargs['empty_label'] = '---------'
        return super().formfield_for_foreignkey(db_field, request, **kwargs)

    def dimension_display(self, obj):
        return obj.dimension_display

    dimension_display.short_description = "Размер"

from .models import StorageLocation

@admin.register(StorageLocation)

class StorageLocationAdmin(admin.ModelAdmin):
    list_display = ['name', 'responsible_person', 'created_at']
    search_fields = ['name', 'responsible_person']
    fieldsets = (
        ('Основная информация', {
            'fields': ('name', 'responsible_person')
        }),
    )