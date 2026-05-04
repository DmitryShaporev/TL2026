# lumbertrack/admin.py

from django.contrib import admin
from .models import (
    ProductType, WoodSpecies, QualityGrade,
    ProductName, LumberDimension, UnitDimension, ProductItem
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