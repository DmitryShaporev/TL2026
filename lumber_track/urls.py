# lumber_track/urls.py
from django.urls import path
from . import views

app_name = 'lumber_track'

urlpatterns = [
    # Главная и справочники
    path('', views.home_view, name='home'),
    path('directories/', views.directories_view, name='directories'),

    # ========== Типы продукции ==========
    path('directories/producttype/', views.producttype_list, name='producttype_list'),
    path('directories/producttype/create/', views.producttype_create, name='producttype_create'),
    path('directories/producttype/<int:pk>/edit/', views.producttype_edit, name='producttype_edit'),
    path('directories/producttype/<int:pk>/delete/', views.producttype_delete, name='producttype_delete'),

    # ========== Породы древесины ==========
    path('directories/woodspecies/', views.woodspecies_list, name='woodspecies_list'),
    path('directories/woodspecies/create/', views.woodspecies_create, name='woodspecies_create'),
    path('directories/woodspecies/<int:pk>/edit/', views.woodspecies_edit, name='woodspecies_edit'),
    path('directories/woodspecies/<int:pk>/delete/', views.woodspecies_delete, name='woodspecies_delete'),

    # ========== Категории качества ==========
    path('directories/qualitygrade/', views.qualitygrade_list, name='qualitygrade_list'),
    path('directories/qualitygrade/create/', views.qualitygrade_create, name='qualitygrade_create'),
    path('directories/qualitygrade/<int:pk>/edit/', views.qualitygrade_edit, name='qualitygrade_edit'),
    path('directories/qualitygrade/<int:pk>/delete/', views.qualitygrade_delete, name='qualitygrade_delete'),

path('directories/producttype/<int:pk>/data/', views.producttype_data, name='producttype_data'),
    path('directories/woodspecies/<int:pk>/data/', views.woodspecies_data, name='woodspecies_data'),
    path('directories/qualitygrade/<int:pk>/data/', views.qualitygrade_data, name='qualitygrade_data'),


# ========== Наименования изделий ==========
path('directories/productname/', views.productname_list, name='productname_list'),
path('directories/productname/create/', views.productname_create, name='productname_create'),
path('directories/productname/<int:pk>/edit/', views.productname_edit, name='productname_edit'),
path('directories/productname/<int:pk>/delete/', views.productname_delete, name='productname_delete'),
path('directories/productname/<int:pk>/data/', views.productname_data, name='productname_data'),



# lumber_track/urls.py (добавьте в urlpatterns)

# ========== Размеры штучных изделий ==========
path('directories/unitdimension/', views.unitdimension_list, name='unitdimension_list'),
path('directories/unitdimension/create/', views.unitdimension_create, name='unitdimension_create'),
path('directories/unitdimension/<int:pk>/edit/', views.unitdimension_edit, name='unitdimension_edit'),
path('directories/unitdimension/<int:pk>/delete/', views.unitdimension_delete, name='unitdimension_delete'),
path('directories/unitdimension/<int:pk>/data/', views.unitdimension_data, name='unitdimension_data'),

# lumber_track/urls.py (добавьте в urlpatterns)

# ========== Размеры погонажа ==========
path('directories/lumberdimension/', views.lumberdimension_list, name='lumberdimension_list'),
path('directories/lumberdimension/create/', views.lumberdimension_create, name='lumberdimension_create'),
path('directories/lumberdimension/<int:pk>/edit/', views.lumberdimension_edit, name='lumberdimension_edit'),
path('directories/lumberdimension/<int:pk>/delete/', views.lumberdimension_delete, name='lumberdimension_delete'),
path('directories/lumberdimension/<int:pk>/data/', views.lumberdimension_data, name='lumberdimension_data'),

    # ========== Справочник изделий ==========
    path('directories/productitem/', views.productitem_list, name='productitem_list'),
    path('directories/productitem/create/', views.productitem_create, name='productitem_create'),
    path('directories/productitem/<int:pk>/delete/', views.productitem_delete, name='productitem_delete'),

    # ========== API для Select2 (поиск) ==========
    path('api/productname/search/', views.api_search_productname, name='api_search_productname'),
    path('api/woodspecies/search/', views.api_search_woodspecies, name='api_search_woodspecies'),
    path('api/qualitygrade/search/', views.api_search_qualitygrade, name='api_search_qualitygrade'),
    path('api/lumberdim/search/', views.api_search_lumberdim, name='api_search_lumberdim'),
    path('api/unitdim/search/', views.api_search_unitdim, name='api_search_unitdim'),

    # ========== API для Select2 (быстрое добавление) ==========
    path('api/productname/add/', views.api_add_productname, name='api_add_productname'),
    path('api/woodspecies/add/', views.api_add_woodspecies, name='api_add_woodspecies'),
    path('api/qualitygrade/add/', views.api_add_qualitygrade, name='api_add_qualitygrade'),
    path('api/lumberdim/add/', views.api_add_lumberdimension, name='api_add_lumberdimension'),
    path('api/unitdim/add/', views.api_add_unitdimension, name='api_add_unitdimension'),

# lumber_track/urls.py - добавьте в urlpatterns

    path('documents/', views.documents_page, name='documents_page'),

# Журналы документов
path('documents/initial/', views.document_journal, {'doc_type': 1}, name='document_initial_journal'),
path('documents/income/', views.document_journal, {'doc_type': 2}, name='document_income_journal'),
path('documents/outcome/', views.document_journal, {'doc_type': 3}, name='document_outcome_journal'),

# Создание, редактирование, удаление документа
path('documents/create/<int:doc_type>/', views.document_create, name='document_create'),
#path('documents/<int:pk>/edit/', views.document_edit, name='document_edit'),
path('documents/<int:pk>/delete/', views.document_delete, name='document_delete'),

# API
path('api/lumberdim/<int:pk>/', views.api_get_lumberdim_data, name='api_lumberdim_data'),
path('api/dimension/add/', views.api_add_dimension, name='api_add_dimension'),

path('api/productname/add/', views.api_add_productname, name='api_add_productname'),


# ========== Места хранения ==========
path('directories/storagelocation/', views.storagelocation_list, name='storagelocation_list'),
path('directories/storagelocation/create/', views.storagelocation_create, name='storagelocation_create'),
path('directories/storagelocation/<int:pk>/edit/', views.storagelocation_edit, name='storagelocation_edit'),
path('directories/storagelocation/<int:pk>/delete/', views.storagelocation_delete, name='storagelocation_delete'),
path('directories/storagelocation/<int:pk>/data/', views.storagelocation_data, name='storagelocation_data'),


# Редактирование документа
path('documents/<int:pk>/edit/', views.document_edit, name='document_edit'),

# Страница отчетов
path('reports/', views.reports_page, name='reports_page'),

# Отчет: Поступление
path('reports/income/', views.report_income, name='report_income'),
path('reports/income/result/', views.report_income_result, name='report_income_result'),

# Отчет: На склад
path('reports/to-stock/', views.report_to_stock, name='report_to_stock'),
path('reports/to-stock/result/', views.report_to_stock_result, name='report_to_stock_result'),

# Отчет: В магазин
path('reports/to-shop/', views.report_to_shop, name='report_to_shop'),
path('reports/to-shop/result/', views.report_to_shop_result, name='report_to_shop_result'),

# Отчет: Движение продукции
path('reports/movement/', views.report_movement, name='report_movement'),
path('reports/movement/result/', views.report_movement_result, name='report_movement_result'),

# Отчет: Сводный по категориям
path('reports/category/', views.report_category, name='report_category'),
path('reports/category/result/', views.report_category_result, name='report_category_result'),

]