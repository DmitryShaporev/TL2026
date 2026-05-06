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
]