# lumbertrack/models.py

from django.db import models


class ProductType(models.Model):
    name = models.CharField(max_length=50, unique=True, verbose_name="Название")

    class Meta:
        verbose_name = "Тип продукции"
        verbose_name_plural = "Типы продукции"

    def __str__(self):
        return self.name


class WoodSpecies(models.Model):
    name = models.CharField(max_length=50, unique=True, verbose_name="Порода")

    class Meta:
        verbose_name = "Порода древесины"
        verbose_name_plural = "Породы древесины"

    def __str__(self):
        return self.name


class QualityGrade(models.Model):
    code = models.CharField(max_length=10, unique=True, verbose_name="Код")
    description = models.CharField(max_length=100, blank=True, verbose_name="Описание")

    class Meta:
        verbose_name = "Категория качества"
        verbose_name_plural = "Категории качества"

    def __str__(self):
        return self.code


# lumber_track/models.py

class StorageLocation(models.Model):
    name = models.CharField(max_length=100, unique=True, verbose_name="Название")
    responsible_person = models.CharField(
        max_length=100,
        blank=True,
        verbose_name="Ответственный (ФИО)"
    )
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата создания")

    class Meta:
        verbose_name = "Место хранения"
        verbose_name_plural = "Места хранения"
        ordering = ['name']

    def __str__(self):
        return self.name

class ProductName(models.Model):
    name = models.CharField(max_length=100, unique=True, verbose_name="Наименование")
    product_type = models.ForeignKey(ProductType, on_delete=models.PROTECT, verbose_name="Тип продукции")

    class Meta:
        verbose_name = "Наименование изделия"
        verbose_name_plural = "Наименования изделий"
        ordering = ['product_type', 'name', ]

    def __str__(self):
        return self.name


# ========== ДВЕ ТАБЛИЦЫ РАЗМЕРОВ ==========

class LumberDimension(models.Model):
    """Размеры для погонажных изделий"""
    thickness = models.PositiveIntegerField(verbose_name="Толщина, мм")
    width = models.PositiveIntegerField(verbose_name="Ширина, мм")
    length = models.PositiveIntegerField(verbose_name="Длина, мм")

    class Meta:
        ordering = ['thickness', 'width', 'length']
        unique_together = ['thickness', 'width', 'length']
        verbose_name = "Размер погонажа"
        verbose_name_plural = "Размеры погонажа"

    @property
    def volume_m3(self):
        """Объем в кубометрах"""
        return (self.thickness * self.width * self.length) / 1_000_000_000

    @property
    def area_m2(self):
        """Площадь в кв.метрах"""
        return (self.width * self.length) / 1_000_000

    def __str__(self):
        return f"{self.thickness}x{self.width}x{self.length} мм"


class UnitDimension(models.Model):
    """Размеры для штучных изделий"""
    length = models.PositiveIntegerField(verbose_name="Длина, мм")
    width = models.PositiveIntegerField(verbose_name="Ширина, мм")
    height = models.PositiveIntegerField(verbose_name="Высота, мм")

    class Meta:
        ordering = ['length', 'width', 'height']
        unique_together = ['length', 'width', 'height']
        verbose_name = "Размер изделия"
        verbose_name_plural = "Размеры изделий"

    def __str__(self):
        return f"{self.length}x{self.width}x{self.height} мм"


# ========== ИТОГОВЫЙ СПРАВОЧНИК ==========

class ProductItem(models.Model):
    product_name = models.ForeignKey(ProductName, on_delete=models.PROTECT, verbose_name="Наименование")
    species = models.ForeignKey(WoodSpecies, on_delete=models.PROTECT, verbose_name="Порода")
    grade = models.ForeignKey(QualityGrade, on_delete=models.PROTECT, verbose_name="Категория")

    # Две возможные размерности (одна из них будет null)
    lumber_dim = models.ForeignKey(LumberDimension, on_delete=models.PROTECT,
                                   null=True, blank=True, verbose_name="Размеры (погонаж)")
    unit_dim = models.ForeignKey(UnitDimension, on_delete=models.PROTECT,
                                 null=True, blank=True, verbose_name="Размеры (штучные)")

    is_active = models.BooleanField(default=True, verbose_name="Активно")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата создания")

    class Meta:
        verbose_name = "Изделие"
        verbose_name_plural = "Изделия"
        # Уникальность с учетом типа размеров
        unique_together = ['product_name', 'species', 'grade', 'lumber_dim', 'unit_dim']

    @property
    def full_name(self):
        parts = [self.product_name.name, self.species.name, self.grade.code]

        if self.lumber_dim:
            parts.append(str(self.lumber_dim))
            if self.product_name.product_type.name == "Погонаж":
                parts.append(f"{self.lumber_dim.volume_m3:.4f} м³")
        elif self.unit_dim:
            parts.append(str(self.unit_dim))

        return ", ".join(parts)

    def __str__(self):
        return self.full_name


# lumber_track/models.py

class Document(models.Model):
    DOCUMENT_TYPES = [
        (1, 'Начальные остатки'),
        (2, 'Приход (выпуск)'),
        (3, 'Расход (отгрузка)'),
    ]

    doc_type = models.PositiveSmallIntegerField(choices=DOCUMENT_TYPES, verbose_name="Тип документа")
    doc_number = models.CharField(max_length=50,  verbose_name="Номер документа")
    doc_date = models.DateField(verbose_name="Дата документа")
    note = models.TextField(blank=True, verbose_name="Примечание")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата создания")

    # Новое поле - место хранения
    location = models.ForeignKey(StorageLocation, on_delete=models.PROTECT, related_name='documents_from',
                                 verbose_name="Откуда")
    to_location = models.ForeignKey(StorageLocation, on_delete=models.PROTECT, null=True, blank=True,
                                    related_name='documents_to', verbose_name="Куда")

    class Meta:
        ordering = ['-doc_date', '-created_at']
        verbose_name = "Документ"
        verbose_name_plural = "Документы"
        unique_together = ['doc_type', 'doc_number']

    def __str__(self):
        return f"{self.get_doc_type_display()} №{self.doc_number} от {self.doc_date}"



class DocumentItem(models.Model):
    document = models.ForeignKey(Document, on_delete=models.CASCADE, related_name='items', verbose_name="Документ")

    product_name = models.ForeignKey(ProductName, on_delete=models.PROTECT, verbose_name="Наименование")
    species = models.ForeignKey(WoodSpecies, on_delete=models.PROTECT, verbose_name="Порода")
    grade = models.ForeignKey(QualityGrade, on_delete=models.PROTECT, verbose_name="Категория")
    lumber_dim = models.ForeignKey(LumberDimension, on_delete=models.PROTECT, null=True, blank=True,
                                   verbose_name="Размер (погонаж)")
    unit_dim = models.ForeignKey(UnitDimension, on_delete=models.PROTECT, null=True, blank=True,
                                 verbose_name="Размер (штучный)")

    quantity = models.PositiveIntegerField(default=0, verbose_name="Количество (шт)")

    class Meta:
        verbose_name = "Позиция документа"
        verbose_name_plural = "Позиции документов"

    def __str__(self):
        return f"{self.product_name.name} - {self.quantity} шт"

    @property
    def dimension_display(self):
        if self.lumber_dim:
            return f"{self.lumber_dim.thickness}-{self.lumber_dim.width}-{self.lumber_dim.length} мм"
        elif self.unit_dim:
            return f"{self.unit_dim.length}-{self.unit_dim.width}-{self.unit_dim.height} мм"
        return "—"

    @property
    def volume_m3(self):
        if self.lumber_dim:
            return self.lumber_dim.volume_m3 * self.quantity
        return 0

    @property
    def area_m2(self):
        if self.lumber_dim:
            return self.lumber_dim.area_m2 * self.quantity
        return 0


# lumber_track/models.py

from django.db.models import Sum, Q
from datetime import date


def get_available_stocks(location_id=None, as_of_date=None):
    """
    Возвращает все позиции с актуальными остатками на складе
    """
    if as_of_date is None:
        as_of_date = date.today()

    # Получаем все уникальные комбинации из DocumentItem
    # с группировкой по полям
    from django.db import connection

    # Формируем запрос на получение позиций с остатками
    query = """
        SELECT 
            di.product_name_id,
            di.species_id, 
            di.grade_id,
            di.lumber_dim_id,
            di.unit_dim_id,
            COALESCE(SUM(CASE WHEN d.doc_type = 1 THEN di.quantity ELSE 0 END), 0) +
            COALESCE(SUM(CASE WHEN d.doc_type = 2 THEN di.quantity ELSE 0 END), 0) -
            COALESCE(SUM(CASE WHEN d.doc_type = 3 THEN di.quantity ELSE 0 END), 0) as balance
        FROM lumber_track_documentitem di
        JOIN lumber_track_document d ON di.document_id = d.id
        WHERE d.doc_date <= %s
    """

    params = [as_of_date]

    if location_id:
        query += " AND d.location_id = %s"
        params.append(location_id)

    query += """
        GROUP BY 
            di.product_name_id,
            di.species_id, 
            di.grade_id,
            di.lumber_dim_id,
            di.unit_dim_id
        HAVING balance > 0
        ORDER BY di.product_name_id
    """

    with connection.cursor() as cursor:
        cursor.execute(query, params)
        results = cursor.fetchall()

    return results