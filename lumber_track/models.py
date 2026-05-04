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