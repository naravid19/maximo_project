# path : maximo_project\maximo_app\models.py

from django.db import models

# Create your models here.
#! Plant Type
class PlantType(models.Model):
    plant_code = models.CharField(max_length=8, unique=True, verbose_name="plant_code")  # รหัสโรงไฟฟ้า
    plant_type_th = models.CharField(max_length=100, verbose_name="plant_type_th")  # ประเภทโรงไฟฟ้า (TH)
    plant_type_eng = models.CharField(max_length=100, verbose_name="plant_type_eng")  # ประเภทโรงไฟฟ้า (ENG)
    act_types = models.ManyToManyField('ActType', related_name='plant_types')   # เชื่อมโยงกับ ActType ผ่าน ManyToManyField
    work_types = models.ManyToManyField('WorkType', related_name='plant_types')  # ความสัมพันธ์ Many-to-Many กับ WorkType
    units = models.ManyToManyField('Unit', related_name='plant_types')
    class Meta:
        verbose_name = "Plant Type"
        verbose_name_plural = "Plant Type"

    def __str__(self):
        return f"{self.plant_code}"

    def save(self, *args, **kwargs):
        self.plant_code = self.plant_code.upper()
        super().save(*args, **kwargs)

#! Site
class Site(models.Model):
    site_id = models.CharField(max_length=8, unique=True, verbose_name="site_id")  # SITEID
    site_name = models.CharField(max_length=100, verbose_name="site_name")  # คำอธิบาย
    organization = models.CharField(max_length=8, verbose_name="organization")  # องค์กร
    plant_types = models.ManyToManyField(PlantType, related_name='sites')   # เพิ่ม ManyToManyField เพื่อเชื่อมโยงกับ PlantType
    class Meta:
        verbose_name = "Site"
        verbose_name_plural = "Site"

    def __str__(self):
        return f"{self.site_id}"

    def save(self, *args, **kwargs):
        # แปลง plant_code ให้เป็นตัวพิมพ์ใหญ่
        self.site_id = self.site_id.upper()
        super().save(*args, **kwargs)

#! ChildSite
class ChildSite(models.Model):
    site_id = models.CharField(max_length=8, unique=True, verbose_name="site_id")
    site_name = models.CharField(max_length=100, verbose_name="site_name")
    parent_site = models.ForeignKey(Site, null=True, blank=True, on_delete=models.CASCADE, related_name='child_sites')  # เชื่อมโยงกับ Site หลัก
    class Meta:
        verbose_name = "Child Site"
        verbose_name_plural = "Child Site"

    def __str__(self):
        return f"{self.site_id}"

    def save(self, *args, **kwargs):
        # แปลง plant_code ให้เป็นตัวพิมพ์ใหญ่
        self.site_id = self.site_id.upper()
        super().save(*args, **kwargs)

#! Unit
class Unit(models.Model):
    unit_code = models.CharField(max_length=8, unique=True, verbose_name="unit_code")  # Block/Unit
    description = models.CharField(max_length=100, blank=True, null=True, verbose_name="description")  # คำอธิบาย
    class Meta:
        verbose_name = "Unit"
        verbose_name_plural = "Unit"

    def __str__(self):
        return f"{self.unit_code}"

#! WORKTYPE
class WorkType(models.Model):
    worktype = models.CharField(max_length=8, unique=True, verbose_name="worktype")  # WORKTYPE
    description = models.CharField(max_length=100, verbose_name="description")  # คำอธิบาย
    class Meta:
        verbose_name = "Work Type"
        verbose_name_plural = "Work Type"
        
    def __str__(self):
        return f"{self.worktype}"

    def save(self, *args, **kwargs):
        self.worktype = self.worktype.upper()
        super().save(*args, **kwargs)

#! ACTTYPE
class ActType(models.Model):
    acttype = models.CharField(max_length=8, unique=True, verbose_name="acttype")  # ACTTYPE
    description = models.CharField(max_length=100, verbose_name="description")  # คำอธิบาย
    code = models.CharField(max_length=8, verbose_name="code")  # CODE
    class Meta:
        verbose_name = "Act Type"
        verbose_name_plural = "Act Type"
    
    def __str__(self):
        return f"{self.acttype}"

    def save(self, *args, **kwargs):
        self.acttype = self.acttype.upper()
        super().save(*args, **kwargs)

#! WBS
class WBSCode(models.Model):
    wbs_code = models.CharField(max_length=8, unique=True, verbose_name="wbs_code")  # WBS Code
    description = models.CharField(max_length=100, verbose_name="description")  # คำอธิบาย
    class Meta:
        verbose_name = "WBS"
        verbose_name_plural = "WBS"
    def __str__(self):
        return f"{self.wbs_code}"
    
    def save(self, *args, **kwargs):
        self.wbs_code = self.wbs_code.upper()
        super().save(*args, **kwargs)

#! STATUS
class Status(models.Model):
    status = models.CharField(max_length=8, unique=True, verbose_name="status") # Status
    description = models.CharField(max_length=100, verbose_name="description")   # คำอธิบาย
    class Meta:
        verbose_name = "Status"
        verbose_name_plural = "Status"    
    
    def __str__(self):
        return f"{self.status}"
    
    def save(self, *args, **kwargs):
        self.status = self.status.upper()
        super().save(*args, **kwargs)