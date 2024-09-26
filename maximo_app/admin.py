from django.contrib import admin
from maximo_app.models import Site, PlantType, Unit, WorkType, ActType, WBSCode, Status
# Register your models here.

@admin.register(Site)
class SiteAdmin(admin.ModelAdmin):
    list_display = ('id', 'site_id', 'site_name', 'organization')  # ฟิลด์ที่จะแสดงในหน้า Admin
    search_fields = ('id', 'site_id', 'site_name', 'organization')  # ฟิลด์ที่สามารถค้นหาได้ใน Admin
    list_filter = ('site_id', 'site_name', 'organization')  # ฟิลด์ที่ใช้สำหรับการกรองข้อมูล
    fields = ('site_id', 'site_name', 'organization')  # กำหนดฟิลด์ที่จะแสดงในฟอร์มเพิ่มหรือแก้ไขข้อมูล
    
    def has_add_permission(self, request):
        return request.user.is_superuser  # เฉพาะ superuser ที่สามารถเพิ่มข้อมูลได้

@admin.register(PlantType)
class PlantTypeAdmin(admin.ModelAdmin):
    list_display = ('id', 'plant_code', 'plant_type_th', 'plant_type_eng')
    search_fields = ('id', 'plant_code', 'plant_type_th', 'plant_type_eng')
    list_filter = ('plant_code', 'plant_type_th', 'plant_type_eng')
    fields = ('plant_code', 'plant_type_th', 'plant_type_eng')
    
    def has_add_permission(self, request):
        return request.user.is_superuser

@admin.register(Unit)
class UnitAdmin(admin.ModelAdmin):
    list_display = ('id', 'unit_code', 'description')
    search_fields = ('id', 'unit_code', 'description')
    fields = ('unit_code', 'description')


@admin.register(WorkType)
class WorkTypeAdmin(admin.ModelAdmin):
    list_display = ('id', 'worktype', 'description')
    search_fields = ('id', 'worktype', 'description')
    fields = ('worktype', 'description')
    
    def has_add_permission(self, request):
        return request.user.is_superuser

@admin.register(ActType)
class ActTypeAdmin(admin.ModelAdmin):
    list_display = ('id', 'acttype', 'description', 'code', 'remark')
    search_fields = ('id', 'acttype', 'description', 'code')
    list_filter = ('acttype',)
    fields = ('acttype', 'description', 'code', 'remark')
    
    def has_add_permission(self, request):
        return request.user.is_superuser

@admin.register(WBSCode)
class WBSCodeAdmin(admin.ModelAdmin):
    list_display = ('id', 'wbs_code', 'description')
    search_fields = ('id', 'wbs_code', 'description')
    list_filter = ('wbs_code',)
    fields = ('wbs_code', 'description')
    
    def has_add_permission(self, request):
        return request.user.is_superuser

@admin.register(Status)
class StatusAdmin(admin.ModelAdmin):
    list_display = ('id', 'status', 'description')
    search_fields = ('id', 'status', 'description')
    list_filter = ('status',)
    fields = ('status', 'description')
    
    def has_add_permission(self, request):
        return request.user.is_superuser