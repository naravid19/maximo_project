from django.contrib import admin
from maximo_app.models import Site, PlantType, Unit, WorkType, ActType, WBSCode, Status
# Register your models here.

@admin.register(PlantType)
class PlantTypeAdmin(admin.ModelAdmin):
    list_display = ('plant_code', 'plant_type_th', 'plant_type_eng')
    search_fields = ('plant_code', 'plant_type_th', 'plant_type_eng')
    list_filter = ('plant_code', 'plant_type_th', 'plant_type_eng')
    fields = ('plant_code', 'plant_type_th', 'plant_type_eng')
    
    def has_add_permission(self, request):
        return request.user.is_superuser
    
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

@admin.register(Site)
class SiteAdmin(admin.ModelAdmin):
    list_display = ('site_id', 'site_name', 'organization', 'get_plant_types')  # ฟิลด์ที่จะแสดงในหน้า Admin
    search_fields = ('site_id', 'site_name', 'organization')  # ฟิลด์ที่สามารถค้นหาได้ใน Admin
    list_filter = ('site_id', 'site_name', 'organization')  # ฟิลด์ที่ใช้สำหรับการกรองข้อมูล
    fields = ('site_id', 'site_name', 'organization')  # กำหนดฟิลด์ที่จะแสดงในฟอร์มเพิ่มหรือแก้ไขข้อมูล
    
    def get_plant_types(self, obj):
        return ", ".join([plant_type.plant_code for plant_type in obj.plant_types.all()])
    get_plant_types.short_description = 'Plant Types'
    
    def has_add_permission(self, request):
        return request.user.is_superuser  # เฉพาะ superuser ที่สามารถเพิ่มข้อมูลได้
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser  # เฉพาะ superuser ที่สามารถแก้ไขข้อมูลได้

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser  # เฉพาะ superuser ที่สามารถลบข้อมูลได้

@admin.register(Unit)
class UnitAdmin(admin.ModelAdmin):
    list_display = ('unit_code', 'description')
    search_fields = ('unit_code', 'description')
    fields = ('unit_code', 'description')
    
    def has_add_permission(self, request):
        return request.user.is_superuser
    
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

@admin.register(WorkType)
class WorkTypeAdmin(admin.ModelAdmin):
    list_display = ('worktype', 'description')
    search_fields = ('worktype', 'description')
    fields = ('worktype', 'description')
    
    def has_add_permission(self, request):
        return request.user.is_superuser
    
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

@admin.register(ActType)
class ActTypeAdmin(admin.ModelAdmin):
    list_display = ('acttype', 'description', 'code', 'remark')
    search_fields = ('acttype', 'description', 'code')
    list_filter = ('acttype',)
    fields = ('acttype', 'description', 'code', 'remark')
    
    def has_add_permission(self, request):
        return request.user.is_superuser
    
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

@admin.register(WBSCode)
class WBSCodeAdmin(admin.ModelAdmin):
    list_display = ('wbs_code', 'description')
    search_fields = ('wbs_code', 'description')
    list_filter = ('wbs_code',)
    fields = ('wbs_code', 'description')
    
    def has_add_permission(self, request):
        return request.user.is_superuser
    
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

@admin.register(Status)
class StatusAdmin(admin.ModelAdmin):
    list_display = ('status', 'description')
    search_fields = ('status', 'description')
    list_filter = ('status',)
    fields = ('status', 'description')
    
    def has_add_permission(self, request):
        return request.user.is_superuser
    
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser