from django.contrib import admin
from maximo_app.models import PlantType, Site, ChildSite, Unit, WorkType, ActType, WBSCode, Status
# Register your models here.

@admin.register(PlantType)
class PlantTypeAdmin(admin.ModelAdmin):
    list_display = ('plant_code', 'plant_type_th', 'plant_type_eng', 'get_act_types', 'get_work_types', 'get_units')
    search_fields = ('plant_code', 'plant_type_th', 'plant_type_eng', 'act_types__acttype', 'work_types__worktype', 'units__unit_code')
    list_filter = ('act_types__acttype', 'work_types__worktype', 'units__unit_code')
    fields = ('plant_code', 'plant_type_th', 'plant_type_eng', 'act_types', 'work_types', 'units')
    
    def get_act_types(self, obj):
        act_types = obj.act_types.all()
        if act_types.exists():
            return ", ".join([acttype.acttype for acttype in act_types])
        return ""
    get_act_types.short_description = 'Acttype'
    
    def get_work_types(self, obj):
        work_types = obj.work_types.all()
        if work_types.exists():
            return ", ".join([worktype.worktype for worktype in work_types])
        return ""
    get_work_types.short_description = 'Work Types'

    def get_units(self, obj):
        units = obj.units.all()
        if units.exists():
            return ", ".join([unit.unit_code for unit in units])
        return ""
    get_units.short_description = 'Units'

    def has_add_permission(self, request):
        return request.user.is_superuser
    
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

@admin.register(Site)
class SiteAdmin(admin.ModelAdmin):
    list_display = ('site_id', 'site_name', 'organization', 'get_plant_types')  # ฟิลด์ที่จะแสดงในหน้า Admin
    search_fields = ('site_id', 'site_name', 'organization', 'plant_types__plant_code')  # ฟิลด์ที่สามารถค้นหาได้ใน Admin
    list_filter = ('plant_types__plant_code',)  # ฟิลด์ที่ใช้สำหรับการกรองข้อมูล
    fields = ('site_id', 'site_name', 'organization', 'plant_types')  # กำหนดฟิลด์ที่จะแสดงในฟอร์มเพิ่มหรือแก้ไขข้อมูล
    
    def get_plant_types(self, obj):
        plant_types = obj.plant_types.all()
        if plant_types.exists():
            return ", ".join([plant_type.plant_code  for plant_type in plant_types])
        return ""
    get_plant_types.short_description = 'Plant Types'
    
    def has_add_permission(self, request):
        return request.user.is_superuser  # เฉพาะ superuser ที่สามารถเพิ่มข้อมูลได้
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser  # เฉพาะ superuser ที่สามารถแก้ไขข้อมูลได้

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser  # เฉพาะ superuser ที่สามารถลบข้อมูลได้

@admin.register(ChildSite)
class ChildSiteAdmin(admin.ModelAdmin):
    list_display = ('site_id', 'site_name', 'get_parent_sites')
    search_fields = ('site_id', 'site_name', 'parent_site__site_id')
    list_filter = ('parent_site__site_id',)
    fields = ('site_id', 'site_name', 'parent_site')
    
    def get_parent_sites(self, obj):
        return obj.parent_site.site_id if obj.parent_site else ''
    get_parent_sites.short_description = 'Parent Site'
    
    def has_add_permission(self, request):
        return request.user.is_superuser 
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

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
    list_filter = ('description',)
    fields = ('worktype', 'description')
    
    def has_add_permission(self, request):
        return request.user.is_superuser
    
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

@admin.register(ActType)
class ActTypeAdmin(admin.ModelAdmin):
    list_display = ('acttype', 'description', 'code')
    search_fields = ('acttype', 'description', 'code')
    list_filter = ('description',)
    fields = ('acttype', 'description', 'code',)
    
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
    list_filter = ('description', )
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
    list_filter = ('description', )
    fields = ('status', 'description')
    
    def has_add_permission(self, request):
        return request.user.is_superuser
    
    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser