# path : maximo_project/maximo_app/urls.py

from django.urls import path
from . import views

# กำหนด custom error handlers
handler404 = 'maximo_app.views.custom_404'
handler500 = 'maximo_app.views.custom_500'

urlpatterns = [
    
    path('', views.index, name='index'),
    
    # URL สำหรับ AJAX
    path('filter-plant-type/', views.filter_plant_type, name='filter_plant_type'),
    path('filter-child-site/', views.filter_child_site, name='filter_child_site'),
    path('filter-site/', views.filter_site, name='filter_site'),
    path('filter-work-type/', views.filter_worktype, name='filter_worktype'),
    path('filter-act-type/', views.filter_acttype, name='filter_acttype'),
    path('filter-wbs/', views.filter_wbs, name='filter_wbs'),
    path('filter-wo-status/', views.filter_wostatus, name='filter_wostatus'),
    
    # URLs สำหรับดาวน์โหลดไฟล์
    path('download-user-schedule/', views.download_user_schedule, name='download_user_schedule'),
    path('download-user-location/', views.download_user_location, name='download_user_location'),
    path('download-comment/', views.download_comment_file, name='download_comment_file'),
    # path('download-job-plan-task/', views.download_job_plan_task_file, name='download_job_plan_task_file'),
    # path('download-job-plan-labor/', views.download_job_plan_labor_file, name='download_job_plan_labor_file'),
    # path('download-pm-plan/', views.download_pm_plan_file, name='download_pm_plan_file'),
    # path('download-template/', views.download_template_file, name='download_template_file'),
    path('download-template/<str:version>/', views.download_template_file, name='download_template_file'),
    path('download-original/', views.download_original_template, name='download_original_template'),
    path('download-schedule/', views.download_schedule, name='download_schedule'),
    path('download-example-schedule/', views.download_example_schedule, name='download_example_schedule'),
    path('download-example-template/', views.download_example_template, name='download_example_template'),
]