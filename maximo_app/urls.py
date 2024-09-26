# path : maximo_project/maximo_app/urls.py

from django.urls import path
from . import views

# กำหนด custom error handlers
handler404 = 'maximo_app.views.custom_404'
handler500 = 'maximo_app.views.custom_500'

urlpatterns = [
    path('', views.index, name='index'),
    path('test/', views.test, name='test'),
    path('download_comment_file/', views.download_comment_file, name='download_comment_file'),
    path('download_job_plan_task_file/', views.download_job_plan_task_file, name='download_job_plan_task_file'),
    path('download_job_plan_labor_file/', views.download_job_plan_labor_file, name='download_job_plan_labor_file'),
    path('download_pm_plan_file/', views.download_pm_plan_file, name='download_pm_plan_file'),
    path('download_template_file/', views.download_template_file, name='download_template_file'),
    path('filter_plant_type/', views.filter_plant_type, name='filter_plant_type'),
    path('filter_site/', views.filter_site, name='filter_site'),
    path('filter_worktype/', views.filter_worktype, name='filter_worktype'),
    path('filter_acttype/', views.filter_acttype, name='filter_acttype'),
    path('filter_wbs/', views.filter_wbs, name='filter_wbs'),
    path('filter_wostatus/', views.filter_wostatus, name='filter_wostatus'),
    # path('details/', views.details, name='details'),  # แสดงหน้ารายละเอียด
]