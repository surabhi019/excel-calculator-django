from django.urls import path

from . import views

app_name = "myapp"

urlpatterns = [
    path('', views.index, name='index'),
    path('download_excel', views.download_excel, name='download_excel'),
]
