from django.urls import path
from .views import *

app_name = 'pdf'

urlpatterns = [
    path('', create_report, name='create'),
]