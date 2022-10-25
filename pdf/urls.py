from django.urls import path
from .views import *

app_name = 'pdf'

urlpatterns = [
    path('', home, name='home'),
]