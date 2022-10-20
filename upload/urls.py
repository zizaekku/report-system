from django.urls import path
from django.urls import path
from django.conf import settings
from django.conf.urls.static import static

from .views import *

app_name = 'upload'

urlpatterns = [
    path('', upload, name='upload'),
]

if settings.DEBUG:
    urlpatterns += static(
        settings.MEDIA_URL,
        document_root=settings.MEDIA_ROOT
    )