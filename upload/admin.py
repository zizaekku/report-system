from django.contrib import admin

# Register your models here.
from upload.models import *

class DataAdmin(admin.ModelAdmin):
    list_display = [f.name for f in Data._meta.fields]

class ImageAdmin(admin.ModelAdmin):
    list_display = [f.name for f in Image._meta.fields]


admin.site.register(Data, DataAdmin)
admin.site.register(Image, ImageAdmin)