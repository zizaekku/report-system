from django.contrib import admin
from .models import *

# Register your models here.
class MappingAdmin(admin.ModelAdmin):
    list_display = [f.name for f in Mapping._meta.fields]

admin.site.register(Mapping, MappingAdmin)