from django.contrib import admin

# Register your models here.
from upload.models import Data

class DataAdmin(admin.ModelAdmin):
    list_display = [f.name for f in Data._meta.fields]


admin.site.register(Data, DataAdmin)