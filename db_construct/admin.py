from django.contrib import admin

# Register your models here.
from .models import Class

class ClassAdmin(admin.ModelAdmin):
    list_display = [f.name for f in Class._meta.fields]


admin.site.register(Class, ClassAdmin)