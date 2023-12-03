from django.contrib import admin

# Register your models here.

from .models import Dispatch_Engin


class AdminDispatch_Engin (admin.ModelAdmin):
    list_display = ('Vehicule', 'Area', 'Description')  

from import_export.admin import ImportExportModelAdmin 

@admin.register(Dispatch_Engin)
class ViewAdmin (ImportExportModelAdmin) : 
    pass


