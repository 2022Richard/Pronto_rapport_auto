from django.urls import path
from . import views


app_name = 'rapport'

urlpatterns = [

    # Accueil
    path('', views.accueil, name='accueil'), 

    # Rapport
    path('rapport', views.rapport, name='rapport'), 

    # Contact 
    path('contact', views.contact, name='contact'), 

     ]