from django.db.models.signals import pre_save
from django.dispatch import receiver
from django.contrib.auth.models import User
from .models import Aide_Engin
""""
@receiver(pre_save, sender=Aide_Engin)
def pre_save_handler(sender, instance, **kwargs):
    print("Pre-save signal received. About to save:",)
"""
