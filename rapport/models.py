from django.db import models

# Create your models here.
   
class Dispatch_Engin(models.Model):
    Vehicule = models.CharField(max_length=30, )
    Area = models.CharField(max_length=30,)
    Description = models.CharField(max_length=200,)

    def __str__(self):
        return self.Description 
      
    
    
