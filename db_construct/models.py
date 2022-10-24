from django.db import models

# Create your models here.
class Class(models.Model):
    class_key = models.CharField(max_length=5, primary_key=True)
    depth = models.IntegerField()
    ancestor = models.CharField(max_length=5)
    name = models.CharField(max_length=100)