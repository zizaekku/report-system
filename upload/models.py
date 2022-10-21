from django.db import models

# Create your models here.
class Data(models.Model):
    data = models.TextField()

class Image(models.Model):
    image = models.ImageField()