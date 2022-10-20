from django.db import models

# Create your models here.
class Mapping(models.Model):
    col = models.CharField(max_length=2)
    row = models.IntegerField()

    def __str__(self):
        return self.col + str(self.row)