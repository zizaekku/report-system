# Generated by Django 4.1.2 on 2022-10-26 01:51

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('upload', '0002_image'),
    ]

    operations = [
        migrations.AlterField(
            model_name='image',
            name='image',
            field=models.TextField(),
        ),
    ]
