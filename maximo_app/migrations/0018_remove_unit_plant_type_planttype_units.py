# Generated by Django 5.1 on 2024-10-02 15:13

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('maximo_app', '0017_unit_plant_type'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='unit',
            name='plant_type',
        ),
        migrations.AddField(
            model_name='planttype',
            name='units',
            field=models.ManyToManyField(related_name='plant_types', to='maximo_app.unit'),
        ),
    ]