# Generated by Django 5.1 on 2024-10-02 14:45

import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('maximo_app', '0016_remove_acttype_remark'),
    ]

    operations = [
        migrations.AddField(
            model_name='unit',
            name='plant_type',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='units', to='maximo_app.planttype'),
        ),
    ]
