# Generated by Django 5.1 on 2024-09-20 07:00

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('maximo_app', '0003_alter_planttype_options'),
    ]

    operations = [
        migrations.AlterModelOptions(
            name='planttype',
            options={'verbose_name': 'ประเภทโรงไฟฟ้า', 'verbose_name_plural': 'Plant Type'},
        ),
    ]
