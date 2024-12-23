# Generated by Django 5.1 on 2024-09-20 08:40

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('maximo_app', '0006_unit_alter_site_site_id'),
    ]

    operations = [
        migrations.CreateModel(
            name='WorkType',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('worktype', models.CharField(max_length=8, unique=True, verbose_name='WORKTYPE')),
                ('description', models.CharField(max_length=100, verbose_name='คำอธิบาย')),
            ],
            options={
                'verbose_name': 'WORKTYPE',
                'verbose_name_plural': 'WORKTYPE',
            },
        ),
    ]
