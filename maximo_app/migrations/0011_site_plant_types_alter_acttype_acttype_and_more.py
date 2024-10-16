# Generated by Django 5.1 on 2024-10-01 02:56

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('maximo_app', '0010_status'),
    ]

    operations = [
        migrations.AddField(
            model_name='site',
            name='plant_types',
            field=models.ManyToManyField(related_name='sites', to='maximo_app.planttype'),
        ),
        migrations.AlterField(
            model_name='acttype',
            name='acttype',
            field=models.CharField(max_length=8, unique=True, verbose_name='acttype'),
        ),
        migrations.AlterField(
            model_name='acttype',
            name='code',
            field=models.CharField(max_length=8, verbose_name='code'),
        ),
        migrations.AlterField(
            model_name='acttype',
            name='description',
            field=models.CharField(max_length=100, verbose_name='description'),
        ),
        migrations.AlterField(
            model_name='acttype',
            name='remark',
            field=models.CharField(blank=True, max_length=100, null=True, verbose_name='remark (ไม่มีใน)'),
        ),
        migrations.AlterField(
            model_name='planttype',
            name='plant_code',
            field=models.CharField(max_length=8, unique=True, verbose_name='plant_code'),
        ),
        migrations.AlterField(
            model_name='planttype',
            name='plant_type_eng',
            field=models.CharField(max_length=100, verbose_name='plant_type_eng'),
        ),
        migrations.AlterField(
            model_name='planttype',
            name='plant_type_th',
            field=models.CharField(max_length=100, verbose_name='plant_type_th'),
        ),
        migrations.AlterField(
            model_name='site',
            name='organization',
            field=models.CharField(max_length=8, verbose_name='organization'),
        ),
        migrations.AlterField(
            model_name='site',
            name='site_id',
            field=models.CharField(max_length=8, unique=True, verbose_name='site_id'),
        ),
        migrations.AlterField(
            model_name='site',
            name='site_name',
            field=models.CharField(max_length=100, verbose_name='site_name'),
        ),
        migrations.AlterField(
            model_name='status',
            name='description',
            field=models.CharField(max_length=100, verbose_name='description'),
        ),
        migrations.AlterField(
            model_name='status',
            name='status',
            field=models.CharField(max_length=8, unique=True, verbose_name='status'),
        ),
        migrations.AlterField(
            model_name='unit',
            name='description',
            field=models.CharField(blank=True, max_length=100, null=True, verbose_name='description'),
        ),
        migrations.AlterField(
            model_name='unit',
            name='unit_code',
            field=models.CharField(max_length=8, unique=True, verbose_name='unit_code'),
        ),
        migrations.AlterField(
            model_name='wbscode',
            name='description',
            field=models.CharField(max_length=100, verbose_name='description'),
        ),
        migrations.AlterField(
            model_name='wbscode',
            name='wbs_code',
            field=models.CharField(max_length=8, unique=True, verbose_name='wbs_code'),
        ),
        migrations.AlterField(
            model_name='worktype',
            name='description',
            field=models.CharField(max_length=100, verbose_name='description'),
        ),
        migrations.AlterField(
            model_name='worktype',
            name='worktype',
            field=models.CharField(max_length=8, unique=True, verbose_name='worktype'),
        ),
    ]
