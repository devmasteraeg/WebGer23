# Generated by Django 4.2.5 on 2023-10-18 20:06

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('processo', '0005_alter_andamentoadm_tipo_andamento'),
    ]

    operations = [
        migrations.AddField(
            model_name='auditoria',
            name='objeto_descricao',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
    ]
