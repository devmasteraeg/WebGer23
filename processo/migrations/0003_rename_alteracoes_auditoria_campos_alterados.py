# Generated by Django 4.2.5 on 2023-10-09 13:46

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('processo', '0002_alter_auditoria_alteracoes'),
    ]

    operations = [
        migrations.RenameField(
            model_name='auditoria',
            old_name='alteracoes',
            new_name='campos_alterados',
        ),
    ]
