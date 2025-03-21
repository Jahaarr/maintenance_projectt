# Generated by Django 5.1.7 on 2025-03-18 04:15

import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Equipement',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nom', models.CharField(max_length=100)),
                ('reference', models.CharField(max_length=50)),
                ('modele', models.CharField(max_length=50)),
                ('localisation', models.CharField(max_length=100)),
                ('date_ajout', models.DateTimeField(auto_now_add=True)),
            ],
        ),
        migrations.CreateModel(
            name='Maintenance',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('type', models.CharField(max_length=50)),
                ('date_intervention', models.DateTimeField()),
                ('duree', models.IntegerField()),
                ('statut', models.CharField(max_length=20)),
                ('rapport', models.TextField(blank=True)),
            ],
        ),
        migrations.CreateModel(
            name='ConsommationPiece',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('quantite', models.IntegerField()),
                ('maintenance', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='consommations', to='maintenance_app.maintenance')),
            ],
        ),
        migrations.CreateModel(
            name='PieceRechange',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nom', models.CharField(max_length=100)),
                ('reference', models.CharField(max_length=50)),
                ('stock_actuel', models.IntegerField()),
                ('consommation', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='pieces', to='maintenance_app.consommationpiece')),
            ],
        ),
        migrations.CreateModel(
            name='SousEnsemble',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nom', models.CharField(max_length=100)),
                ('reference', models.CharField(max_length=50)),
                ('modele', models.CharField(max_length=50)),
                ('statut', models.CharField(max_length=20)),
                ('date_ajout', models.DateTimeField(auto_now_add=True)),
                ('quantite_installee', models.IntegerField(default=0)),
                ('relais_disponible', models.IntegerField(default=0)),
                ('en_attente_revision', models.IntegerField(default=0)),
                ('encours_revision', models.IntegerField(default=0)),
                ('corps_disponible', models.IntegerField(default=0)),
                ('equipement', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='sous_ensembles', to='maintenance_app.equipement')),
            ],
        ),
        migrations.CreateModel(
            name='PlanMaintenance',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('type', models.CharField(max_length=50)),
                ('date_prev', models.DateTimeField()),
                ('statut', models.CharField(max_length=20)),
                ('sous_ensemble', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='plans', to='maintenance_app.sousensemble')),
            ],
        ),
        migrations.AddField(
            model_name='maintenance',
            name='sous_ensemble',
            field=models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='maintenances', to='maintenance_app.sousensemble'),
        ),
        migrations.CreateModel(
            name='HSEIncident',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('date_incident', models.DateTimeField()),
                ('description', models.TextField()),
                ('gravite', models.CharField(max_length=20)),
                ('sous_ensemble', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='incidents', to='maintenance_app.sousensemble')),
            ],
        ),
    ]
