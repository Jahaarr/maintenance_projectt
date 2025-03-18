from django.db import models

class Equipement(models.Model):
    nom = models.CharField(max_length=100, unique=True)  # Ajout de unique=True    reference = models.CharField(max_length=50)
    reference = models.CharField(max_length=50)
    modele = models.CharField(max_length=50)
    localisation = models.CharField(max_length=100)
    date_ajout = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.nom

class SousEnsemble(models.Model):
    equipement = models.ForeignKey(Equipement, on_delete=models.CASCADE, related_name='sous_ensembles')
    nom = models.CharField(max_length=100)
    reference = models.CharField(max_length=50)
    modele = models.CharField(max_length=50)
    statut = models.CharField(max_length=20)
    date_ajout = models.DateTimeField(auto_now_add=True)
    quantite_installee = models.IntegerField(default=0)
    relais_disponible = models.IntegerField(default=0)
    en_attente_revision = models.IntegerField(default=0)
    encours_revision = models.IntegerField(default=0)
    corps_disponible = models.IntegerField(default=0)

    def __str__(self):
        return f"{self.nom} ({self.equipement.nom})"

class Maintenance(models.Model):
    sous_ensemble = models.ForeignKey(SousEnsemble, on_delete=models.CASCADE, related_name='maintenances')
    type = models.CharField(max_length=50)
    date_intervention = models.DateTimeField()
    duree = models.IntegerField()
    statut = models.CharField(max_length=20)
    rapport = models.TextField(blank=True)

class PlanMaintenance(models.Model):
    sous_ensemble = models.ForeignKey(SousEnsemble, on_delete=models.CASCADE, related_name='plans')
    type = models.CharField(max_length=50)
    date_prev = models.DateTimeField()
    statut = models.CharField(max_length=20)

class HSEIncident(models.Model):
    sous_ensemble = models.ForeignKey(SousEnsemble, on_delete=models.CASCADE, related_name='incidents')
    date_incident = models.DateTimeField()
    description = models.TextField()
    gravite = models.CharField(max_length=20)

class ConsommationPiece(models.Model):
    maintenance = models.ForeignKey(Maintenance, on_delete=models.CASCADE, related_name='consommations')
    quantite = models.IntegerField()

class PieceRechange(models.Model):
    consommation = models.ForeignKey(ConsommationPiece, on_delete=models.CASCADE, related_name='pieces')
    nom = models.CharField(max_length=100)
    reference = models.CharField(max_length=50)
    stock_actuel = models.IntegerField()