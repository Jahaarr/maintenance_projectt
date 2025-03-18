# maintenance_app/views.py
from django.shortcuts import render, redirect
from django.http import HttpResponse
import pandas as pd
from .forms import SousEnsembleForm
from .models import Equipement, SousEnsemble

def equipment_list(request):
    equipments = Equipement.objects.all()
    return render(request, 'maintenance_app/equipment_list.html', {'equipments': equipments})

def equipment_detail(request, equipement_id):
    equipement = Equipement.objects.get(id=equipement_id)
    sous_ensembles = SousEnsemble.objects.filter(equipement=equipement)
    return render(request, 'maintenance_app/equipment_detail.html', {'equipement': equipement, 'sous_ensembles': sous_ensembles})

def update_sous_ensemble(request, sous_ensemble_id):
    sous_ensemble = SousEnsemble.objects.get(id=sous_ensemble_id)
    if request.method == 'POST':
        form = SousEnsembleForm(request.POST, instance=sous_ensemble)
        if form.is_valid():
            form.save()
            return redirect('equipment_detail', equipement_id=sous_ensemble.equipement.id)
    else:
        form = SousEnsembleForm(instance=sous_ensemble)
    return render(request, 'maintenance_app/update_sous_ensemble.html', {'form': form, 'sous_ensemble': sous_ensemble})

def export_excel(request):
    sous_ensembles = SousEnsemble.objects.all()
    data = {
        'Equipement': [se.equipement.nom for se in sous_ensembles],
        'Sous-ensemble': [se.nom for se in sous_ensembles],
        'Quantité SE installée': [se.quantite_installee for se in sous_ensembles],
        'Relais disponible': [se.relais_disponible for se in sous_ensembles],
        'En attente': [se.en_attente_revision for se in sous_ensembles],
        'En cours': [se.encours_revision for se in sous_ensembles],
        'Corps disponible': [se.corps_disponible for se in sous_ensembles]
    }
    df = pd.DataFrame(data)
    response = HttpResponse(content_type='application/vnd.ms-excel')
    response['Content-Disposition'] = 'attachment; filename="sous_ensembles.xlsx"'
    df.to_excel(response, index=False)
    return response

def dashboard(request):
    """Affiche le tableau de bord avec des statistiques et des alertes."""
    # Statistiques globales
    total_equipments = Equipement.objects.count()
    sous_ensembles = SousEnsemble.objects.all()
    total_sous_ensembles = sous_ensembles.count()
    en_attente = sum(se.en_attente_revision for se in sous_ensembles)
    en_cours = sum(se.encours_revision for se in sous_ensembles)
    disponibles = sum(se.relais_disponible + se.corps_disponible for se in sous_ensembles)

    # Alertes : sous-ensembles avec relais = 0 et en attente ou en cours de révision
    alerts = sous_ensembles.filter(relais_disponible=0, en_attente_revision__gt=0)

    # Données pour les graphiques
    # Graphique 1 : Répartition des sous-ensembles par statut
    stats_by_status = {
        'En attente': en_attente,
        'En cours': en_cours,
        'Disponibles': disponibles,
    }

    # Graphique 2 : Nombre de sous-ensembles par équipement
    stats_by_equipment = {}
    for equip in Equipement.objects.all():
        count = SousEnsemble.objects.filter(equipement=equip).count()
        if count > 0:
            stats_by_equipment[equip.nom] = count

    context = {
        'total_equipments': total_equipments,
        'total_sous_ensembles': total_sous_ensembles,
        'en_attente': en_attente,
        'en_cours': en_cours,
        'disponibles': disponibles,
        'alerts': alerts,
        'stats_by_status': stats_by_status,
        'stats_by_equipment': stats_by_equipment,
    }
    return render(request, 'maintenance_app/dashboard.html', context)