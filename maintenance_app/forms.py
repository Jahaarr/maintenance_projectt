# maintenance_app/forms.py
from django import forms
from .models import SousEnsemble

class SousEnsembleForm(forms.ModelForm):
    class Meta:
        model = SousEnsemble
        fields = ['quantite_installee', 'relais_disponible', 'en_attente_revision', 'encours_revision', 'corps_disponible']