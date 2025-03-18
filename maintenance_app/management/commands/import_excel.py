# maintenance_app/management/commands/import_excel.py
import pandas as pd
from django.core.management.base import BaseCommand
from maintenance_app.models import Equipement, SousEnsemble

class Command(BaseCommand):
    help = 'Importe les données des équipements et sous-ensembles depuis un fichier Excel'

    def add_arguments(self, parser):
        parser.add_argument('file_path', type=str, help='Chemin du fichier Excel à importer')

    def handle(self, *args, **options):
        file_path = options['file_path']
        try:
            # Lit le fichier Excel
            df = pd.read_excel(file_path)

            # Parcourt chaque ligne du fichier
            for index, row in df.iterrows():
                # Crée ou récupère l'équipement
                equipement, created = Equipement.objects.get_or_create(
                    nom=row['Equipement'],
                    defaults={
                        'reference': row.get('Référence', ''),
                        'modele': row.get('Modèle', ''),
                        'localisation': row.get('Localisation', '')
                    }
                )

                # Crée ou met à jour le sous-ensemble
                sous_ensemble, created = SousEnsemble.objects.get_or_create(
                    equipement=equipement,
                    nom=row['Sous-ensemble'],
                    defaults={
                        'reference': row.get('Référence SE', ''),
                        'modele': row.get('Modèle SE', ''),
                        'statut': row.get('Statut', ''),
                        'quantite_installee': row.get('Quantité SE installée', 0),
                        'relais_disponible': row.get('Sous-ensemble relais disponible (révisé)', 0),
                        'en_attente_revision': row.get('Sous-ensemble en attente révision', 0),
                        'encours_revision': row.get('Sous-ensemble encours de révision', 0),
                        'corps_disponible': row.get('Corps de Sous-ensembles disponibles (révisable)', 0)
                    }
                )
                self.stdout.write(self.style.SUCCESS(f'Importé : {equipement.nom} - {sous_ensemble.nom}'))

            self.stdout.write(self.style.SUCCESS('Importation terminée avec succès'))
        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Erreur lors de l\'importation : {str(e)}'))