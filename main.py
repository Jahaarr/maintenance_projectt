import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QComboBox, QPushButton, QLineEdit, QFrame, QMessageBox,
                             QTableWidget, QTableWidgetItem, QTabWidget, QScrollArea)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
import pandas as pd
import numpy as np

class EquipmentApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Equipment Management App")
        self.setGeometry(100, 100, 1000, 700)

        # Apply a professional light theme to the main window
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f6fa;
            }
            QTabWidget::pane {
                border: 1px solid #dfe6e9;
                background-color: #ffffff;
            }
            QTabBar::tab {
                background: #dfe6e9;
                color: #2d3436;
                padding: 10px 20px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                font: 12pt "Segoe UI";
            }
            QTabBar::tab:selected {
                background: #ffffff;
                border-bottom: 2px solid #0984e3;
            }
            QTabBar::tab:hover {
                background: #b2bec3;
            }
        """)

        # Central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Load the Excel file with multiple sheets
        try:
            self.excel_data = pd.read_excel("ENGINS.xlsx", sheet_name=None, header=None)
        except FileNotFoundError:
            self.excel_data = {}
            QMessageBox.critical(self, "Error", "Excel file 'ENGINS.xlsx' not found.")
            return

        # Define expected headers for each sheet type
        self.sheet_configs = {
            "Park engin": {
                "headers": ["Equipement", "MLE", "DMS", "TYPE", "N° DES SERIES", "SITUATION"],
                "numeric_cols": [],
                "filter_col": "SITUATION"
            },
            "Cartographie moteur": {
                "headers": ["Equipement", "Sous-ensemble", "Criticité", "Quantité SE installée",
                            "Sous-ensemble relais disponible (révisé)", "Sous-ensemble en attente révision",
                            "Sous-ensemble encours de révision", "Corps de Sous-ensembles disponibles (révisable)"],
                "numeric_cols": ["Quantité SE installée", "Sous-ensemble relais disponible (révisé)",
                                 "Sous-ensemble en attente révision", "Sous-ensemble encours de révision",
                                 "Corps de Sous-ensembles disponibles (révisable)"],
                "filter_col": "Criticité"
            },
            "Cartographie transmission": {
                "headers": ["Equipement", "Sous-ensemble", "Criticité", "Quantité SE installée",
                            "Sous-ensemble relais disponible (révisé)", "Sous-ensemble en attente révision",
                            "Sous-ensemble encours de révision", "Corps de Sous-ensembles disponibles (révisable)"],
                "numeric_cols": ["Quantité SE installée", "Sous-ensemble relais disponible (révisé)",
                                 "Sous-ensemble en attente révision", "Sous-ensemble encours de révision",
                                 "Corps de Sous-ensembles disponibles (révisable)"],
                "filter_col": "Criticité"
            },
            "Cartographie Engin": {
                "headers": ["Equipement", "Sous-ensemble", "Criticité", "Quantité SE installée",
                            "Sous-ensemble relais disponible (révisé)", "Sous-ensemble en attente révision",
                            "Sous-ensemble encours de révision", "Corps de Sous-ensembles disponibles (révisable)"],
                "numeric_cols": ["Quantité SE installée", "Sous-ensemble relais disponible (révisé)",
                                 "Sous-ensemble en attente révision", "Sous-ensemble encours de révision",
                                 "Corps de Sous-ensembles disponibles (révisable)"],
                "filter_col": "Criticité"
            },
            "Performances BG": {
                "headers": ["équipement", "Sous-ensemble", "date de changement 1", "OT", "Compteur de changement 1",
                            "date de changement 2", "OT", "Compteur de changement 2", "date de changement 3", "OT",
                            "Compteur de changement 3", "date de changement 4", "OT", "Compteur de changement 4",
                            "date de changement 5", "OT", "Compteur de changement 5", "date de changement 6", "OT",
                            "Compteur de changement 6", "compteur actuel S45/2024", "PERFORMANCE"],
                "numeric_cols": ["Compteur de changement 1", "Compteur de changement 2", "Compteur de changement 3",
                                 "Compteur de changement 4", "Compteur de changement 5", "Compteur de changement 6",
                                 "compteur actuel S45/2024", "PERFORMANCE"],
                "filter_col": None
            },
            "Performances YSF": {
                "headers": ["équipement", "Sous-ensemble", "date de changement 1", "OT", "Compteur de changement 1",
                            "date de changement 2", "OT", "Compteur de changement 2", "date de changement 3", "OT",
                            "Compteur de changement 3", "date de changement 4", "OT", "Compteur de changement 4",
                            "date de changement 5", "OT", "Compteur de changement 5", "date de changement 6", "OT",
                            "Compteur de changement 6", "compteur actuel S45/2024", "PERFORMANCE"],
                "numeric_cols": ["Compteur de changement 1", "Compteur de changement 2", "Compteur de changement 3",
                                 "Compteur de changement 4", "Compteur de changement 5", "Compteur de changement 6",
                                 "compteur actuel S45/2024", "PERFORMANCE"],
                "filter_col": None
            },
            "Programme 2025 BG": {
                "headers": ["Type d'engin", "Equipement", "Sous-ensemble", "Qte v1", "Qte v2", "Qte v3",
                            "Devis unitaire", "Cout V2", "Cout V3", "Commentaire", "SECTION AFFECTATION"],
                "numeric_cols": ["Qte v1", "Qte v2", "Qte v3", "Devis unitaire", "Cout V2", "Cout V3"],
                "filter_col": "SECTION AFFECTATION"
            },
            "Programme 2025 YSF": {
                "headers": ["Equipement", "Engin", "REP", "Sous ensemble", "Seuil HM", "HM cumulés",
                            "Devis unitaire", "Qte [V1]", "Cout V1", "Qte [V2]", "Cout [V2]", "OBS",
                            "SECTION AFFECTATION"],
                "numeric_cols": ["Seuil HM", "HM cumulés", "Devis unitaire", "Qte [V1]", "Cout V1", "Qte [V2]", "Cout [V2]"],
                "filter_col": "SECTION AFFECTATION"
            }
        }

        # Preprocess sheets to extract relevant sections
        self.processed_data = {}
        for sheet_name, df in self.excel_data.items():
            self.processed_data[sheet_name] = self.preprocess_sheet(df, sheet_name)

        # Current sheet data
        self.df = None
        self.current_sheet_config = None

        # Title
        title_label = QLabel("Equipment Data Management")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont("Segoe UI", 18, QFont.Weight.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2d3436; margin-bottom: 10px;")
        main_layout.addWidget(title_label)

        # Tabs for different sections
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # Tab 1: Equipment Overview
        self.tab_equipment = QWidget()
        self.tabs.addTab(self.tab_equipment, "Equipment Overview")
        self.setup_equipment_tab()

        # Tab 2: Update Data
        self.tab_update = QWidget()
        self.tabs.addTab(self.tab_update, "Update Data")
        self.setup_update_tab()

        # Tab 3: Dashboard
        self.tab_dashboard = QWidget()
        self.tabs.addTab(self.tab_dashboard, "Dashboard")
        self.setup_dashboard_tab()

    def preprocess_sheet(self, df, sheet_name):
        if df.empty:
            print(f"Sheet {sheet_name} is empty. Skipping preprocessing.")
            return pd.DataFrame()

        # Check if the sheet is in the configuration
        if sheet_name not in self.sheet_configs:
            print(f"Unknown sheet: {sheet_name}. Skipping preprocessing.")
            return pd.DataFrame()

        config = self.sheet_configs[sheet_name]
        target_headers = config["headers"]

        # Find the header row by checking for a row that contains all target headers
        header_row = None
        for i in range(min(10, len(df))):  # Check up to the first 10 rows to skip metadata
            row = df.iloc[i].astype(str).str.strip()
            # Clean the row values to handle special characters and newlines
            row = row.str.replace(r'\n', ' ').str.replace(r'\s+', ' ', regex=True)
            if all(any(col in val for val in row.values) for col in target_headers):
                header_row = i
                break

        if header_row is not None:
            # Read the sheet again, starting from the header row
            df_section = pd.read_excel("ENGINS.xlsx", sheet_name=sheet_name, skiprows=header_row)
            # Clean column names
            df_section.columns = df_section.columns.str.strip().str.replace(r'\n', ' ').str.replace(r'\s+', ' ', regex=True)
            # Ensure all expected headers are present, fill missing ones with NaN
            for col in target_headers:
                if col not in df_section.columns:
                    df_section[col] = pd.NA
            # Reorder columns to match target_headers
            df_section = df_section[target_headers]
            # Drop rows where the first column is NaN
            df_section = df_section.dropna(subset=[target_headers[0]], how="all")
            # Convert numeric columns to appropriate types
            for col in config["numeric_cols"]:
                if col in df_section.columns:
                    df_section[col] = pd.to_numeric(df_section[col], errors='coerce').fillna(0)

            # Add a Section column for BG and YSF in specific sheets
            if sheet_name in ["Cartographie moteur", "Cartographie transmission", "Cartographie Engin"]:
                df_section['Section'] = pd.NA
                current_section = None
                for idx, row in df_section.iterrows():
                    if row['Equipement'] in ['BG', 'YSF']:
                        current_section = row['Equipement']
                    else:
                        df_section.at[idx, 'Section'] = current_section
                # Drop rows that are section headers (BG or YSF)
                df_section = df_section[~df_section['Equipement'].isin(['BG', 'YSF'])]

            return df_section
        else:
            print(f"Could not find the expected section in sheet {sheet_name}.")
            return pd.DataFrame()

    # Tab 1: Equipment Overview
    def setup_equipment_tab(self):
        layout = QVBoxLayout(self.tab_equipment)
        layout.setSpacing(10)

        # Sheet selection
        sheet_frame = QFrame()
        sheet_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        sheet_layout = QHBoxLayout(sheet_frame)
        sheet_label = QLabel("Select Sheet:")
        sheet_label.setFont(QFont("Segoe UI", 12))
        sheet_label.setStyleSheet("color: #2d3436;")
        sheet_layout.addWidget(sheet_label)

        self.sheet_combo_equipment = QComboBox()
        self.sheet_combo_equipment.setFont(QFont("Segoe UI", 12))
        self.sheet_combo_equipment.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                padding: 5px;
                border-radius: 3px;
                color: #2d3436;
            }
            QComboBox::drop-down {
                border-left: 1px solid #dfe6e9;
                padding-right: 5px;
            }
            QComboBox:hover {
                border: 1px solid #0984e3;
            }
        """)
        self.sheet_combo_equipment.addItems(list(self.processed_data.keys()))
        self.sheet_combo_equipment.currentTextChanged.connect(self.load_sheet_equipment)
        sheet_layout.addWidget(self.sheet_combo_equipment)
        layout.addWidget(sheet_frame)

        # Section filter frame (for BG/YSF)
        section_frame = QFrame()
        section_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        section_layout = QHBoxLayout(section_frame)
        section_label = QLabel("Filter by Section:")
        section_label.setFont(QFont("Segoe UI", 12))
        section_label.setStyleSheet("color: #2d3436;")
        section_layout.addWidget(section_label)

        self.section_combo = QComboBox()
        self.section_combo.setFont(QFont("Segoe UI", 12))
        self.section_combo.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                padding: 5px;
                border-radius: 3px;
                color: #2d3436;
            }
            QComboBox::drop-down {
                border-left: 1px solid #dfe6e9;
                padding-right: 5px;
            }
            QComboBox:hover {
                border: 1px solid #0984e3;
            }
        """)
        self.section_combo.addItem("All")
        self.section_combo.currentTextChanged.connect(self.update_equipment_table)
        section_layout.addWidget(self.section_combo)
        layout.addWidget(section_frame)

        # Filter frame
        filter_frame = QFrame()
        filter_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        filter_layout = QHBoxLayout(filter_frame)
        filter_label = QLabel("Filter by:")
        filter_label.setFont(QFont("Segoe UI", 12))
        filter_label.setStyleSheet("color: #2d3436;")
        filter_layout.addWidget(filter_label)

        self.filter_combo = QComboBox()
        self.filter_combo.setFont(QFont("Segoe UI", 12))
        self.filter_combo.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                padding: 5px;
                border-radius: 3px;
                color: #2d3436;
            }
            QComboBox::drop-down {
                border-left: 1px solid #dfe6e9;
                padding-right: 5px;
            }
            QComboBox:hover {
                border: 1px solid #0984e3;
            }
        """)
        self.filter_combo.addItem("All")
        self.filter_combo.currentTextChanged.connect(self.update_equipment_table)
        filter_layout.addWidget(self.filter_combo)
        layout.addWidget(filter_frame)

        # Table to display equipment data
        self.equipment_table = QTableWidget()
        self.equipment_table.setStyleSheet("""
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                gridline-color: #dfe6e9;
                color: #2d3436;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QTableWidget::item:alternate {
                background-color: #f5f6fa;
            }
            QHeaderView::section {
                background-color: #dfe6e9;
                color: #2d3436;
                padding: 5px;
                border: 1px solid #dfe6e9;
                font: bold 12pt "Segoe UI";
            }
        """)
        self.equipment_table.setAlternatingRowColors(True)
        layout.addWidget(self.equipment_table)

        # Load initial sheet
        if self.sheet_combo_equipment.currentText():
            self.load_sheet_equipment(self.sheet_combo_equipment.currentText())

    def load_sheet_equipment(self, sheet_name):
        if sheet_name:
            self.df = self.processed_data[sheet_name]
            self.current_sheet_config = self.sheet_configs.get(sheet_name, {})
            if not self.df.empty:
                # Update section dropdown for Cartographie sheets
                if sheet_name in ["Cartographie moteur", "Cartographie transmission", "Cartographie Engin"] and 'Section' in self.df.columns:
                    section_values = sorted(self.df['Section'].dropna().unique())
                    self.section_combo.clear()
                    self.section_combo.addItem("All")
                    self.section_combo.addItems([str(val) for val in section_values if val in ["BG", "YSF"]])
                else:
                    self.section_combo.clear()
                    self.section_combo.addItem("All")

                # Update filter dropdown based on the filter column
                filter_col = self.current_sheet_config.get("filter_col")
                if filter_col and filter_col in self.df.columns:
                    filter_values = sorted(self.df[filter_col].dropna().unique())
                    self.filter_combo.clear()
                    self.filter_combo.addItem("All")
                    self.filter_combo.addItems([str(val) for val in filter_values])
                else:
                    self.filter_combo.clear()
                    self.filter_combo.addItem("All")
                self.update_equipment_table()
            else:
                QMessageBox.critical(self, "Error", f"No valid data found in sheet {sheet_name}.")

    def update_equipment_table(self):
        if self.df is None:
            return

        # Filter by section
        section_value = self.section_combo.currentText()
        filtered_df = self.df
        if section_value != "All" and 'Section' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Section'] == section_value]

        # Filter by additional filter column
        filter_value = self.filter_combo.currentText()
        filter_col = self.current_sheet_config.get("filter_col")
        if filter_value != "All" and filter_col and filter_col in filtered_df.columns:
            filtered_df = filtered_df[filtered_df[filter_col] == filter_value]

        # Set up table
        self.equipment_table.clear()
        columns = self.current_sheet_config.get("headers", [])
        self.equipment_table.setColumnCount(len(columns))
        self.equipment_table.setHorizontalHeaderLabels(columns)
        self.equipment_table.setRowCount(len(filtered_df))

        # Populate table
        for row_idx, (_, row) in enumerate(filtered_df.iterrows()):
            for col_idx, col in enumerate(columns):
                item = QTableWidgetItem(str(row.get(col, "")))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)  # Make cells read-only
                self.equipment_table.setItem(row_idx, col_idx, item)

        self.equipment_table.resizeColumnsToContents()

    # Tab 2: Update Data
    def setup_update_tab(self):
        layout = QVBoxLayout(self.tab_update)
        layout.setSpacing(15)

        # Sheet selection
        sheet_frame = QFrame()
        sheet_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        sheet_layout = QHBoxLayout(sheet_frame)
        sheet_label = QLabel("Select Sheet:")
        sheet_label.setFont(QFont("Segoe UI", 12))
        sheet_label.setStyleSheet("color: #2d3436;")
        sheet_layout.addWidget(sheet_label)

        self.sheet_combo_update = QComboBox()
        self.sheet_combo_update.setFont(QFont("Segoe UI", 12))
        self.sheet_combo_update.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                padding: 5px;
                border-radius: 3px;
                color: #2d3436;
            }
            QComboBox::drop-down {
                border-left: 1px solid #dfe6e9;
                padding-right: 5px;
            }
            QComboBox:hover {
                border: 1px solid #0984e3;
            }
        """)
        self.sheet_combo_update.addItems(list(self.processed_data.keys()))
        self.sheet_combo_update.currentTextChanged.connect(self.load_sheet_update)
        sheet_layout.addWidget(self.sheet_combo_update)
        layout.addWidget(sheet_frame)

        # Input frame
        self.input_frame = QFrame()
        self.input_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        input_layout = QVBoxLayout(self.input_frame)
        input_layout.setSpacing(10)

        # Equipment selection
        equipment_row = QHBoxLayout()
        equipment_label = QLabel("Select Equipment:")
        equipment_label.setFont(QFont("Segoe UI", 12))
        equipment_label.setStyleSheet("color: #2d3436;")
        equipment_row.addWidget(equipment_label)

        self.equipment_combo = QComboBox()
        self.equipment_combo.setFont(QFont("Segoe UI", 12))
        self.equipment_combo.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                padding: 5px;
                border-radius: 3px;
                color: #2d3436;
            }
            QComboBox::drop-down {
                border-left: 1px solid #dfe6e9;
                padding-right: 5px;
            }
            QComboBox:hover {
                border: 1px solid #0984e3;
            }
        """)
        self.equipment_combo.currentTextChanged.connect(self.update_sous_ensemble)
        equipment_row.addWidget(self.equipment_combo)
        input_layout.addLayout(equipment_row)

        # Sous-ensemble selection
        sous_ensemble_row = QHBoxLayout()
        sous_ensemble_label = QLabel("Select Sous-ensemble:")
        sous_ensemble_label.setFont(QFont("Segoe UI", 12))
        sous_ensemble_label.setStyleSheet("color: #2d3436;")
        sous_ensemble_row.addWidget(sous_ensemble_label)

        self.sous_ensemble_combo = QComboBox()
        self.sous_ensemble_combo.setFont(QFont("Segoe UI", 12))
        self.sous_ensemble_combo.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                padding: 5px;
                border-radius: 3px;
                color: #2d3436;
            }
            QComboBox::drop-down {
                border-left: 1px solid #dfe6e9;
                padding-right: 5px;
            }
            QComboBox:hover {
                border: 1px solid #0984e3;
            }
        """)
        self.sous_ensemble_combo.currentTextChanged.connect(self.display_data)
        sous_ensemble_row.addWidget(self.sous_ensemble_combo)
        input_layout.addLayout(sous_ensemble_row)

        # Data display frame
        self.data_frame = QFrame()
        self.data_frame.setStyleSheet("""
            QFrame {
                background-color: #f5f6fa;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        self.data_layout = QVBoxLayout(self.data_frame)
        self.data_layout.setContentsMargins(10, 10, 10, 10)
        self.data_layout.setSpacing(10)

        self.labels = {}

        input_layout.addWidget(self.data_frame)
        layout.addWidget(self.input_frame)

        # Save button
        save_button = QPushButton("Save Changes")
        save_button.setFont(QFont("Segoe UI", 12))
        save_button.setStyleSheet("""
            QPushButton {
                background-color: #0984e3;
                color: #ffffff;
                padding: 10px 20px;
                border-radius: 5px;
                border: none;
                font: bold 12pt "Segoe UI";
            }
            QPushButton:hover {
                background-color: #0870c2;
            }
            QPushButton:pressed {
                background-color: #065aa1;
            }
        """)
        save_button.clicked.connect(self.save_data)
        layout.addWidget(save_button, alignment=Qt.AlignmentFlag.AlignCenter)

        # Load initial sheet
        if self.sheet_combo_update.currentText():
            self.load_sheet_update(self.sheet_combo_update.currentText())

    def load_sheet_update(self, sheet_name):
        if sheet_name:
            self.df = self.processed_data[sheet_name]
            self.current_sheet_config = self.sheet_configs.get(sheet_name, {})
            if not self.df.empty:
                # Update equipment dropdown
                equipment_col = "Equipement" if "Equipement" in self.df.columns else "équipement"
                equipment_list = sorted(self.df[equipment_col].dropna().unique())
                self.equipment_combo.clear()
                self.equipment_combo.addItems([str(e) for e in equipment_list])
                self.sous_ensemble_combo.clear()

                # Update data fields dynamically based on numeric columns
                for i in reversed(range(self.data_layout.count())):
                    widget = self.data_layout.itemAt(i).widget()
                    if widget:
                        widget.deleteLater()

                self.labels.clear()
                numeric_cols = self.current_sheet_config.get("numeric_cols", [])
                for col in numeric_cols:
                    if col in self.df.columns:
                        row_layout = QHBoxLayout()
                        label = QLabel(f"{col}:")
                        label.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
                        label.setStyleSheet("color: #2d3436;")
                        row_layout.addWidget(label)

                        entry = QLineEdit()
                        entry.setFont(QFont("Segoe UI", 12))
                        entry.setStyleSheet("""
                            QLineEdit {
                                background-color: #ffffff;
                                border: 1px solid #dfe6e9;
                                padding: 5px;
                                border-radius: 3px;
                                color: #2d3436;
                            }
                            QLineEdit:hover {
                                border: 1px solid #0984e3;
                            }
                        """)
                        entry.setFixedHeight(35)
                        row_layout.addWidget(entry)

                        self.labels[col] = entry
                        self.data_layout.addLayout(row_layout)

                self.clear_data_fields()
            else:
                QMessageBox.critical(self, "Error", f"No valid data found in sheet {sheet_name}.")

    def update_sous_ensemble(self, equipment):
        if equipment and self.df is not None:
            sous_ensemble_col = "Sous-ensemble" if "Sous-ensemble" in self.df.columns else "Sous ensemble"
            if sous_ensemble_col in self.df.columns:
                equipment_col = "Equipement" if "Equipement" in self.df.columns else "équipement"
                sous_ensemble_list = sorted(self.df[self.df[equipment_col] == equipment][sous_ensemble_col].dropna().unique())
                self.sous_ensemble_combo.clear()
                self.sous_ensemble_combo.addItems([str(s) for s in sous_ensemble_list])
                self.clear_data_fields()

    def display_data(self, sous_ensemble):
        equipment = self.equipment_combo.currentText()
        if equipment and sous_ensemble and self.df is not None:
            equipment_col = "Equipement" if "Equipement" in self.df.columns else "équipement"
            sous_ensemble_col = "Sous-ensemble" if "Sous-ensemble" in self.df.columns else "Sous ensemble"
            if sous_ensemble_col in self.df.columns:
                row = self.df[(self.df[equipment_col] == equipment) &
                              (self.df[sous_ensemble_col] == sous_ensemble)]
                if not row.empty:
                    for col, entry in self.labels.items():
                        entry.setText(str(row[col].values[0]))

    def clear_data_fields(self):
        for entry in self.labels.values():
            entry.clear()

    def save_data(self):
        sheet_name = self.sheet_combo_update.currentText()
        equipment = self.equipment_combo.currentText()
        sous_ensemble = self.sous_ensemble_combo.currentText()

        if sheet_name and equipment and sous_ensemble and self.df is not None:
            equipment_col = "Equipement" if "Equipement" in self.df.columns else "équipement"
            sous_ensemble_col = "Sous-ensemble" if "Sous-ensemble" in self.df.columns else "Sous ensemble"
            if sous_ensemble_col in self.df.columns:
                idx = self.df[(self.df[equipment_col] == equipment) &
                              (self.df[sous_ensemble_col] == sous_ensemble)].index
                if not idx.empty:
                    # Validate inputs
                    for col, entry in self.labels.items():
                        try:
                            value = float(entry.text()) if entry.text() else 0
                            if value < 0:
                                QMessageBox.critical(self, "Error", f"Value for {col} cannot be negative.")
                                return
                            self.df.at[idx[0], col] = value
                        except ValueError:
                            QMessageBox.critical(self, "Error", f"Invalid value for {col}. Please enter a number.")
                            return

                    try:
                        with pd.ExcelWriter("ENGINS.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                            for sheet, sheet_data in self.processed_data.items():
                                if sheet == sheet_name:
                                    # Remove the temporary 'Section' column before saving
                                    sheet_data = self.df.drop(columns=['Section'], errors='ignore')
                                sheet_data.to_excel(writer, sheet_name=sheet, index=False)
                        QMessageBox.information(self, "Success", "Data saved successfully!")
                        # Refresh the equipment table and dashboard
                        self.load_sheet_equipment(self.sheet_combo_equipment.currentText())
                        self.update_dashboard()
                    except Exception as e:
                        QMessageBox.critical(self, "Error", f"Failed to save data: {e}")
                else:
                    QMessageBox.critical(self, "Error", "Selected equipment and sous-ensemble not found in data.")
            else:
                QMessageBox.critical(self, "Error", "Sous-ensemble column not found in the selected sheet.")
        else:
            QMessageBox.critical(self, "Error", "Please select a sheet, equipment, and sous-ensemble.")

    # Tab 3: Dashboard
    def setup_dashboard_tab(self):
        layout = QVBoxLayout(self.tab_dashboard)
        layout.setSpacing(15)

        # Sheet selection for dashboard
        sheet_frame = QFrame()
        sheet_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        sheet_layout = QHBoxLayout(sheet_frame)
        sheet_label = QLabel("Select Sheet:")
        sheet_label.setFont(QFont("Segoe UI", 12))
        sheet_label.setStyleSheet("color: #2d3436;")
        sheet_layout.addWidget(sheet_label)

        self.sheet_combo_dashboard = QComboBox()
        self.sheet_combo_dashboard.setFont(QFont("Segoe UI", 12))
        self.sheet_combo_dashboard.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                padding: 5px;
                border-radius: 3px;
                color: #2d3436;
            }
            QComboBox::drop-down {
                border-left: 1px solid #dfe6e9;
                padding-right: 5px;
            }
            QComboBox:hover {
                border: 1px solid #0984e3;
            }
        """)
        self.sheet_combo_dashboard.addItems(list(self.processed_data.keys()))
        self.sheet_combo_dashboard.currentTextChanged.connect(self.update_dashboard)
        sheet_layout.addWidget(self.sheet_combo_dashboard)
        layout.addWidget(sheet_frame)

        # Dashboard content
        self.dashboard_frame = QFrame()
        self.dashboard_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        dashboard_layout = QVBoxLayout(self.dashboard_frame)
        dashboard_layout.setSpacing(10)

        # Statistics
        self.stats_label = QLabel("Statistics:")
        self.stats_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        self.stats_label.setStyleSheet("color: #2d3436; margin-bottom: 5px;")
        dashboard_layout.addWidget(self.stats_label)

        self.stats_text = QLabel()
        self.stats_text.setFont(QFont("Segoe UI", 12))
        self.stats_text.setStyleSheet("color: #2d3436; background-color: #f5f6fa; padding: 10px; border-radius: 5px;")
        dashboard_layout.addWidget(self.stats_text)

        # Alerts
        self.alerts_label = QLabel("Alerts:")
        self.alerts_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        self.alerts_label.setStyleSheet("color: #2d3436; margin-top: 10px; margin-bottom: 5px;")
        dashboard_layout.addWidget(self.alerts_label)

        self.alerts_area = QScrollArea()
        self.alerts_area.setWidgetResizable(True)
        self.alerts_area.setStyleSheet("""
            QScrollArea {
                background-color: #f5f6fa;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
            }
        """)
        self.alerts_widget = QWidget()
        self.alerts_layout = QVBoxLayout(self.alerts_widget)
        self.alerts_area.setWidget(self.alerts_widget)
        dashboard_layout.addWidget(self.alerts_area)

        layout.addWidget(self.dashboard_frame)

        # Load initial dashboard
        if self.sheet_combo_dashboard.currentText():
            self.update_dashboard()

    def update_dashboard(self):
        sheet_name = self.sheet_combo_dashboard.currentText()
        if not sheet_name or sheet_name not in self.processed_data:
            return

        df = self.processed_data[sheet_name]
        config = self.sheet_configs.get(sheet_name, {})
        if df.empty:
            self.stats_text.setText("No data available.")
            return

        # Calculate statistics based on sheet type
        equipment_col = "Equipement" if "Equipement" in df.columns else "équipement"
        sous_ensemble_col = "Sous-ensemble" if "Sous-ensemble" in df.columns else "Sous ensemble"

        total_equipments = len(df[equipment_col].unique()) if equipment_col in df.columns else 0
        total_sous_ensembles = len(df[sous_ensemble_col].dropna()) if sous_ensemble_col in df.columns else 0

        stats = f"Total Equipments: {total_equipments}\n"
        if sous_ensemble_col in df.columns:
            stats += f"Total Sous-ensembles: {total_sous_ensembles}\n"

        # Sheet-specific statistics
        if sheet_name in ["Cartographie moteur", "Cartographie transmission", "Cartographie Engin"]:
            awaiting_revision = int(df["Sous-ensemble en attente révision"].sum())
            in_progress = int(df["Sous-ensemble encours de révision"].sum())
            stats += (
                f"Sous-ensembles Awaiting Revision: {awaiting_revision}\n"
                f"Sous-ensembles In Progress: {in_progress}"
            )
        elif sheet_name in ["Programme 2025 BG", "Programme 2025 YSF"]:
            total_cost = 0
            if "Cout V2" in df.columns:
                total_cost += df["Cout V2"].sum()
            if "Cout [V2]" in df.columns:
                total_cost += df["Cout [V2]"].sum()
            if "Cout V1" in df.columns:
                total_cost += df["Cout V1"].sum()
            stats += f"Total Estimated Cost: {total_cost:.2f}"

        self.stats_text.setText(stats)

        # Clear previous alerts
        for i in reversed(range(self.alerts_layout.count())):
            widget = self.alerts_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        # Generate alerts for specific sheets
        if sheet_name in ["Cartographie moteur", "Cartographie transmission", "Cartographie Engin"]:
            critical_rows = df[
                (df["Sous-ensemble relais disponible (révisé)"] == 0) &
                (df["Sous-ensemble en attente révision"] > 0)
            ]
            for _, row in critical_rows.iterrows():
                alert = f"Critical: {row['Equipement']} - {row['Sous-ensemble']} has 0 available and {row['Sous-ensemble en attente révision']} awaiting revision."
                alert_label = QLabel(alert)
                alert_label.setFont(QFont("Segoe UI", 12))
                alert_label.setStyleSheet("color: #d63031; padding: 5px; background-color: #ffcccc; border-radius: 3px;")
                self.alerts_layout.addWidget(alert_label)

            if critical_rows.empty:
                no_alert = QLabel("No critical alerts.")
                no_alert.setFont(QFont("Segoe UI", 12))
                no_alert.setStyleSheet("color: #2d3436; padding: 5px; background-color: #e6ffed; border-radius: 3px;")
                self.alerts_layout.addWidget(no_alert)
        else:
            no_alert = QLabel("Alerts not applicable for this sheet.")
            no_alert.setFont(QFont("Segoe UI", 12))
            no_alert.setStyleSheet("color: #2d3436; padding: 5px; background-color: #e6ffed; border-radius: 3px;")
            self.alerts_layout.addWidget(no_alert)

def main():
    app = QApplication(sys.argv)
    window = EquipmentApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()