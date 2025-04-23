import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QComboBox, QPushButton, QLineEdit, QFrame, QMessageBox,
                             QTableWidget, QTableWidgetItem, QTabWidget, QScrollArea, QCheckBox, QButtonGroup, QRadioButton)

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
            self.excel_data = pd.read_excel("Cartographie SE par atelier.xlsx", sheet_name=None, header=None)
        except FileNotFoundError:
            self.excel_data = {}
            QMessageBox.critical(self, "Error", "Excel file 'Cartographie SE par atelier.xlsx' not found.")
            return

        # Preprocess sheets to extract relevant sections
        self.processed_data = {}
        for sheet_name, df in self.excel_data.items():
            self.processed_data[sheet_name] = self.preprocess_sheet(df, sheet_name)

        # Current sheet data
        self.df = None

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

        # Tab 3: Dashboard
        self.tab_workflow = QWidget()
        self.tabs.addTab(self.tab_workflow, "Workflow")
        self.setup_workflow_tab()


    def preprocess_sheet(self, df, sheet_name):
        if df.empty:
            print(f"Sheet {sheet_name} is empty. Skipping preprocessing.")
            return pd.DataFrame()

        # Find the header row
        target_header = ["Equipement", "Sous-ensemble", "Quantité SE installée",
                         "Sous-ensemble relais disponible (révisé)", "Sous-ensemble en attente révision",
                         "Sous-ensemble encours de révision", "Corps de Sous-ensembles disponibles (révisable)"]
        header_row = None
        for i in range(len(df)):
            row = df.iloc[i].astype(str).str.strip()
            if all(col in row.values for col in target_header):
                header_row = i
                break

        if header_row is not None:
            df_section = pd.read_excel("Cartographie SE par atelier.xlsx", sheet_name=sheet_name, skiprows=header_row)
            df_section = df_section.dropna(subset=["Equipement", "Sous-ensemble"], how="all")
            df_section.columns = df_section.columns.str.strip()
            # Convert numeric columns to appropriate types
            numeric_cols = ["Quantité SE installée", "Sous-ensemble relais disponible (révisé)",
                            "Sous-ensemble en attente révision", "Sous-ensemble encours de révision",
                            "Corps de Sous-ensembles disponibles (révisable)"]
            for col in numeric_cols:
                df_section[col] = pd.to_numeric(df_section[col], errors='coerce').fillna(0)
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

        # Installation filter
        install_frame = QFrame()
        install_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        install_layout = QHBoxLayout(install_frame)
        install_label = QLabel("Filter by Installation:")
        install_label.setFont(QFont("Segoe UI", 12))
        install_label.setStyleSheet("color: #2d3436;")
        install_layout.addWidget(install_label)

        self.install_combo = QComboBox()
        self.install_combo.setFont(QFont("Segoe UI", 12))
        self.install_combo.setStyleSheet("""
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
        self.install_combo.addItem("All")
        self.install_combo.currentTextChanged.connect(self.update_equipment_table)
        install_layout.addWidget(self.install_combo)
        layout.addWidget(install_frame)

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
            if not self.df.empty:
                # Update installation dropdown
                installations = sorted(self.df.get("Unnamed: 0", pd.Series([])).dropna().unique())
                self.install_combo.clear()
                self.install_combo.addItem("All")
                self.install_combo.addItems(installations)
                self.update_equipment_table()
            else:
                QMessageBox.critical(self, "Error", f"No valid data found in sheet {sheet_name}.")

    def update_equipment_table(self):
        if self.df is None:
            return

        # Filter by installation
        install_filter = self.install_combo.currentText()
        if install_filter == "All":
            filtered_df = self.df
        else:
            filtered_df = self.df[self.df["Unnamed: 0"] == install_filter]

        # Set up table
        self.equipment_table.clear()
        columns = ["Equipement", "Sous-ensemble", "Quantité SE installée",
                   "Sous-ensemble relais disponible (révisé)", "Sous-ensemble en attente révision",
                   "Sous-ensemble encours de révision", "Corps de Sous-ensembles disponibles (révisable)"]
        self.equipment_table.setColumnCount(len(columns))
        self.equipment_table.setHorizontalHeaderLabels(columns)
        self.equipment_table.setRowCount(len(filtered_df))

        # Populate table
        for row_idx, (_, row) in enumerate(filtered_df.iterrows()):
            for col_idx, col in enumerate(columns):
                item = QTableWidgetItem(str(row[col]))
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
        data_layout = QVBoxLayout(self.data_frame)
        data_layout.setContentsMargins(10, 10, 10, 10)
        data_layout.setSpacing(10)

        self.labels = {}
        self.columns_to_display = [
            "Quantité SE installée", "Sous-ensemble relais disponible (révisé)",
            "Sous-ensemble en attente révision", "Sous-ensemble encours de révision",
            "Corps de Sous-ensembles disponibles (révisable)"
        ]

        for col in self.columns_to_display:
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
            data_layout.addLayout(row_layout)

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
            if not self.df.empty:
                equipment_list = sorted(self.df["Equipement"].dropna().unique())
                self.equipment_combo.clear()
                self.equipment_combo.addItems(equipment_list)
                self.sous_ensemble_combo.clear()
                self.clear_data_fields()
            else:
                QMessageBox.critical(self, "Error", f"No valid data found in sheet {sheet_name}.")

    def update_sous_ensemble(self, equipment):
        if equipment and self.df is not None:
            sous_ensemble_list = sorted(self.df[self.df["Equipement"] == equipment]["Sous-ensemble"].dropna().unique())
            self.sous_ensemble_combo.clear()
            self.sous_ensemble_combo.addItems(sous_ensemble_list)
            self.clear_data_fields()

    def display_data(self, sous_ensemble):
        equipment = self.equipment_combo.currentText()
        if equipment and sous_ensemble and self.df is not None:
            row = self.df[(self.df["Equipement"] == equipment) &
                          (self.df["Sous-ensemble"] == sous_ensemble)]
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
            idx = self.df[(self.df["Equipement"] == equipment) &
                          (self.df["Sous-ensemble"] == sous_ensemble)].index
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
                    with pd.ExcelWriter("Cartographie SE par atelier.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                        for sheet, sheet_data in self.processed_data.items():
                            if sheet == sheet_name:
                                sheet_data = self.df
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

    def setup_workflow_tab(self):
        layout = QVBoxLayout(self.tab_workflow)
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

        # Checklist for revision process
        checklist_frame = QFrame()
        checklist_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        checklist_layout = QVBoxLayout(checklist_frame)
        
        checklist_label = QLabel("S/E Revision Process Workflow:")
        checklist_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        checklist_label.setStyleSheet("color: #2d3436; margin-bottom: 10px;")
        checklist_layout.addWidget(checklist_label)

        # Dictionary to store widget states
        self.checklist_widgets = {}
        self.radio_groups = {}

        # Helper function to create a step label
        def create_step_label(text, indent=0):
            step_frame = QFrame()
            step_frame.setStyleSheet("""
                QFrame {
                    border: 1px solid #dfe6e9;
                    border-radius: 3px;
                    padding: 5px;
                    background-color: #f8f9fa;
                }
            """)
            step_layout = QHBoxLayout(step_frame)
            step_layout.setContentsMargins(indent * 20, 2, 2, 2)
            label = QLabel(text)
            label.setFont(QFont("Segoe UI", 11))
            label.setStyleSheet("color: #2d3436; border: none;")
            step_layout.addWidget(label)
            return step_frame

        # Helper function to create a decision point with Oui/Non radio buttons
        def create_decision_point(text, indent=0):
            decision_frame = QFrame()
            decision_frame.setStyleSheet("""
                QFrame {
                    border: 1px solid #0984e3;
                    border-radius: 3px;
                    padding: 5px;
                    background-color: #e7f3ff;
                }
            """)
            decision_layout = QVBoxLayout(decision_frame)
            decision_layout.setContentsMargins(indent * 20, 2, 2, 2)
            
            label = QLabel(text)
            label.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
            label.setStyleSheet("color: #2d3436; border: none;")
            decision_layout.addWidget(label)

            radio_layout = QHBoxLayout()
            group = QButtonGroup(decision_frame)
            oui_radio = QRadioButton("Oui")
            non_radio = QRadioButton("Non")
            oui_radio.setFont(QFont("Segoe UI", 10))
            non_radio.setFont(QFont("Segoe UI", 10))
            oui_radio.setStyleSheet("color: #2d3436; padding: 2px;")
            non_radio.setStyleSheet("color: #2d3436; padding: 2px;")
            group.addButton(oui_radio)
            group.addButton(non_radio)
            radio_layout.addWidget(oui_radio)
            radio_layout.addWidget(non_radio)
            decision_layout.addLayout(radio_layout)

            return decision_frame, group, oui_radio, non_radio

        # Build the workflow following the diagram
        # Step 1: Expert S/E à réviser
        step1 = create_step_label("Expert S/E à réviser")
        checklist_layout.addWidget(step1)
        self.checklist_widgets["Expert S/E à réviser"] = step1

        # Decision 1: Besoin en PDR?
        decision1, group1, oui1, non1 = create_decision_point("Besoin en PDR?")
        checklist_layout.addWidget(decision1)
        self.radio_groups["Besoin en PDR?"] = (group1, oui1, non1)

        # Oui path for Decision 1
        oui_frame1 = QFrame()
        oui_layout1 = QVBoxLayout(oui_frame1)
        step2 = create_step_label("Identification S/E et besoin en PDR, équipement, logistique", indent=1)
        oui_layout1.addWidget(step2)
        self.checklist_widgets["Identification S/E et besoin en PDR, équipement, logistique"] = step2

        # Decision 2: Besoin en MEC?
        decision2, group2, oui2, non2 = create_decision_point("Besoin en MEC?", indent=1)
        oui_layout1.addWidget(decision2)
        self.radio_groups["Besoin en MEC?"] = (group2, oui2, non2)

        # Oui path for Decision 2
        oui_frame2 = QFrame()
        oui_layout2 = QVBoxLayout(oui_frame2)
        step3 = create_step_label("Récupération équipement", indent=2)
        oui_layout2.addWidget(step3)
        self.checklist_widgets["Récupération équipement"] = step3
        oui_layout1.addWidget(oui_frame2)
        oui_frame2.setVisible(False)  # Initially hidden

        checklist_layout.addWidget(oui_frame1)
        oui_frame1.setVisible(False)  # Initially hidden

        # Processus de préparation (common path)
        step4 = create_step_label("Processus de préparation")
        checklist_layout.addWidget(step4)
        self.checklist_widgets["Processus de préparation"] = step4

        step5 = create_step_label("Récupération Bon sortie OT une fois la fiche de préparation est en position sur le tableau de préparation.")
        checklist_layout.addWidget(step5)
        self.checklist_widgets["Récupération Bon sortie OT"] = step5

        # Decision 3: S/E critique?
        decision3, group3, oui3, non3 = create_decision_point("S/E critique (Moteur thermique, moteur de roue, redacteur, Tracks...)")
        checklist_layout.addWidget(decision3)
        self.radio_groups["S/E critique"] = (group3, oui3, non3)

        # Oui path for Decision 3
        oui_frame3 = QFrame()
        oui_layout3 = QVBoxLayout(oui_frame3)
        step6 = create_step_label("Établissement d'un planning de révision", indent=1)
        oui_layout3.addWidget(step6)
        self.checklist_widgets["Établissement d'un planning de révision"] = step6
        checklist_layout.addWidget(oui_frame3)
        oui_frame3.setVisible(False)  # Initially hidden

        # Decision 4: Intervention nécessitant un outillage ou un réglage spécifique?
        decision4, group4, oui4, non4 = create_decision_point("Intervention nécessitant un outillage ou un réglage spécifique ou présenté pour le personnel ?")
        checklist_layout.addWidget(decision4)
        self.radio_groups["Intervention spécifique"] = (group4, oui4, non4)

        # Oui path for Decision 4
        oui_frame4 = QFrame()
        oui_layout4 = QVBoxLayout(oui_frame4)
        step7 = create_step_label("Préparation G.O.", indent=1)
        oui_layout4.addWidget(step7)
        self.checklist_widgets["Préparation G.O."] = step7
        checklist_layout.addWidget(oui_frame4)
        oui_frame4.setVisible(False)  # Initially hidden

        # Final steps
        step8 = create_step_label("Lancement des travaux de révision S/E")
        checklist_layout.addWidget(step8)
        self.checklist_widgets["Lancement des travaux de révision S/E"] = step8

        step9 = create_step_label("Instruction de la carte d'identification du S/E et la déplacer dans la zone (En cours de révision)")
        checklist_layout.addWidget(step9)
        self.checklist_widgets["Instruction de la carte d'identification"] = step9

        # Connect radio buttons to show/hide relevant sections
        oui1.toggled.connect(lambda checked: oui_frame1.setVisible(checked))
        oui2.toggled.connect(lambda checked: oui_frame2.setVisible(checked))
        oui3.toggled.connect(lambda checked: oui_frame3.setVisible(checked))
        oui4.toggled.connect(lambda checked: oui_frame4.setVisible(checked))

        # Load previous states if any
        self.load_checklist_states()

        layout.addWidget(checklist_frame)

    def save_checklist_state(self, step, state):
        # Placeholder for storage
        if step in self.checklist_widgets:
            self.checklist_widgets[step].setEnabled(bool(state))
        elif step in self.radio_groups:
            group, oui, non = self.radio_groups[step]
            if state == "Oui":
                oui.setChecked(True)
            elif state == "Non":
                non.setChecked(True)

    def load_checklist_states(self):
        # Placeholder for loading
        pass

    def update_dashboard(self):
        sheet_name = self.sheet_combo_dashboard.currentText()
        if not sheet_name or sheet_name not in self.processed_data:
            return

        df = self.processed_data[sheet_name]
        if df.empty:
            self.stats_text.setText("No data available.")
            return

        # Calculate statistics
        total_equipments = len(df["Equipement"].unique())
        total_sous_ensembles = len(df)
        awaiting_revision = int(df["Sous-ensemble en attente révision"].sum())
        in_progress = int(df["Sous-ensemble encours de révision"].sum())
        stats = (
            f"Total Equipments: {total_equipments}\n"
            f"Total Sous-ensembles: {total_sous_ensembles}\n"
            f"Sous-ensembles Awaiting Revision: {awaiting_revision}\n"
            f"Sous-ensembles In Progress: {in_progress}"
        )
        self.stats_text.setText(stats)

        # Clear previous alerts
        for i in reversed(range(self.alerts_layout.count())):
            widget = self.alerts_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        # Generate alerts
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

def main():
    app = QApplication(sys.argv)
    window = EquipmentApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()