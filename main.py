import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QComboBox, QPushButton, QLineEdit, QFrame, QMessageBox)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
import pandas as pd

class EquipmentApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Equipment Management App")
        self.setGeometry(100, 100, 800, 600)
        
        # Central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        
        # Load the Excel file with multiple sheets
        try:
            self.excel_data = pd.read_excel("ENGINS.xlsx", sheet_name=None, header=None)
        except FileNotFoundError:
            self.excel_data = {}
            print("Excel file not found. Created an empty dictionary.")
        
        # Preprocess sheets to extract relevant sections
        self.processed_data = {}
        for sheet_name, df in self.excel_data.items():
            self.processed_data[sheet_name] = self.preprocess_sheet(df, sheet_name)
        
        # Current sheet data
        self.df = None
        
        # Title
        title_label = QLabel("Equipment Data Management")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont("Arial", 16, QFont.Weight.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c3e50;")
        main_layout.addWidget(title_label)
        
        # Sheet selection frame
        sheet_frame = QFrame()
        sheet_frame.setStyleSheet("background-color: #d5e8f5; padding: 5px;")
        sheet_layout = QHBoxLayout(sheet_frame)
        sheet_layout.setContentsMargins(10, 5, 10, 5)
        
        sheet_label = QLabel("Select Sheet:")
        sheet_label.setFont(QFont("Arial", 12))
        sheet_layout.addWidget(sheet_label)
        
        self.sheet_combo = QComboBox()
        self.sheet_combo.setFont(QFont("Arial", 12))
        self.sheet_combo.addItems(list(self.processed_data.keys()))
        self.sheet_combo.currentTextChanged.connect(self.load_sheet)
        sheet_layout.addWidget(self.sheet_combo)
        
        main_layout.addWidget(sheet_frame)
        
        # Input frame
        self.input_frame = QFrame()
        self.input_frame.setStyleSheet("background-color: #d5e8f5; padding: 10px;")
        input_layout = QVBoxLayout(self.input_frame)
        input_layout.setContentsMargins(10, 10, 10, 10)
        input_layout.setSpacing(10)
        
        # Equipment selection
        equipment_row = QHBoxLayout()
        equipment_label = QLabel("Select Equipment:")
        equipment_label.setFont(QFont("Arial", 12))
        equipment_row.addWidget(equipment_label)
        
        self.equipment_combo = QComboBox()
        self.equipment_combo.setFont(QFont("Arial", 12))
        self.equipment_combo.currentTextChanged.connect(self.update_sous_ensemble)
        equipment_row.addWidget(self.equipment_combo)
        input_layout.addLayout(equipment_row)
        
        # Sous-ensemble selection
        sous_ensemble_row = QHBoxLayout()
        sous_ensemble_label = QLabel("Select Sous-ensemble:")
        sous_ensemble_label.setFont(QFont("Arial", 12))
        sous_ensemble_row.addWidget(sous_ensemble_label)
        
        self.sous_ensemble_combo = QComboBox()
        self.sous_ensemble_combo.setFont(QFont("Arial", 12))
        self.sous_ensemble_combo.currentTextChanged.connect(self.display_data)
        sous_ensemble_row.addWidget(self.sous_ensemble_combo)
        input_layout.addLayout(sous_ensemble_row)
        
        main_layout.addWidget(self.input_frame)
        
        # Data display frame
        self.data_frame = QFrame()
        self.data_frame.setStyleSheet("background-color: #f0f9ff; padding: 10px;")
        data_layout = QVBoxLayout(self.data_frame)
        data_layout.setContentsMargins(10, 10, 10, 10)
        data_layout.setSpacing(10)
        
        self.labels = {}
        self.columns_to_display = [
            "Criticité", "Quantité SE installée", "Sous-ensemble relais disponible (révisé)",
            "Sous-ensemble en attente révision", "Sous-ensemble encours de révision",
            "Corps de Sous-ensembles disponibles (révisable)"
        ]
        
        for col in self.columns_to_display:
            row_layout = QHBoxLayout()
            
            label = QLabel(f"{col}:")
            label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
            label.setStyleSheet("color: #34495e;")
            row_layout.addWidget(label)
            
            entry = QLineEdit()
            entry.setFont(QFont("Arial", 12))
            entry.setStyleSheet("background-color: white; color: #2c3e50;")
            entry.setFixedHeight(30)
            row_layout.addWidget(entry)
            
            self.labels[col] = entry
            data_layout.addLayout(row_layout)
        
        main_layout.addWidget(self.data_frame, stretch=1)
        
        # Save button
        save_button = QPushButton("Save Changes")
        save_button.setFont(QFont("Arial", 12))
        save_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 8px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        save_button.clicked.connect(self.save_data)
        main_layout.addWidget(save_button, alignment=Qt.AlignmentFlag.AlignCenter)
    
    def preprocess_sheet(self, df, sheet_name):
        # Check if the DataFrame is empty
        if df.empty:
            print(f"Sheet {sheet_name} is empty. Skipping preprocessing.")
            return pd.DataFrame()

        # If the sheet is "Cartographie Engin", extract the "Sous-ensemble Details" section
        if sheet_name == "Cartographie Engin":
            # Find the header row for the "Sous-ensemble Details" section
            target_header = ["Equipement", "Sous-ensemble", "Criticité", "Quantité SE installée", 
                             "Sous-ensemble relais disponible (révisé)", "Sous-ensemble en attente révision", 
                             "Sous-ensemble encours de révision", "Corps de Sous-ensembles disponibles (révisable)"]
            header_row = None
            for i in range(len(df)):
                row = df.iloc[i].astype(str).str.strip()
                if all(col in row.values for col in target_header):
                    header_row = i
                    break
            
            if header_row is not None:
                # Read the data starting from the header row
                df_section = pd.read_excel("ENGINS.xlsx", sheet_name=sheet_name, skiprows=header_row)
                # Filter rows until we hit an empty row or irrelevant section
                df_section = df_section.dropna(subset=["Equipement", "Sous-ensemble"], how="all")
                # Ensure column names are consistent
                df_section.columns = df_section.columns.str.strip()
                return df_section
            else:
                print(f"Could not find the 'Sous-ensemble Details' section in sheet {sheet_name}.")
                return pd.DataFrame()
        else:
            # For other sheets, check if the first row contains the expected columns
            target_header = ["Equipement", "Sous-ensemble", "Criticité", "Quantité SE installée", 
                             "Sous-ensemble relais disponible (révisé)", "Sous-ensemble en attente révision", 
                             "Sous-ensemble encours de révision", "Corps de Sous-ensembles disponibles (révisable)"]
            first_row = df.iloc[0].astype(str).str.strip()
            if all(col in first_row.values for col in target_header):
                # If the first row is the header, re-read the sheet with the first row as header
                df = pd.read_excel("ENGINS.xlsx", sheet_name=sheet_name, header=0)
            # Convert column names to strings and strip whitespace
            df.columns = df.columns.astype(str).str.strip()
            return df
    
    def load_sheet(self, sheet_name):
        if sheet_name:
            self.df = self.processed_data[sheet_name]
            if not self.df.empty:
                # Update equipment dropdown
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
        sheet_name = self.sheet_combo.currentText()
        equipment = self.equipment_combo.currentText()
        sous_ensemble = self.sous_ensemble_combo.currentText()
        
        if sheet_name and equipment and sous_ensemble and self.df is not None:
            idx = self.df[(self.df["Equipement"] == equipment) & 
                          (self.df["Sous-ensemble"] == sous_ensemble)].index
            if not idx.empty:
                for col, entry in self.labels.items():
                    self.df.at[idx[0], col] = entry.text()
                try:
                    with pd.ExcelWriter("ENGINS.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                        for sheet, sheet_data in self.processed_data.items():
                            if sheet == sheet_name:
                                sheet_data = self.df
                            sheet_data.to_excel(writer, sheet_name=sheet, index=False)
                    QMessageBox.information(self, "Success", "Data saved successfully!")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to save data: {e}")
            else:
                QMessageBox.critical(self, "Error", "Selected equipment and sous-ensemble not found in data.")
        else:
            QMessageBox.critical(self, "Error", "Please select a sheet, equipment, and sous-ensemble.")

def main():
    app = QApplication(sys.argv)
    window = EquipmentApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()