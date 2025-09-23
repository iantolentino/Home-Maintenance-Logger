"""
Home Maintenance Logger
- Modern UI using PySide6
- CRUD -> data.json (auto-saved) and Export to Excel (data_YYYY-MM-DD.xlsx)
- No login required
- Window always centered
"""

import sys
import os
import json
import uuid
from datetime import datetime

import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QTextEdit,
    QDateEdit, QComboBox, QDoubleSpinBox, QPushButton, QHBoxLayout,
    QVBoxLayout, QTableWidget, QTableWidgetItem, QMessageBox, QFileDialog,
    QHeaderView, QAbstractItemView
)
from PySide6.QtCore import Qt, QDate, QTimer

DATA_JSON = "data.json"


class HomeMaintenanceLogger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Home Maintenance Logger")
        self.setMinimumSize(1000, 600)

        self._center_timer = QTimer(self)
        self._center_timer.setInterval(400)  # ms
        self._center_timer.timeout.connect(self.center)
        self._center_timer.start()

        self.data = []  # list of dicts
        self.current_selected_row = None

        self.init_ui()
        self.load_data()
        self.center()

    # --- Center window
    def center(self):
        screen = QApplication.primaryScreen()
        if not screen:
            return
        geo = screen.availableGeometry()
        x = geo.x() + (geo.width() - self.width()) // 2
        y = geo.y() + (geo.height() - self.height()) // 2
        self.move(x, y)

    # --- UI setup
    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout()
        central.setLayout(main_layout)

        # --- Left form
        form_layout = QVBoxLayout()

        self.input_id = QLineEdit()
        self.input_id.setReadOnly(True)
        self.input_id.setPlaceholderText("Auto-generated ID")

        self.input_date = QDateEdit()
        self.input_date.setCalendarPopup(True)
        self.input_date.setDate(QDate.currentDate())

        self.input_category = QComboBox()
        self.input_category.setEditable(True)
        self.input_category.addItems([
            "Plumbing", "Electrical", "Appliances", "HVAC", "Garden", "Roofing", "Other"
        ])

        self.input_item = QLineEdit()
        self.input_item.setPlaceholderText("e.g., Water heater, AC filter")

        self.input_desc = QTextEdit()
        self.input_desc.setPlaceholderText("Description of maintenance/issue")

        self.input_cost = QDoubleSpinBox()
        self.input_cost.setPrefix("₱ ")
        self.input_cost.setMaximum(1_000_000)
        self.input_cost.setDecimals(2)

        self.input_status = QComboBox()
        self.input_status.addItems(["Planned", "In Progress", "Completed", "Skipped"])

        self.input_next_service = QDateEdit()
        self.input_next_service.setCalendarPopup(True)
        self.input_next_service.setDate(QDate.currentDate())

        self.input_notes = QTextEdit()
        self.input_notes.setPlaceholderText("Optional notes")

        # Buttons
        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton("Add")
        self.btn_update = QPushButton("Update")
        self.btn_delete = QPushButton("Delete")
        self.btn_clear = QPushButton("Clear")
        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_update)
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addWidget(self.btn_clear)

        # Utilities
        util_layout = QHBoxLayout()
        self.btn_export = QPushButton("Export Excel")
        self.btn_import = QPushButton("Import JSON")
        self.btn_save = QPushButton("Save JSON")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search item/category/notes...")
        self.btn_search = QPushButton("Search")
        util_layout.addWidget(self.btn_export)
        util_layout.addWidget(self.btn_import)
        util_layout.addWidget(self.btn_save)

        form_layout.addWidget(QLabel("Record ID")); form_layout.addWidget(self.input_id)
        form_layout.addWidget(QLabel("Date")); form_layout.addWidget(self.input_date)
        form_layout.addWidget(QLabel("Category")); form_layout.addWidget(self.input_category)
        form_layout.addWidget(QLabel("Item")); form_layout.addWidget(self.input_item)
        form_layout.addWidget(QLabel("Description")); form_layout.addWidget(self.input_desc)
        form_layout.addWidget(QLabel("Cost")); form_layout.addWidget(self.input_cost)
        form_layout.addWidget(QLabel("Status")); form_layout.addWidget(self.input_status)
        form_layout.addWidget(QLabel("Next Service Date")); form_layout.addWidget(self.input_next_service)
        form_layout.addWidget(QLabel("Notes")); form_layout.addWidget(self.input_notes)
        form_layout.addLayout(btn_layout)
        form_layout.addLayout(util_layout)
        form_layout.addWidget(self.search_input)
        form_layout.addWidget(self.btn_search)

        left_widget = QWidget()
        left_widget.setLayout(form_layout)
        left_widget.setMaximumWidth(420)

        # --- Right table
        right_layout = QVBoxLayout()
        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels([
            "ID", "Date", "Category", "Item", "Description",
            "Cost", "Status", "Next Service", "Notes"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        right_layout.addWidget(self.table)

        right_widget = QWidget()
        right_widget.setLayout(right_layout)

        main_layout.addWidget(left_widget)
        main_layout.addWidget(right_widget)

        # Connect
        self.btn_add.clicked.connect(self.add_entry)
        self.btn_update.clicked.connect(self.update_entry)
        self.btn_delete.clicked.connect(self.delete_entry)
        self.btn_clear.clicked.connect(self.clear_form)
        self.btn_export.clicked.connect(self.export_excel)
        self.btn_import.clicked.connect(self.import_json)
        self.btn_save.clicked.connect(self.save_data)
        self.btn_search.clicked.connect(self.search_entries)
        self.table.cellDoubleClicked.connect(self.load_selected_row)

    # --- Data persistence
    def load_data(self):
        if os.path.exists(DATA_JSON):
            try:
                with open(DATA_JSON, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
            except Exception:
                self.data = []
        else:
            self.data = []
        self.refresh_table()

    def save_data(self):
        try:
            with open(DATA_JSON, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(self, "Save Error", str(e))

    def export_excel(self):
        if not self.data:
            QMessageBox.warning(self, "No Data", "No records to export.")
            return
        df = pd.DataFrame(self.data)
        today = datetime.now().strftime("%Y-%m-%d")
        default_filename = f"home_maintenance_{today}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Export Excel", default_filename, "Excel Files (*.xlsx)")
        if path:
            df.to_excel(path, index=False)

    def import_json(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import JSON", "", "JSON Files (*.json)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, list):
                self.data.extend(loaded)
                self.refresh_table()
        except Exception as e:
            QMessageBox.critical(self, "Import Error", str(e))

    # --- CRUD
    def add_entry(self):
        entry = self.read_form()
        if not entry:
            return
        entry["id"] = uuid.uuid4().hex
        self.data.append(entry)
        self.refresh_table()
        self.save_data()
        self.clear_form()

    def update_entry(self):
        row = self.table.currentRow()
        if row < 0:
            return
        entry = self.read_form()
        if not entry:
            return
        entry["id"] = self.table.item(row, 0).text()
        for i, rec in enumerate(self.data):
            if rec.get("id") == entry["id"]:
                self.data[i] = entry
                break
        self.refresh_table()
        self.save_data()
        self.clear_form()

    def delete_entry(self):
        row = self.table.currentRow()
        if row < 0:
            return
        rid = self.table.item(row, 0).text()
        self.data = [r for r in self.data if r.get("id") != rid]
        self.refresh_table()
        self.save_data()
        self.clear_form()

    def clear_form(self):
        self.input_id.clear()
        self.input_date.setDate(QDate.currentDate())
        self.input_category.setCurrentIndex(0)
        self.input_item.clear()
        self.input_desc.clear()
        self.input_cost.setValue(0.0)
        self.input_status.setCurrentIndex(0)
        self.input_next_service.setDate(QDate.currentDate())
        self.input_notes.clear()

    def read_form(self):
        item = self.input_item.text().strip()
        if not item:
            QMessageBox.warning(self, "Missing field", "Item name is required.")
            return None
        return {
            "id": self.input_id.text() or "",
            "date": self.input_date.date().toString("yyyy-MM-dd"),
            "category": self.input_category.currentText(),
            "item": item,
            "description": self.input_desc.toPlainText().strip(),
            "cost": float(self.input_cost.value()),
            "status": self.input_status.currentText(),
            "next_service_date": self.input_next_service.date().toString("yyyy-MM-dd"),
            "notes": self.input_notes.toPlainText().strip()
        }

    def refresh_table(self):
        self.table.setRowCount(0)
        for rec in self.data:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(rec.get("id", "")))
            self.table.setItem(row, 1, QTableWidgetItem(rec.get("date", "")))
            self.table.setItem(row, 2, QTableWidgetItem(rec.get("category", "")))
            self.table.setItem(row, 3, QTableWidgetItem(rec.get("item", "")))
            self.table.setItem(row, 4, QTableWidgetItem(rec.get("description", "")))
            self.table.setItem(row, 5, QTableWidgetItem(f"{rec.get('cost', 0):.2f}"))
            self.table.setItem(row, 6, QTableWidgetItem(rec.get("status", "")))
            self.table.setItem(row, 7, QTableWidgetItem(rec.get("next_service_date", "")))
            self.table.setItem(row, 8, QTableWidgetItem(rec.get("notes", "")))

    def load_selected_row(self, row, column=None):
        if row < 0:
            return
        rid = self.table.item(row, 0).text()
        rec = next((r for r in self.data if r.get("id") == rid), None)
        if not rec:
            return
        self.input_id.setText(rec.get("id", ""))
        self.input_date.setDate(QDate.fromString(rec.get("date", ""), "yyyy-MM-dd"))
        self.input_category.setCurrentText(rec.get("category", ""))
        self.input_item.setText(rec.get("item", ""))
        self.input_desc.setPlainText(rec.get("description", ""))
        self.input_cost.setValue(float(rec.get("cost", 0)))
        self.input_status.setCurrentText(rec.get("status", "Planned"))
        self.input_next_service.setDate(QDate.fromString(rec.get("next_service_date", ""), "yyyy-MM-dd"))
        self.input_notes.setPlainText(rec.get("notes", ""))

    def search_entries(self):
        q = self.search_input.text().strip().lower()
        if not q:
            self.refresh_table()
            return
        filtered = [r for r in self.data if
                    q in r.get("item", "").lower()
                    or q in r.get("category", "").lower()
                    or q in r.get("notes", "").lower()
                    or q in r.get("description", "").lower()]
        self.table.setRowCount(0)
        for rec in filtered:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(rec.get("id", "")))
            self.table.setItem(row, 1, QTableWidgetItem(rec.get("date", "")))
            self.table.setItem(row, 2, QTableWidgetItem(rec.get("category", "")))
            self.table.setItem(row, 3, QTableWidgetItem(rec.get("item", "")))
            self.table.setItem(row, 4, QTableWidgetItem(rec.get("description", "")))
            self.table.setItem(row, 5, QTableWidgetItem(f"{rec.get('cost', 0):.2f}"))
            self.table.setItem(row, 6, QTableWidgetItem(rec.get("status", "")))
            self.table.setItem(row, 7, QTableWidgetItem(rec.get("next_service_date", "")))
            self.table.setItem(row, 8, QTableWidgetItem(rec.get("notes", "")))


def main():
    app = QApplication(sys.argv)

    # Try to apply qt-material theme
    try:
        from qt_material import apply_stylesheet
        apply_stylesheet(app, theme='light_blue.xml')
    except Exception:
        app.setStyleSheet("""
            QWidget { font-family: "Segoe UI", Roboto, Arial; font-size: 11pt; }
            QPushButton { padding: 8px; border-radius: 6px; }
            QLineEdit, QTextEdit, QDateEdit, QComboBox { padding: 6px; border-radius: 6px; }
        """)

    window = HomeMaintenanceLogger()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
