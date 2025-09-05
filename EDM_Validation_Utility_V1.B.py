import sys
import pandas as pd
import sqlite3
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTextEdit, QLabel, QFileDialog, QListWidget, QAbstractItemView,
    QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView, QMenu
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QColor, QPalette
from PyQt5.QtWidgets import QSizePolicy

class ExcelSQLValidatorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EDM Validation Utility")
        self.setGeometry(100, 100, 1200, 800)

        self.db_conn = None
        self.data_files_loaded = {}
        self.test_cases_df = None
        self.validation_results = []
        self.manual_sql_result_table = None

        self.themes = ["Light", "Dark", "Blue"]
        self.current_theme = 0  # Start with Light

        self.init_ui()
        self.apply_styles()

    def apply_styles(self):
        theme = self.themes[self.current_theme]
        if theme == "Light":
            self.setStyleSheet("""
                QWidget { background-color: #f9fafb; font-family: Segoe UI, Arial, sans-serif; font-size: 13px; }
                QLabel#HeaderLabel { font-size: 28px; font-weight: bold; color: #0d47a1; padding: 10px 0 20px 0; }
                QLabel { color: #263238; font-weight: bold; }
                QPushButton { background-color: #1565c0; color: #fff; border-radius: 6px; padding: 8px 18px; font-size: 14px; font-weight: bold; }
                QPushButton:disabled { background-color: #b0bec5; color: #eceff1; }
                QPushButton:hover { background-color: #1976d2; }
                QTableWidget { background-color: #fff; border: 1px solid #90caf9; font-size: 13px; color: #263238; }
                QHeaderView::section { background-color: #1565c0; color: #fff; font-weight: bold; font-size: 13px; border: none; padding: 6px; }
                QListWidget { background-color: #e3f2fd; border: 1px solid #90caf9; color: #263238; }
                QTextEdit { background-color: #e3f2fd; border: 1px solid #90caf9; font-size: 13px; color: #263238; }
            """)
        elif theme == "Dark":
            self.setStyleSheet("""
                QWidget { background-color: #181c24; color: #e0e0e0; font-family: Segoe UI, Arial, sans-serif; font-size: 13px; }
                QLabel#HeaderLabel { font-size: 28px; font-weight: bold; color: #90caf9; padding: 10px 0 20px 0; }
                QLabel { color: #b0bec5; font-weight: bold; }
                QPushButton { background-color: #263859; color: #fff; border-radius: 6px; padding: 8px 18px; font-size: 14px; font-weight: bold; }
                QPushButton:disabled { background-color: #616161; color: #bdbdbd; }
                QPushButton:hover { background-color: #1976d2; }
                QTableWidget { background-color: #23272e; border: 1px solid #90caf9; font-size: 13px; color: #e0e0e0; }
                QHeaderView::section { background-color: #263859; color: #fff; font-weight: bold; font-size: 13px; border: none; padding: 6px; }
                QListWidget { background-color: #23272e; border: 1px solid #90caf9; color: #e0e0e0; }
                QTextEdit { background-color: #23272e; border: 1px solid #90caf9; font-size: 13px; color: #e0e0e0; }
            """)
        elif theme == "Blue":
            self.setStyleSheet("""
                QWidget { background-color: #eaf6fb; font-family: Segoe UI, Arial, sans-serif; font-size: 13px; }
                QLabel#HeaderLabel { font-size: 28px; font-weight: bold; color: #01579b; padding: 10px 0 20px 0; }
                QLabel { color: #01579b; font-weight: bold; }
                QPushButton { background-color: #005792; color: #fff; border-radius: 6px; padding: 8px 18px; font-size: 14px; font-weight: bold; }
                QPushButton:disabled { background-color: #b3e5fc; color: #eceff1; }
                QPushButton:hover { background-color: #0288d1; }
                QTableWidget { background-color: #fff; border: 1px solid #0288d1; font-size: 13px; color: #01579b; }
                QHeaderView::section { background-color: #0288d1; color: #fff; font-weight: bold; font-size: 13px; border: none; padding: 6px; }
                QListWidget { background-color: #b3e5fc; border: 1px solid #0288d1; color: #01579b; }
                QTextEdit { background-color: #b3e5fc; border: 1px solid #0288d1; font-size: 13px; color: #01579b; }
            """)

    def switch_theme(self):
        self.current_theme = (self.current_theme + 1) % len(self.themes)
        self.apply_styles()

    def init_ui(self):
        main_layout = QVBoxLayout()

        # --- Header ---
        header_label = QLabel("EDM Validation Utility")
        header_label.setObjectName("HeaderLabel")
        header_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(header_label)

        # --- Theme Switcher ---
        theme_btn = QPushButton("Switch Theme")
        theme_btn.clicked.connect(self.switch_theme)
        main_layout.addWidget(theme_btn, alignment=Qt.AlignRight)

        # --- File Selection Area ---
        file_selection_group_layout = QHBoxLayout()

        # Data Files Selection
        data_file_layout = QVBoxLayout()
        data_file_label = QLabel("1. Select Excel Data Files:")
        data_file_layout.addWidget(data_file_label)
        self.add_data_file_button = QPushButton("Add Data Excel File(s)")
        self.add_data_file_button.clicked.connect(self.add_data_excel_files)
        data_file_layout.addWidget(self.add_data_file_button)
        self.loaded_data_files_list = QListWidget()
        self.loaded_data_files_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        data_file_layout.addWidget(self.loaded_data_files_list)
        self.remove_data_file_button = QPushButton("Remove Selected Data File(s)")
        self.remove_data_file_button.clicked.connect(self.remove_data_excel_files)
        data_file_layout.addWidget(self.remove_data_file_button)

        file_selection_group_layout.addLayout(data_file_layout)

        # Test Case File Selection
        tc_file_layout = QVBoxLayout()
        tc_file_label = QLabel("2. Select Excel Test Case File:")
        tc_file_layout.addWidget(tc_file_label)
        self.load_tc_file_button = QPushButton("Load Test Case Excel File")
        self.load_tc_file_button.clicked.connect(self.load_test_case_excel)
        tc_file_layout.addWidget(self.load_tc_file_button)
        self.tc_file_path_label = QLabel("No test case file loaded.")
        tc_file_layout.addWidget(self.tc_file_path_label)
        self.view_tc_button = QPushButton("View Test Cases")
        self.view_tc_button.clicked.connect(self.view_test_cases)
        self.view_tc_button.setEnabled(False)
        tc_file_layout.addWidget(self.view_tc_button)
        tc_file_layout.addStretch(1)
        file_selection_group_layout.addLayout(tc_file_layout)

        main_layout.addLayout(file_selection_group_layout)

        # --- Actions ---
        action_layout = QHBoxLayout()
        self.run_validation_button = QPushButton("3. Run Validation")
        self.run_validation_button.clicked.connect(self.run_validation)
        self.run_validation_button.setEnabled(False)
        action_layout.addWidget(self.run_validation_button)

        self.clear_all_button = QPushButton("Clear All")
        self.clear_all_button.clicked.connect(self.clear_all)
        action_layout.addWidget(self.clear_all_button)

        main_layout.addLayout(action_layout)

        # --- Validation Report Area ---
        report_label = QLabel("Validation Report:")
        main_layout.addWidget(report_label)
        self.report_table = QTableWidget()
        self.report_table.setColumnCount(5)
        self.report_table.setHorizontalHeaderLabels(
            ["TC Name", "Status", "Expected Result", "Actual Result", "Error/Details"]
        )
        self.report_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.report_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.report_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.report_table.customContextMenuRequested.connect(self.show_report_table_context_menu)
        main_layout.addWidget(self.report_table)

        self.save_report_button = QPushButton("Save Report to Excel")
        self.save_report_button.clicked.connect(self.save_report_to_excel)
        self.save_report_button.setEnabled(False)
        main_layout.addWidget(self.save_report_button)

        # --- Manual SQL Execution Area (Side by Side) ---
        manual_sql_area = QHBoxLayout()
        # Left: SQL Editor
        left_sql_layout = QVBoxLayout()
        manual_sql_label = QLabel("Manual SQL Query:")
        left_sql_layout.addWidget(manual_sql_label)
        self.manual_sql_input = QTextEdit()
        self.manual_sql_input.setPlaceholderText("Type your SQL query here...")
        self.manual_sql_input.setMinimumHeight(120)
        self.manual_sql_input.setSizePolicy(self.manual_sql_input.sizePolicy().horizontalPolicy(), QSizePolicy.Expanding)
        left_sql_layout.addWidget(self.manual_sql_input)
        self.run_manual_sql_button = QPushButton("Run SQL")
        self.run_manual_sql_button.clicked.connect(self.run_manual_sql)
        left_sql_layout.addWidget(self.run_manual_sql_button)
        left_sql_layout.addStretch(1)
        manual_sql_area.addLayout(left_sql_layout, 4)  # 40%

        # Right: Output Table
        right_sql_layout = QVBoxLayout()
        output_label = QLabel("SQL Output:")
        right_sql_layout.addWidget(output_label)
        self.manual_sql_result_table = QTableWidget()
        self.manual_sql_result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.manual_sql_result_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.manual_sql_result_table.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
        self.manual_sql_result_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        right_sql_layout.addWidget(self.manual_sql_result_table)
        self.show_tables_button = QPushButton("Show Loaded Tables")
        self.show_tables_button.clicked.connect(self.show_loaded_tables)
        right_sql_layout.addWidget(self.show_tables_button)
        right_sql_layout.addStretch(1)
        manual_sql_area.addLayout(right_sql_layout, 6)  # 60%

        # Add the manual SQL area to the main layout
        main_layout.addLayout(manual_sql_area, stretch=1)

        self.setLayout(main_layout)
        self.update_run_button_state()

    def update_run_button_state(self):
        # Enable run button if data files and test cases are loaded
        if self.data_files_loaded and self.test_cases_df is not None:
            self.run_validation_button.setEnabled(True)
        else:
            self.run_validation_button.setEnabled(False)

    def connect_db(self):
        if self.db_conn:
            self.db_conn.close()
        self.db_conn = sqlite3.connect(":memory:") # In-memory database
        print("Connected to in-memory SQLite database.")

    def add_data_excel_files(self):
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("Excel Files (*.xlsx *.xls)")
        file_dialog.setFileMode(QFileDialog.ExistingFiles)

        if file_dialog.exec_():
            selected_files = file_dialog.selectedFiles()
            # Only connect if not already connected
            if not self.db_conn:
                self.connect_db()
            if self.data_files_loaded is None:
                self.data_files_loaded = {}

            for file_path in selected_files:
                already_loaded = any(self.loaded_data_files_list.item(i).text() == file_path
                                     for i in range(self.loaded_data_files_list.count()))
                if not already_loaded:
                    try:
                        with pd.ExcelFile(file_path) as xls:
                            sheet_names = xls.sheet_names
                            loaded_sheets = []

                            for sheet_name in sheet_names:
                                df = pd.read_excel(xls, sheet_name=sheet_name)
                                base_name = os.path.splitext(os.path.basename(file_path))[0]
                                table_name = f'{base_name}.{sheet_name}'
                                # Allow dot and underscore in table name
                                table_name = "".join(c for c in table_name if c.isalnum() or c in ['.', '_'])
                                df.to_sql(table_name, self.db_conn, if_exists='replace', index=False)
                                loaded_sheets.append(table_name)
                                print(f"Loaded '{sheet_name}' from '{base_name}' into table '{table_name}'")

                        self.data_files_loaded[file_path] = loaded_sheets
                        self.loaded_data_files_list.addItem(file_path)
                        QMessageBox.information(self, "Success", f"Loaded '{os.path.basename(file_path)}' with sheets: {', '.join(sheet_names)}.")
                    except Exception as e:
                        QMessageBox.warning(self, "Error Loading Data File", f"Could not load '{file_path}': {e}")
            self.update_run_button_state()

    def remove_data_excel_files(self):
        selected_items = self.loaded_data_files_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select file(s) to remove.")
            return

        for item in selected_items:
            file_path = item.text()
            # Remove tables loaded from this file
            if file_path in self.data_files_loaded:
                tables_to_drop = self.data_files_loaded[file_path]
                cursor = self.db_conn.cursor()
                for table_name in tables_to_drop:
                    try:
                        cursor.execute(f'DROP TABLE IF EXISTS "{table_name}"')
                    except Exception as e:
                        print(f"Error dropping table {table_name}: {e}")
                del self.data_files_loaded[file_path]
            # Remove from UI
            row = self.loaded_data_files_list.row(item)
            self.loaded_data_files_list.takeItem(row)

        self.update_run_button_state()


    def load_test_case_excel(self):
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("Excel Files (*.xlsx *.xls)")
        file_path, _ = file_dialog.getOpenFileName(self, "Load Test Case Excel File")

        if file_path:
            try:
                # Assuming test cases are in the first sheet
                self.test_cases_df = pd.read_excel(file_path)
                # Ensure required columns exist
                required_cols = ["TC_Name", "SQL_Query", "Expected_Result"]
                if not all(col in self.test_cases_df.columns for col in required_cols):
                    raise ValueError(f"Test case file must contain columns: {', '.join(required_cols)}")

                self.tc_file_path_label.setText(f"Loaded: {os.path.basename(file_path)}")
                self.view_tc_button.setEnabled(True)
                QMessageBox.information(self, "Success", f"Test cases loaded from '{os.path.basename(file_path)}'.")
            except Exception as e:
                self.test_cases_df = None
                self.tc_file_path_label.setText("No test case file loaded.")
                self.view_tc_button.setEnabled(False)
                QMessageBox.warning(self, "Error Loading TC File", f"Could not load test cases: {e}")
            self.update_run_button_state()

    def view_test_cases(self):
        if self.test_cases_df is not None:
            # Create a new window or dialog to show the test cases
            tc_viewer = QWidget()
            tc_viewer.setWindowTitle("Loaded Test Cases")
            tc_viewer.setGeometry(200, 200, 800, 600)
            layout = QVBoxLayout()
            table = QTableWidget()
            table.setColumnCount(len(self.test_cases_df.columns))
            table.setHorizontalHeaderLabels(self.test_cases_df.columns.tolist())
            table.setRowCount(len(self.test_cases_df))

            for i, row in self.test_cases_df.iterrows():
                for j, col_name in enumerate(self.test_cases_df.columns):
                    table.setItem(i, j, QTableWidgetItem(str(row[col_name])))
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.setEditTriggers(QTableWidget.NoEditTriggers)
            layout.addWidget(table)
            tc_viewer.setLayout(layout)
            tc_viewer.show()
            # This will create a local reference. For a persistent window,
            # you might need to store it as a member of the main app or a dialog.
            self.tc_viewer_window = tc_viewer
        else:
            QMessageBox.information(self, "No Test Cases", "No test case file is currently loaded.")

    def run_validation(self):
        self.validation_results = []
        self.report_table.setRowCount(0) # Clear previous results

        if not self.db_conn:
            QMessageBox.warning(self, "No Data", "No Excel data files loaded. Please load data first.")
            return
        if self.test_cases_df is None:
            QMessageBox.warning(self, "No Test Cases", "No test case file loaded. Please load test cases first.")
            return

        cursor = self.db_conn.cursor()

        for index, tc in self.test_cases_df.iterrows():
            tc_name = tc['TC_Name']
            sql_query = tc['SQL_Query']
            expected_result = str(tc['Expected_Result']).strip() # Convert to string for consistent comparison

            status = "FAIL"
            actual_result_str = ""
            error_details = ""

            try:
                cursor.execute(sql_query)
                query_results = cursor.fetchall()
                actual_result_str = str(query_results) # Default representation of results

                # --- Compare Actual vs. Expected ---
                if "0 rows" in expected_result:
                    if not query_results: # If query_results is empty
                        status = "PASS"
                    else:
                        actual_result_str = f"{len(query_results)} rows found."
                elif expected_result.startswith("COUNT = "):
                    expected_count = int(expected_result.split("=")[1].strip())
                    actual_count = len(query_results)
                    actual_result_str = f"COUNT = {actual_count}"
                    if actual_count == expected_count:
                        status = "PASS"
                elif expected_result.lower() == "no records": # Specific common expectation
                    if not query_results:
                        status = "PASS"
                    else:
                        actual_result_str = f"{len(query_results)} records found."
                elif expected_result.lower() == "records exist": # Specific common expectation
                    if query_results:
                        status = "PASS"
                        actual_result_str = f"{len(query_results)} records exist."
                    else:
                        actual_result_str = "No records found."
                elif query_results and len(query_results) == 1 and len(query_results[0]) == 1:
                    # If single scalar value returned
                    if str(query_results[0][0]).strip() == expected_result:
                        status = "PASS"
                        actual_result_str = str(query_results[0][0])
                    else:
                        actual_result_str = str(query_results[0][0])
                else:
                    # Generic comparison: just stringify both and compare
                    # This might be tricky for complex query results (order of columns/rows)
                    if actual_result_str == expected_result:
                         status = "PASS"
                    else:
                        error_details = f"Generic comparison failed. Actual: '{actual_result_str}', Expected: '{expected_result}'"

            except sqlite3.Error as e:
                status = "ERROR"
                error_details = f"SQL Error: {e}"
                actual_result_str = "N/A"
            except Exception as e:
                status = "ERROR"
                error_details = f"Unexpected Error: {e}"
                actual_result_str = "N/A"

            self.validation_results.append({
                "TC Name": tc_name,
                "Status": status,
                "Expected Result": expected_result,
                "Actual Result": actual_result_str,
                "Error/Details": error_details,
                "SQL_Query": sql_query # Include original query for full report
            })

        self.display_results_in_table()
        self.save_report_button.setEnabled(True)
        QMessageBox.information(self, "Validation Complete", "All test cases have been executed.")

    def display_results_in_table(self):
        self.report_table.setRowCount(len(self.validation_results))
        for row_idx, result in enumerate(self.validation_results):
            self.report_table.setItem(row_idx, 0, QTableWidgetItem(result["TC Name"]))
            self.report_table.setItem(row_idx, 1, QTableWidgetItem(result["Status"]))
            self.report_table.setItem(row_idx, 2, QTableWidgetItem(result["Expected Result"]))
            self.report_table.setItem(row_idx, 3, QTableWidgetItem(result["Actual Result"]))
            self.report_table.setItem(row_idx, 4, QTableWidgetItem(result["Error/Details"]))

            # Color code status
            if result["Status"] == "PASS":
                self.report_table.item(row_idx, 1).setBackground(Qt.green)
            elif result["Status"] == "FAIL":
                self.report_table.item(row_idx, 1).setBackground(Qt.red)
            elif result["Status"] == "ERROR":
                self.report_table.item(row_idx, 1).setBackground(Qt.darkRed)


    def save_report_to_excel(self):
        if not self.validation_results:
            QMessageBox.warning(self, "No Report", "No validation results to save.")
            return

        file_dialog = QFileDialog()
        file_dialog.setDefaultSuffix("xlsx")
        file_path, _ = file_dialog.getSaveFileName(self, "Save Validation Report", "", "Excel Files (*.xlsx)")

        if file_path:
            try:
                report_df = pd.DataFrame(self.validation_results)
                report_df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Report Saved", f"Validation report saved to '{file_path}'.")
            except Exception as e:
                QMessageBox.critical(self, "Error Saving Report", f"Could not save report: {e}")

    def clear_all(self):
        if self.db_conn:
            self.db_conn.close()
            self.db_conn = None
        self.data_files_loaded = {}
        self.test_cases_df = None
        self.validation_results = []

        self.loaded_data_files_list.clear()
        self.tc_file_path_label.setText("No test case file loaded.")
        self.view_tc_button.setEnabled(False)
        self.report_table.setRowCount(0)
        self.save_report_button.setEnabled(False)
        self.update_run_button_state()
        QMessageBox.information(self, "Cleared", "All loaded data and test cases have been cleared.")

    def run_manual_sql(self):
        sql = self.manual_sql_input.toPlainText().strip()
        self.manual_sql_result_table.clear()
        if not sql:
            QMessageBox.warning(self, "No SQL", "Please enter a SQL query.")
            return
        if not self.db_conn:
            QMessageBox.warning(self, "No Data", "No data loaded. Please load Excel data files first.")
            return
        try:
            cursor = self.db_conn.cursor()
            # Try to execute the whole input as a single statement (works for CTEs, UNIONs, etc.)
            cursor.execute(sql)
            if cursor.description:  # SELECT or similar
                rows = cursor.fetchall()
                columns = [desc[0] for desc in cursor.description]
                self.manual_sql_result_table.setColumnCount(len(columns))
                self.manual_sql_result_table.setHorizontalHeaderLabels(columns)
                self.manual_sql_result_table.setRowCount(len(rows))
                for i, row in enumerate(rows):
                    for j, value in enumerate(row):
                        self.manual_sql_result_table.setItem(i, j, QTableWidgetItem(str(value)))
            else:  # Non-SELECT (INSERT/UPDATE/DELETE)
                self.db_conn.commit()
                self.manual_sql_result_table.setColumnCount(1)
                self.manual_sql_result_table.setRowCount(1)
                self.manual_sql_result_table.setHorizontalHeaderLabels(["Result"])
                self.manual_sql_result_table.setItem(0, 0, QTableWidgetItem("Query executed successfully."))
        except Exception as e:
            self.manual_sql_result_table.setColumnCount(1)
            self.manual_sql_result_table.setRowCount(1)
            self.manual_sql_result_table.setHorizontalHeaderLabels(["Error"])
            self.manual_sql_result_table.setItem(0, 0, QTableWidgetItem(str(e)))

    def show_loaded_tables(self):
        if not self.db_conn:
            QMessageBox.information(self, "No DB", "No database loaded.")
            return
        cursor = self.db_conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = [row[0] for row in cursor.fetchall()]
        QMessageBox.information(self, "Loaded Tables", "\n".join(tables))

    def show_report_table_context_menu(self, pos):
        item = self.report_table.itemAt(pos)
        if item and item.column() == 0:  # Only for TC Name column
            row = item.row()
            menu = QMenu(self)
            copy_sql_action = menu.addAction("Copy SQL")
            action = menu.exec_(self.report_table.viewport().mapToGlobal(pos))
            if action == copy_sql_action:
                # Get SQL for this row from self.validation_results
                if 0 <= row < len(self.validation_results):
                    sql = self.validation_results[row].get("SQL_Query", "")
                    if sql:
                        QApplication.clipboard().setText(sql)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelSQLValidatorApp()
    ex.show()
    sys.exit(app.exec_())
