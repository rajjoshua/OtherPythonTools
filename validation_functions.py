def check_row_count(db_conn, table_name):
    cursor = db_conn.cursor()
    cursor.execute(f'SELECT COUNT(*) FROM "{table_name}"')
    return str(cursor.fetchone()[0])

def always_pass():
    return "PASS"

def custom_logic_example(arg1, arg2):
    # Any logic you want
    return str(int(arg1) + int(arg2))

import numpy as np
from PyQt5.QtWidgets import QMessageBox

def sum_array(arr):
    # arr should be a list of numbers
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setWindowTitle("Numpy Sum")
    msg.setText(f"Sum of array: {np.sum(arr)}")
    msg.exec_()
    return str(np.sum(arr))