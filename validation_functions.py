def check_row_count(db_conn, table_name):
    cursor = db_conn.cursor()
    cursor.execute(f'SELECT COUNT(*) FROM "{table_name}"')
    return str(cursor.fetchone()[0])

def always_pass():
    return "PASS"

def custom_logic_example(arg1, arg2):
    # Any logic you want
    return str(int(arg1) + int(arg2))