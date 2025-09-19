import mysql.connector
import pandas as pd
from openpyxl import load_workbook

def export_mysql_query_to_template():
    config = {
        'user': 'your_username',
        'password': 'your_password',
        'host': 'localhost',
        'database': 'your_database'
    }

    query = """
    SELECT
        column_name AS abc,
        column_name2 AS def
    FROM
        your_table
    WHERE
        some_condition
    """

    template_path = "./template.xls"       
    output_path = "./src/public/query_result_filled.xlsx"  

    cell_map = {
        'abc': 'B2',
        'def': 'C2',
    }

    try:
        conn = mysql.connector.connect(**config)
        df = pd.read_sql(query, conn)

        wb = load_workbook(template_path)
        ws = wb.active  

        for col_name, start_cell in cell_map.items():
            if col_name not in df.columns:
                print(f"⚠️ Column '{col_name}' not found in query results, skipping.")
                continue

            start_col = ws[start_cell].column  
            start_row = ws[start_cell].row    

            for i, value in enumerate(df[col_name], start=start_row):
                ws.cell(row=i, column=start_col, value=value)

        
        wb.save(output_path)
        print(f"Data written to template and saved as {output_path}")

    except mysql.connector.Error as err:
        print(f" MySQL error: {err}")

    except Exception as e:
        print(f" Unexpected error: {e}")

    finally:
        if 'conn' in locals() and conn.is_connected():
            conn.close()

if __name__ == "__main__":
    export_mysql_query_to_template()
