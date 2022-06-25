import pandas as pd


class rdp_connection:
    def __init__(self, host, username, password):
        self.host = host
        self.username = username
        self.password = password


class excel_data:
    def __init__(self, excel_path):
        self.rdp_connections = []
        df = pd.read_excel(excel_path)
        for row in df.iterrows():
            row_values = row[1].values
            if str(row_values[0]) and str(row_values[1]) != "nan" and str(row_values[2]) != "nan":
                self.rdp_connections.append(
                    rdp_connection(host=row_values[0], username=row_values[1], password=row_values[2]))
