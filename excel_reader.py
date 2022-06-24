import pandas as pd


class rdp_connection:
    def __init__(self, host, username, password):
        self.host = host
        self.username = username
        self.password = password

    def host_toString(self):
        return f"Host: {self.host}"

    def username_toString(self):
        return f"Username: {self.username}"

    def password_toString(self):
        return f"Password: {self.password}"

    def divider_toString(self, div="-"):
        return div * max(len(self.host_toString()), len(self.username_toString()), len(self.password_toString()))

    def __str__(self):
        divider = self.divider_toString()
        return divider + "\n" + self.host_toString() + "\n" + self.username_toString() + "\n" + self.password_toString() + "\n" + divider


class excel_data:
    def __init__(self, excel_path=None):
        if excel_path is None:
            print("No Path given")
            exit(0)
        self.rdp_connections = []

        df = pd.read_excel(excel_path)
        for row in df.iterrows():
            row_values = row[1].values
            if str(row_values[0]) and str(row_values[1]) != "nan" and str(row_values[2]) != "nan":
                self.rdp_connections.append(
                    rdp_connection(host=row_values[0], username=row_values[1], password=row_values[2]))

    def getUser(self, index=0):
        if index >= len(self.rdp_connections):
            index = 0
        return self.rdp_connections[index]
