import pyodbc
import pandas as pd


class MicrosoftAccessService:

    def __init__(self, file_location):
        connection_string = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + file_location
        self.conn = pyodbc.connect(connection_string)
        self.cursor = self.conn.cursor()
        print("Done initializing MicrosoftAccessService with " + file_location)

    def get_conn(self):
        return self.conn

    def get_cursor(self):
        return self.cursor

    def run_query(self, query):
        return pd.read_sql(query, self.conn)
