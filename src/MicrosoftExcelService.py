import pandas as pd


class MicrosoftExcelService:

    def __init__(self, file_location):
        self.file_location = file_location
        self.df = pd.read_excel(self.file_location)
        print("Done initializing MicrosoftExcelService with " + file_location)

    def read_excel(self):
        return self.df

    def get_unique_values_from_column_name(self, column_name):
        return self.df[column_name].dropna().unique()