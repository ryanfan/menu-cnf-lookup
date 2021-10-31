import pandas as pd
import xlsxwriter.utility

import src.MicrosoftAccessService as MicrosoftAccessService
import src.MicrosoftExcelService as MicrosoftExcelService


class NutrientLookupService:

    def __init__(self):
        self.premade_ingredients_excel = None
        self.cnf_ingredients_database = None
        self.recipes_excel = None

    def add_premade_ingredients_excel(self, premade_ingredients_excel):
        self.premade_ingredients_excel = MicrosoftExcelService.MicrosoftExcelService(premade_ingredients_excel)

    def add_cnf_ingredients_database(self, cnf_ingredients_database):
        self.cnf_ingredients_database = MicrosoftAccessService.MicrosoftAccessService(cnf_ingredients_database)

    def read_recipes(self, recipes_excel):
        self.recipes_excel = MicrosoftExcelService.MicrosoftExcelService(recipes_excel)
        recipes_df = self.recipes_excel.read_excel()

        premades = MicrosoftExcelService.MicrosoftExcelService(r'data/1-Premade food nutrition.xlsx').read_excel()
        monica = MicrosoftExcelService.MicrosoftExcelService(r'data/4- Final Nutrition Analysis.xlsx').read_excel()

        unique_cnf_ingredients = self.recipes_excel.get_unique_values_from_column_name('FoodID')
        UNIQUE_CNF_INGREDIENTS_LIST = tuple(unique_cnf_ingredients.tolist())
        # looking at the recipes to see which are premade and which are from cnf
        NUTRIENT_IDS = tuple({
                                 208: "ENERGY (KILOCALORIES)",
                                 204: "FAT (TOTAL LIPIDS)",
                                 606: "FATTY ACIDS, SATURATED, TOTAL",
                                 605: "FATTY ACIDS, TRANS, TOTAL",
                                 601: "CHOLESTEROL",
                                 307: "SODIUM",
                                 205: "CARBOHYDRATE, TOTAL (BY DIFFERENCE)",
                                 291: "FIBRE, TOTAL DIETARY",
                                 269: "SUGARS, TOTAL",
                                 203: "PROTEIN",
                                 319: "RETINOL",
                                 401: "VITAMIN C",
                                 301: "CALCIUM",
                                 303: "IRON"}.keys()
                             )

        NUTRIENT_IDS_COLUMN_ORDER = [
            'ENERGY (KILOCALORIES)',
            'FAT (TOTAL LIPIDS)',
            'FATTY ACIDS, SATURATED, TOTAL',
            'FATTY ACIDS, TRANS, TOTAL',
            'CHOLESTEROL',
            'SODIUM',
            'CARBOHYDRATE, TOTAL (BY DIFFERENCE)',
            'FIBRE, TOTAL DIETARY',
            'SUGARS, TOTAL',
            'PROTEIN',
            'RETINOL',
            'VITAMIN C',
            'CALCIUM',
            'IRON'
        ]

        relevant_nutrient_info_transposed = self.cnf_ingredients_database.run_query(
            'SELECT * FROM ("Nutrient Amount" na ' +
            'INNER JOIN "Nutrient Name" nn ON na.NutrientID=nn.NutrientID)' +
            'WHERE na.FoodID IN ' + str(UNIQUE_CNF_INGREDIENTS_LIST) + ' ' +
            'AND na.NutrientID IN ' + str(NUTRIENT_IDS)
        ).pivot(index='FoodID', columns='NutrientName', values='NutrientValue')[NUTRIENT_IDS_COLUMN_ORDER]

        relevant_measurement_info = self.cnf_ingredients_database.run_query(
            'SELECT fn.FoodID, fn.FoodDescription, cfa.ConversionFactorValue, mn.MeasureID, mn.MeasureDescription FROM ("Foodname" fn ' +
            'INNER JOIN "ConvFactAmount" cfa ON fn.FoodID=cfa.FoodID)' +
            'INNER JOIN "MeasureName" mn ON mn.MeasureID=cfa.MeasureID ' +
            'WHERE fn.FoodID IN ' + str(UNIQUE_CNF_INGREDIENTS_LIST)
        )

        result = relevant_nutrient_info_transposed.merge(relevant_measurement_info, on="FoodID", how='left')

        relevant_nutrient_info_transposed = None
        relevant_measurement_info = None

        relevant_foodname_info = self.cnf_ingredients_database.run_query(
            'SELECT * FROM "Foodname" fa ' +
            'WHERE fa.FoodID IN ' + str(UNIQUE_CNF_INGREDIENTS_LIST)
        )

        result = relevant_foodname_info.merge(result, on="FoodID", how='left')[
            ['FoodID', 'FoodDescription_x', 'MeasureDescription', 'MeasureID',
             'ConversionFactorValue'] + NUTRIENT_IDS_COLUMN_ORDER]
        result = pd.concat([result, result.iloc[:, 5:19].mul(result['ConversionFactorValue'], 0)], axis=1)
        result["concat_id"] = result["FoodID"].astype(str) + ":" + result["MeasureDescription"].astype(str)
        mid = result['concat_id']
        result.drop(labels=['concat_id'], axis=1, inplace=True)
        result.insert(0, 'concat_id', mid)

        # print(result.loc[result['FoodID'] == 14]["MeasureDescription"].dropna().unique())
        tmp = []
        for index, row in recipes_df.iterrows():
            if pd.isna(row['FoodID']):
                tmp.append([])
                continue
            values = result.loc[result['FoodID'] == int(row['FoodID'])]["MeasureDescription"].dropna().unique()
            tmp.append(values)
        recipes_df['new_col'] = tmp


        with pd.ExcelWriter(r'data/2-Excel-Writer-Output.xlsx', engine='xlsxwriter') as writer:
            amounts, units, in_grams, ratios, premade_measurements, premade_units = [], [], [], [], [], []
            start_row = 2
            for index, row in recipes_df.iterrows():

                if not pd.isna(row['FoodID']):

                    if row['source'] == "CNF":
                        amounts.append("=VALUE(LEFT(J" + str(index + 2) + ", MAX(ISNUMBER(VALUE(MID(J" + str(
                            index + 2) + ",{1,2,3,4,5,6,7,8,9},1)))*{1,2,3,4,5,6,7,8,9})+1-1))")
                        units.append("=TRIM(RIGHT(J" + str(index + 2) + ", LEN(J" + str(
                            index + 2) + ") - MAX(ISNUMBER(VALUE(MID(J" + str(
                            index + 2) + ",{1,2,3,4,5,6,7,8,9},1)))*{1,2,3,4,5,6,7,8,9})))")
                        for idx, num in enumerate([5, 6]):
                            s = "=VLOOKUP(F" + str(index + 2) + "&\":\"&J" + str(
                                index + 2) + ",'CNF Data'!$B$1:$AI$" + str(len(result.index)) + "," + str(num) + ",FALSE)"
                            recipes_df.iloc[[index], [9 + idx]] = s
                        for idx, num in enumerate([21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34]):
                            s = "=VLOOKUP(F" + str(index + 2) + "&\":\"&J" + str(
                                index + 2) + ",'CNF Data'!$B$1:$AI$2" + str(len(result.index)) + "," + str(num) + ",FALSE) * P" + str(index + 2)
                            recipes_df.iloc[[index], [15 + idx]] = s
                        premade_measurements.append("")
                        premade_units.append("")
                        ratios.append(
                            "=C" + str(index + 2) + "/H" + str(index + 2)
                        )
                        in_grams.append(
                            "=C" + str(index + 2) + "*L" + str(index + 2) + "/H" + str(index + 2) + "* 100"
                        )

                    elif row['source'] == "Premade - needs CNF conversion":
                        amounts.append("=VALUE(LEFT(J" + str(index + 2) + ", MAX(ISNUMBER(VALUE(MID(J" + str(
                            index + 2) + ",{1,2,3,4,5,6,7,8,9},1)))*{1,2,3,4,5,6,7,8,9})+1-1))")
                        units.append("=TRIM(RIGHT(J" + str(index + 2) + ", LEN(J" + str(
                            index + 2) + ") - MAX(ISNUMBER(VALUE(MID(J" + str(
                            index + 2) + ",{1,2,3,4,5,6,7,8,9},1)))*{1,2,3,4,5,6,7,8,9})))")
                        for idx, num in enumerate([5, 6]):
                            s = "=VLOOKUP(F" + str(index + 2) + "&\":\"&J" + str(
                                index + 2) + ",'CNF Data'!$B$1:$AI$" + str(len(result.index)) + "," + str(num) + ",FALSE)"
                            recipes_df.iloc[[index], [9 + idx]] = s
                        for idx, num in enumerate([4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]):
                            s = "=VLOOKUP(E" + str(index + 2) + ", Premades!B1: R18, " + str(
                                num) + ", FALSE) * P" + str(index + 2)
                            recipes_df.iloc[[index], [15 + idx]] = s
                        ratios.append(
                            "=C" + str(index + 2) + "/N" + str(index + 2)
                        )
                        in_grams.append(
                            "=C" + str(index + 2) + "*L" + str(index + 2) + "/H" + str(index + 2) + "* 100"
                        )
                        premade_measurements.append(
                            "=VLOOKUP(E" + str(index + 2) + ", Premades!B1: R18, 2, FALSE)"
                        )
                        premade_units.append(
                            "=VLOOKUP(E" + str(index + 2) + ", Premades!B1: R18, 3, FALSE)"
                        )

                elif row['source'] == "Premade":
                    amounts.append("")
                    units.append("")
                    for idx, num in enumerate([4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]):
                        s = "=VLOOKUP(E" + str(index + 2) + ", Premades!B1: R18, " + str(
                            num) + ", FALSE) * P" + str(index + 2)
                        recipes_df.iloc[[index], [15 + idx]] = s
                    ratios.append(
                        "=C" + str(index + 2) + "/N" + str(index + 2)
                    )
                    in_grams.append(
                        "=C" + str(index + 2)
                    )
                    premade_measurements.append(
                        "=VLOOKUP(E" + str(index + 2) + ", Premades!B1: R18, 2, FALSE)"
                    )
                    premade_units.append(
                        "=VLOOKUP(E" + str(index + 2) + ", Premades!B1: R18, 3, FALSE)"
                    )

                elif row['name'] == "Recipe":
                    start_row = index
                    amounts.append("")
                    units.append("")
                    in_grams.append("")
                    ratios.append("")
                    premade_measurements.append("")
                    premade_units.append("")

                elif row['name'] == "Total":
                    amounts.append("")
                    units.append("")
                    in_grams.append("=SUM(M" + str(start_row + 2) + ":M" + str(index + 1) + ")")
                    ratios.append("=M" + str(index + 2) + "/C" + str(index + 2))
                    premade_measurements.append("")
                    premade_units.append("")
                    for idx, num in enumerate([4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]):
                        col = xlsxwriter.utility.xl_col_to_name(15 + idx + 1)
                        s = "=SUM(" + col + str(start_row + 2) + ":" + col + str(index + 1) + ") / P" + str(index + 2)
                        recipes_df.iloc[[index], [15 + idx]] = s

                else:
                    amounts.append("")
                    units.append("")
                    in_grams.append("")
                    ratios.append("")
                    premade_measurements.append("")
                    premade_units.append("")

            recipes_df['amount'] = amounts
            recipes_df['unit'] = units
            recipes_df['in grams'] = in_grams
            recipes_df['ratio'] = ratios
            recipes_df['premade measurement'] = premade_measurements
            recipes_df['premade unit'] = premade_units

            recipes_df.drop(['new_col'], axis=1).to_excel(writer, sheet_name="Recipes")
            result.to_excel(writer, sheet_name="CNF Data")
            premades.to_excel(writer, sheet_name="Premades")
            monica.to_excel(writer, sheet_name="Monica")

            # Assign the workbook and worksheet
            worksheet = writer.sheets['Recipes']

            # Adding the header and Datavalidation list
            for index, row in recipes_df.iterrows():
                print(row['new_col'])
                if not pd.isna(row['FoodID']):
                    worksheet.data_validation('J' + str(index + 2), {'validate': 'list',
                                                                     'source': list(row['new_col'])})
