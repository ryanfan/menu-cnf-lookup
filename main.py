import src.NutrientLookupService as NutrientLookupService

if __name__ == '__main__':
    premade_ingredients_excel = r'data/1-Premade food nutrition.xlsx'
    cnf_ingredients_database = r'data/CNF2015.accdb'
    recipes_excel = r'data/2-Recipes-input.xlsx'

    # Declare output file
    nutrient_lookup = NutrientLookupService.NutrientLookupService()

    # Add premade ingredients
    nutrient_lookup.add_premade_ingredients_excel(premade_ingredients_excel)

    # Add CNF ingredients
    nutrient_lookup.add_cnf_ingredients_database(cnf_ingredients_database)

    nutrient_lookup.read_recipes(recipes_excel)

    # Merge data

    # Save output

    # output file.
