import pandas as pd
from tkinter import filedialog
import copy

# CONSTANTS
TERM_DICT_TEMPLATE = {"termo": None, "classe": None, "subclasse": None, "in_vocabulum?": None}
TERM_COLUMN_NAMES = ["novo_nome", "novo_assunto", "novo_local", "novo_descrel"]

# VARIABLES
term_output = []

if __name__ == "__main__":
    # Load files
    df_prenorm = pd.read_excel(filedialog.askopenfilename(), sheet_name="normalizacao")
    df_vocabulum = pd.read_excel(filedialog.askopenfilename())

    for index, row in df_prenorm.iterrows():
        for column_name in TERM_COLUMN_NAMES:
            cell_content = row[column_name]

            # skip empty / NaN cells
            if pd.isna(cell_content):
                continue

            # split by ";" into terms
            terms = str(cell_content).split(";")

            for term in terms:
                term = term.strip()
                if not term:
                    continue

                # create a fresh dict
                temp_dict = copy.deepcopy(TERM_DICT_TEMPLATE)
                temp_dict["termo"] = term
                temp_dict["classe"] = column_name.replace("novo_", "")
                temp_dict["subclasse"] = row["subtipo"]
                temp_dict["in_vocabulum?"] = "SIM" if term in df_vocabulum["termo"].values else "NÃO"

                # add to results
                term_output.append(temp_dict)

    # build DataFrame directly
    df_output = pd.DataFrame(term_output)

    df_output.to_excel("posnorm_check.xlsx", index=False)
