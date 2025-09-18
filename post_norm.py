import pandas as pd
from tkinter import filedialog
import copy

# CONSTANTS
TERM_DICT_TEMPLATE = {"termo": None, "classe": None, "subclasse": None, "in_vocabulum?": None, "justification":None}
TERM_COLUMN_NAMES = ["novo_nome", "novo_assunto", "novo_local", "novo_descrel"]

def set_justification(termo, classe):
    output = "AVERIGUAR"

    if classe == "descrel": output = "DESCREL"
    if ("[" or "OU") in str(termo): output = "INDETERMINACAO"
    if classe == "entidade coletiva": output = "ENTIDADE COLETIVA"

    return output

# VARIABLES
term_output = []

if __name__ == "__main__":
    # Load files
    df_prenorm = pd.read_excel(filedialog.askopenfilename(), sheet_name="normalizacao")
    df_vocabulum = pd.read_excel(filedialog.askopenfilename())

    for index, row in df_prenorm.iterrows():
        for column_name in TERM_COLUMN_NAMES:
            cell_content = row[column_name]

            if pd.isna(cell_content):
                continue

            terms = str(cell_content).split(";")

            for term in terms:
                term = term.strip()
                if not term:
                    continue

                temp_dict = copy.deepcopy(TERM_DICT_TEMPLATE)
                temp_dict["termo"] = term
                temp_dict["classe"] = column_name.replace("novo_", "")
                temp_dict["subclasse"] = row["subtipo"]
                if term in df_vocabulum["termo"].values:
                    temp_dict["in_vocabulum?"] = "SIM"
                    temp_dict["justification"] = "--"
                else:
                    temp_dict["in_vocabulum?"] = "NÃO"
                    temp_dict["justification"] = set_justification(temp_dict["termo"], temp_dict["classe"])

                term_output.append(temp_dict)

    df_output = pd.DataFrame(term_output)

    df_output.to_excel("postnorm_check.xlsx", index=False)
