import pandas as pd
import os
from pathlib import Path

def edit_data_file(genotypes_data, trait_specific_data):
    "Funkce načte genotypová data uživatele a referenční panel markerů, "
    "provede jejich formátování a"
    " uloží upravené soubory ve sjednocené podobě pro další analýzu."
    # Cesta ke složce transformers
    # folder = "transformers"

    # # Zajištění, že složka existuje
    # os.makedirs(folder, exist_ok=True)

    # # Výstupní cesty
    # output_genotypes = os.path.join(folder, "Genotypes_targeted_updated.xlsx")
    # output_trait_specific = os.path.join(folder, "Trait_specific_updated.xlsx")

    folder = Path("transform_data")
    folder.mkdir(parents=True, exist_ok=True)

    output_genotypes = folder / "Genotypes_targeted_updated.xlsx"
    output_trait_specific = folder / "Trait_specific_updated.xlsx"

    # Načtení dat
    df_1 = pd.read_excel(genotypes_data)
    df_2 = pd.read_excel(trait_specific_data)

    # Vytvoření nového sloupce "Ref/Alt_Allele"
    df_1["Ref/Alt_Allele"] = df_1["Reference allele"].astype(str) + "/" + df_1["Alternate allele(s)"].astype(str)

    df_1["Reference"] = (
    # Očištění sloupce "Reference"(jen na číslo)
        df_1["Reference"]
        .astype(str)
        .str.extract(r"(\d+|[XYM])")[0]
        .apply(lambda x: str(int(x)) if x.isdigit() else x)
    )

    # Úprava: nahradit "-" za "."
    df_2["Ref/Alt_Allele"] = df_2["Ref/Alt_Allele"].astype(str).str.replace("-", ".", regex=False)

    # Uložení do složky transformers
    df_1.to_excel(output_genotypes, index=False)
    df_2.to_excel(output_trait_specific, index=False)

    print(f"Hotovo! Vytvořen soubor: {os.path.abspath(output_genotypes)}")
    print(f"Hotovo! Vytvořen soubor: {os.path.abspath(output_trait_specific)}")
