import pandas as pd
import re
from pathlib import Path

def split_dataset(Same_positin_input):
    "Funkce zpracuje soubor Trait_specific_updated a "
    "automaticky rozdělí data do odpovídajících Excel listů na základě typu traitu"
    OUTPUT_FILE = "Same_Position_Split_22.xlsx"   
    TRAIT_COL = "Trait"                           

    out_folder = Path("distribution")
    out_folder.mkdir(parents=True, exist_ok=True)

    out_path = out_folder / OUTPUT_FILE

    # Pomocná funkce
    def sanitize_sheet_name(name: str) -> str:
        """
        Vyčistí název listu pro Excel:
        - odstraní nepovolené znaky: : \\ / ? * [ ]
        - ořízne na max. 31 znaků (limit Excelu)
        - nahradí prázdné/NaN názvy za 'No_Trait'
        """
        if pd.isna(name) or str(name).strip() == "":
            cleaned = "No_Trait"
        else:
            cleaned = str(name).strip()
            cleaned = re.sub(r"[:\\/?*\[\]]", "_", cleaned)
            cleaned = cleaned[:31]  # Excel limit
            if cleaned == "":
                cleaned = "No_Trait"
        return cleaned

    df = pd.read_excel(Same_positin_input)

    if TRAIT_COL not in df.columns:
        raise ValueError(f"Ve vstupním souboru chybí sloupec '{TRAIT_COL}'.")

    # Úpravy traitů
    df_grouped = df.copy()

    # 1) SPECIÁLNÍ ZACHÁZENÍ PRO Lipoxygenase*
    mask_lipo = df_grouped[TRAIT_COL].astype(str).str.startswith("Lipoxygenase", na=False)
    df_grouped.loc[mask_lipo, TRAIT_COL] = "Lipoxygenase"

    # 2) SPOJENÍ KONKRÉTNÍCH TRAITŮ DO JEDNOHO (přes mapu)
    trait_map = {
        # Pubsecence_color
        "Pubsecence_color_light_tawny/tawny": "Pubsecence_color",
        "Pubsecence_color_tawny/grey": "Pubsecence_color",

        # Pod_color
        "Pod_color_Black/Brown": "Pod_color",
        "Pod_color_Brown/Tan": "Pod_color",

        # Protein_Oil
        "Big_seed/Protein/Oil": "Protein_Oil",
        "Protein/Oil": "Protein_Oil",
        "Protein_oil_content": "Protein_Oil",
        "Fatty_acid_oil": "Protein_Oil",
        
        #Internode_length
        "Internode_length": "Internode_length",
        "Short_internode": "Internode_length",
    }
    df_grouped[TRAIT_COL] = df_grouped[TRAIT_COL].replace(trait_map)

    # === SKUPINOVÁNÍ A ZÁPIS DO VÍCE LISTŮ ===
    groups = df_grouped.groupby(TRAIT_COL, dropna=False)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        def sort_key(val):
            return ("zzz" if pd.isna(val) else str(val).lower())

        for trait_value, subdf in sorted(groups, key=lambda x: sort_key(x[0])):
            sheet_name = sanitize_sheet_name(trait_value)
            subdf.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Hotovo! Vytvořen soubor: {out_path.resolve()}")

