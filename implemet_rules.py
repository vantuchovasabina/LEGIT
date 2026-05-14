import pandas as pd
from pathlib import Path
from gene_rules import GENE_RULES  # musí být dostupné v PYTHONPATH

def label_excel_by_gene_rules(input_xlsx: str, output_xlsx: str | None = None) -> str:
    """
    V každém listu nahradí v těch sloupcích, jejichž hlavička (Variant_position)
    odpovídá klíči v GENE_RULES:
      - 'ref'  -> levá hodnota z GENE_RULES
      - 'alt'  -> pravá hodnota z GENE_RULES
      - 'Het', 'N' zůstávají
    Ostatní sloupce/listy ponechá beze změny.

    Vrací cestu k výstupnímu XLSX.
    """
    in_path = Path(input_xlsx)
    if not in_path.exists():
        raise FileNotFoundError(f"Soubor nenalezen: {in_path}")

    if output_xlsx is None:
        output_xlsx = str(in_path.with_name(in_path.stem + "_labeled.xlsx"))

    xf = pd.ExcelFile(in_path)

    def _to_key(x) -> str | None:
        """
        Z hlavičky sloupce vytáhne 'Variant_position' (pokud je MultiIndex s názvem úrovně),
        jinak použije samotnou hlavičku; vždy vrátí string bez whitespace.
        """
        # MultiIndex sloupec: tuple
        if isinstance(x, tuple):
            # Zkus najít úroveň pojmenovanou 'Variant_position'
            # (pandas ji předává skrz names u MultiIndex-u; k dispozici tady není,
            # proto volíme robustní heuristiku: vyber číslo/řetězec, který vypadá jako klíč)
            for part in x[::-1]:  # často bývá pozice poslední
                if part is None:
                    continue
                s = str(part).strip()
                if s != "":
                    return s
            return None

        # Jednoduchý název sloupce
        if x is None:
            return None
        return str(x).strip()

    def _map_cell(val, rule_pair):
        """Val (buněčná hodnota) → nahrazený štítek podle (left,right)."""
        if pd.isna(val):
            return val
        s = str(val).strip().lower()
        if s == "ref":
            return rule_pair[0]
        if s == "alt":
            return rule_pair[1]
        if s in ("het", "n"):
            return val  # ponecháme jak je (případně můžeš vrátit s původním 'Het'/'N')
        return val  # nic neměníme (např. jiné texty)

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        for sheet in xf.sheet_names:
            df = pd.read_excel(in_path, sheet_name=sheet, dtype=object)

            # Pro každý sloupec zjistíme, zda jeho "position key" je v GENE_RULES
            replace_cols = {}
            for col in df.columns:
                key = _to_key(col)
                if key is None:
                    continue
                # sjednotíme na string přesně jako v GENE_RULES (většinou bez mezer)
                if key in GENE_RULES:
                    replace_cols[col] = GENE_RULES[key]

            # Proveď mapování jen u sloupců, které mají pravidlo
            if replace_cols:
                df = df.copy()
                for col, rule_pair in replace_cols.items():
                    df[col] = df[col].map(lambda v: _map_cell(v, rule_pair))

            # Zapiš list
            df.to_excel(writer, sheet_name=sheet, index=False)

    return str(Path(output_xlsx).resolve())

