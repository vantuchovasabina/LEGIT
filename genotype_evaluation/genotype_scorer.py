import pandas as pd 
from gene_rules import GENE_RULES
from pathlib import Path

# Použijeme jen podskupinu "rules" pro tento skript (matrix)
# if isinstance(GENE_RULES, dict) and "rules" in GENE_RULES:
#     RULES_FOR_MATRIX = GENE_RULES["rules"]
# else:
#     # fallback, kdyby někdy GENE_RULES bylo přímo plochý slovník
#     RULES_FOR_MATRIX = GENE_RULES
# Použijeme jen podskupinu "rules" pro tento skript (matrix)
RULES_FOR_MATRIX = GENE_RULES["rules"]



def create_genotype_matrix(GENO_FILE, SPLIT_FILE):
    "Funkce create_genotype_matrix načte genotypová data vzorků a rozdělený soubor traitů, "
    "propojí je na základě chromozomu a pozice varianty a vytvoří matici genotypů pro jednotlivé traity. "
    "Pro každý marker vyhodnocuje genotypy vzorků vůči referenčním a alternativním alelám podle definovaných pravidel a "
    "výsledky ukládá do Excel souboru s oddělenými listy pro každý trait."
    OUT_FILE = "All_Traits_matrix.xlsx"

    out_folder = Path("genotype_evaluation")
    out_folder.mkdir(parents=True, exist_ok=True)
    out_path = out_folder / OUT_FILE

    # === Sloupce (názvy) ===
    MARKER_COL    = "Suggested_marker_name_(Wm82_REF_first_OR_major_minor_alleles)"
    GENE_NAME_COL = "Gene_name"
    MARKER_ID_COL = "Gene/marker_ID"
    ID_COL        = "ID"
    GENOTYPE_COL  = "Ref/Alt_Allele"  # BEREME z GENO_FILE

    # Ve traitu už nevyžadujeme Ref/Alt_Allele
    REQ_TRAIT_COLS = {
        "Chromosome", "Variant_position", MARKER_COL, GENE_NAME_COL, MARKER_ID_COL
    }

    # ---------- Pomocné parsování ---------- #
    def _parse_two(val):
        """
        Zkusí zparsovat něco jako 'A/T', 'AA/TT', 'AT' apod. na dvě alely.
        Používá se hlavně pro genotypy (buňky vzorků).
        """
        if val is None or pd.isna(val):
            return None, None
        v = str(val).strip()
        if v == "":
            return None, None
        for sep in ("/", "|"):
            if sep in v:
                a, b = [t.strip().upper() for t in v.split(sep, 1)]
                return a, b
        if len(v) == 2 and v.isalpha():
            return v[0].upper(), v[1].upper()
        return None, None

    def parse_refalt_multi(val):
        """
        Parsuje Ref/Alt_Allele, kde může být více alt alel oddělených čárkou.

        Příklady:
            'A/T' -> ('A', ['T'])
            'AAACAAC/A,AAAC' -> ('AAACAAC', ['A', 'AAAC'])
        """
        if val is None or pd.isna(val):
            return None, []

        s = str(val).strip()
        if s == "":
            return None, []

        # forma Ref/Alt1,Alt2,...
        if "/" in s:
            ref_part, alts_part = s.split("/", 1)
            ref = ref_part.strip().upper()
            alts = [a.strip().upper() for a in alts_part.split(",") if a.strip() != ""]
            return ref, alts

        # fallback: zkus původní _parse_two
        ra, rb = _parse_two(val)
        if ra is not None and rb is not None:
            return ra, [rb]

        return None, []

    # ---------- Vyhodnocení buňky vůči Ref/Alt_Allele ---------- #
    def ref_or_alt(genostr, refalt):
        """
        Vrací:
        - 'ALT2' pokud genotyp obsahuje '*'
        - 'N'    pokud genotyp obsahuje '.'
        - 'Het'  pokud genotyp je heterozygot (dvě různé alely)
        - 'ref'  pokud je homozygot referenční alely
        - 'alt'  pokud je homozygot jedné z alternativních alel
        - None   pokud nejde rozumně vyhodnotit
        """

        # SPECIÁLNÍ PRAVIDLO: hvězdička → ALT2
        if isinstance(genostr, str) and "*" in genostr:
            return "ALT2"

        # SPECIÁLNÍ PRAVIDLO: tečka v genotypu → N
        if isinstance(genostr, str) and "." in genostr:
            return "N"

        a, b = _parse_two(genostr)
        if a == "." or b == ".":  # kryje případy jako "A/." nebo "./T"
            return "N"

        # Parsuj Ref/Alt_Allele (může mít víc alt alel)
        ref, alts = parse_refalt_multi(refalt)

        # Speciální pravidlo: pokud je v Ref/Alt_Allele tečka jako jedna z alel, vše je 'ref'
        if ref == "." or any(alt == "." for alt in alts):
            return "ref"

        if a is None or b is None:
            return None  # neparsovatelná/missing buňka

        # Heterozygot: dvě různé alely
        if a != b:
            return "Het"

        # Homozygot: obě alely stejné
        # ref/ref
        if a == ref:
            return "ref"

        # alt/alt (pro kterýkoli z altů)
        if a in alts:
            return "alt"

        return None

    def labels_from_rules(gene_name, variant_position, marker_id):
        """
        Vrátí (ref_text, alt_text) pro danou pozici / marker_id
        podle RULES_FOR_MATRIX (= jen podskupina 'rules' z gene_rules.py).

        Pokud nic nenajde, použije gene_name jako základ:
            ref = gene_name
            alt = gene_name.lower()
        """
        key_pos = str(variant_position).strip()
        if key_pos in RULES_FOR_MATRIX:
            return RULES_FOR_MATRIX[key_pos]

        key_mid = str(marker_id).strip()
        if key_mid in RULES_FOR_MATRIX:
            return RULES_FOR_MATRIX[key_mid]

        return (str(gene_name), str(gene_name).lower())

    # === Načtení genotypů ===
    geno = pd.read_excel(GENO_FILE)
    # zachováváš původní rozmezí sloupců se vzorky
    sample_cols = list(geno.columns[6:21])
    samples = sample_cols

    # Kontrola, že GENO_FILE má Ref/Alt_Allele
    if GENOTYPE_COL not in geno.columns:
        raise ValueError(f"Ve vstupu GENO_FILE chybí sloupec '{GENOTYPE_COL}'.")

    geno = geno.copy()

    # DEDUPE: pokud existují vícenásobné řádky pro stejný (Reference, Position),
    # bereme první výskyt
    geno_nodup = (
        geno
        .sort_values(["Reference", "Position"])
        .drop_duplicates(subset=["Reference", "Position"], keep="first")
    )

    xf = pd.ExcelFile(SPLIT_FILE)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name in xf.sheet_names:
            trait = pd.read_excel(SPLIT_FILE, sheet_name=sheet_name)

            missing = REQ_TRAIT_COLS - set(trait.columns)
            if missing:
                pd.DataFrame(
                    {"info": [f"List '{sheet_name}': chybí sloupce {sorted(missing)}"]}
                ).to_excel(writer, sheet_name=sheet_name, index=False)
                continue

            trait = trait.copy()
            trait[MARKER_COL]    = trait[MARKER_COL].astype(str)
            trait[GENE_NAME_COL] = trait[GENE_NAME_COL].astype(str)
            trait[MARKER_ID_COL] = trait[MARKER_ID_COL].astype(str).str.strip()

            # Řádky, které budeme párovat
            rows = trait.dropna(subset=[MARKER_COL]).copy()
            rows = rows[["Chromosome", "Variant_position", MARKER_COL, GENE_NAME_COL, MARKER_ID_COL]]

            # Merge podle (Chromosome, Variant_position) <-> (Reference, Position)
            merged = pd.merge(
                rows,
                geno_nodup[["Reference", "Position", GENOTYPE_COL] + sample_cols],
                left_on=["Chromosome", "Variant_position"],
                right_on=["Reference", "Position"],
                how="left",
                validate="m:1"
            )

            # unikátní podle varianty: (Chromosome, Variant_position)
            merged_unique = (
                merged
                .drop_duplicates(subset=["Chromosome", "Variant_position"], keep="first")
                .set_index(["Chromosome", "Variant_position"])
            )

            # --- MultiIndex sloupce (jen štítky) --- #
            cols, col_to_key = [], []
            for _, r in rows.iterrows():
                gene_name = r[GENE_NAME_COL]
                chr_val   = r["Chromosome"]
                marker_id = r[MARKER_ID_COL]
                pos_val   = r["Variant_position"]

                ref_text, alt_text = labels_from_rules(gene_name, pos_val, marker_id)
                cols.append((gene_name, chr_val, marker_id, ref_text, alt_text, pos_val))
                col_to_key.append((chr_val, pos_val))

            if not cols:
                pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
                continue

            multi_cols = pd.MultiIndex.from_tuples(
                cols,
                names=[GENE_NAME_COL, "Chromosome", "Gene/marker_ID", "ref", "alt", "Variant_position"]
            )
            matrix = pd.DataFrame(index=samples, columns=multi_cols, dtype=object)

            # --- Naplnění matice --- #
            for col_key, key_tuple in zip(multi_cols, col_to_key):
                gene_name     = col_key[0]
                marker_id_val = col_key[2]
                varpos_val    = col_key[5]

                if key_tuple not in merged_unique.index:
                    continue

                row = merged_unique.loc[key_tuple]
                ref_text, alt_text = labels_from_rules(gene_name, varpos_val, marker_id_val)

                # Ref/Alt_Allele teď bereme z GENO_FILE (vpravo)
                refalt_value = row.get(GENOTYPE_COL, pd.NA)

                for sample in samples:
                    status = ref_or_alt(row.get(sample, pd.NA), refalt_value)

                    if status == "ref":
                        matrix.loc[sample, col_key] = ref_text
                    elif status == "alt":
                        matrix.loc[sample, col_key] = alt_text
                    elif status == "Het":
                        matrix.loc[sample, col_key] = "Het"
                    elif status == "N":
                        matrix.loc[sample, col_key] = "N"
                    elif status == "ALT2":
                        matrix.loc[sample, col_key] = "ALT2"
                    else:
                        matrix.loc[sample, col_key] = pd.NA

            matrix = matrix.sort_index(
                axis=1,
                level=[GENE_NAME_COL, "Chromosome", "Gene/marker_ID", "ref", "alt", "Variant_position"],
                sort_remaining=False
            )

            matrix.to_excel(writer, sheet_name=sheet_name, index=True, merge_cells=True)

    print(f"Hotovo! Vytvořen soubor: {out_path}")

