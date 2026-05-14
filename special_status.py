H3_A_POS = {
    "44380095", "44378704", "44384794", "44386329"
}
H3_B_POS = {
    "44381936", "44383773", "44382980",
    "44381399", "44384047", "44383393", "44382187"
}

# H2:
#   A: {44380095, 44378704, 44384794, 44386329, 44382187} = ALT
#   B: {44381936, 44383773, 44382980, 44381399, 44384047, 44383393} = REF
H2_A_POS = {
    "44380095", "44378704", "44384794", "44386329", "44382187"
}
H2_B_POS = {
    "44381936", "44383773", "44382980",
    "44381399", "44384047", "44383393"
}

# H1_Ref:
#   všechny tyto pozice jsou REF
H1_ALL_REF_POS = {
    "44380095", "44378704", "44384794", "44386329",
    "44381936", "44383773", "44382980",
    "44381399", "44384047", "44383393", "44382187"
}

# Jednoduchá speciální pozice pro highCd / lowCd
SPECIAL_SINGLE = {
    # pokud 4966222 je REF -> highCd, pokud ALT -> lowCd
    "4966222": {"REF": "highCd", "ALT": "lowCd"},
}


def _all_in_state(pos_results, positions, expected):
    """
    pos_results: dict pozice -> "REF"/"ALT"
    positions: množina pozic, které nás zajímají
    expected: "REF" nebo "ALT"

    Bere v úvahu jen ty pozice, které v pos_results skutečně jsou.
    Vrací True jen pokud:
      - máme aspoň jednu takovou pozici
      - a všechny jsou v požadovaném stavu.
    """
    vals = [pos_results[p] for p in positions if p in pos_results]
    if not vals:
        return False
    return all(v == expected for v in vals)


def eval_special_h3_h2(pos_results):
    """
    Vyhodnocuje H3 / H2 / H1_Ref na základě slovníku:
        pos_results = { "pozice": "REF" nebo "ALT", ... }

    Pravidla:
      - H3_delays(P):
          H3_A_POS = REF a H3_B_POS = ALT
      - H2_speedsup(P):
          H2_A_POS = ALT a H2_B_POS = REF
      - H1_Ref:
          H1_ALL_REF_POS = REF

    Vrací:
      - "H3_delays(P)" nebo "H2_speedsup(P)" nebo "H1_Ref"
      - None, pokud se nic netrefí.
    """

    # H3: A = REF, B = ALT
    if _all_in_state(pos_results, H3_A_POS, "REF") and _all_in_state(pos_results, H3_B_POS, "ALT"):
        return "H3_delays(P)"

    # H2: A = ALT, B = REF
    if _all_in_state(pos_results, H2_A_POS, "ALT") and _all_in_state(pos_results, H2_B_POS, "REF"):
        return "H2_speedsup(P)"

    # H1_Ref: všechny vyjmenované pozice jsou REF
    if _all_in_state(pos_results, H1_ALL_REF_POS, "REF"):
        return "H1_Ref"

    return None


def eval_special_single(pos_key, value_ref_alt):
    """
    Vyhodnocuje jednoduchá speciální pravidla jako highCd/lowCd.

    Parametry:
        pos_key        – pozice jako string (např. "4966222")
        value_ref_alt  – "REF" nebo "ALT"

    Vrátí:
        - speciální status jako string ("highCd" / "lowCd"),
        - nebo None pokud se neaplikuje.
    """
    if pos_key in SPECIAL_SINGLE:
        mapping = SPECIAL_SINGLE[pos_key]
        return mapping.get(value_ref_alt, None)

    return None
