from pathlib import Path

SAVE_TO_FILES = True   # True = ukládat po krocích, False = předávat přes return


def maybe_save(df, output_path):
    """
    Uloží DataFrame pouze pokud je SAVE_TO_FILES = True
    """
    if SAVE_TO_FILES and output_path is not None:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(output_path, index=False)
