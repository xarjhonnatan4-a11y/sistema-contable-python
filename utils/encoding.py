import pandas as pd


MARCAS_MOJIBAKE = (
    "Ã",
    "Â",
    "�",
    "â€",
    "â€œ",
    "â€˜",
    "â„",
    "â†",
    "âœ",
    "âš",
    "ðŸ",
)


def _parece_mojibake(texto):
    return any(marca in texto for marca in MARCAS_MOJIBAKE)


def corregir_mojibake(texto):
    """Repara texto UTF-8 que fue interpretado como Windows-1252/Latin-1."""
    if not isinstance(texto, str):
        return texto

    resultado = texto.strip()
    for _ in range(3):
        if not _parece_mojibake(resultado):
            break

        reparado = resultado
        for encoding in ("cp1252", "latin1"):
            try:
                candidato = resultado.encode(encoding).decode("utf-8")
            except UnicodeError:
                continue
            reparado = candidato
            break

        if reparado == resultado:
            break
        resultado = reparado.strip()

    return resultado


def normalizar_columnas_dataframe(df):
    """Limpia espacios y repara mojibake en los encabezados de un DataFrame."""
    df = df.copy()
    df.columns = [corregir_mojibake(str(col)) for col in df.columns]

    if not df.columns.duplicated().any():
        return df

    normalizado = pd.DataFrame(index=df.index)
    for columna in dict.fromkeys(df.columns):
        repetidas = df.loc[:, df.columns == columna]
        if repetidas.shape[1] == 1:
            normalizado[columna] = repetidas.iloc[:, 0]
        else:
            normalizado[columna] = repetidas.bfill(axis=1).iloc[:, 0]

    return normalizado
