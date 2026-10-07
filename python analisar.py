import pandas as pd
from shapely import wkt
from shapely.geometry import Point

# ============================================================
# ARQUIVOS
# ============================================================

ARQUIVO_POSICOES = "Posições Pitanga.xlsx"
ARQUIVO_CERCA = "Mapa sem título- Carregamento Pitanga - Palmeirinha.csv"

ARQUIVO_SAIDA = "Tempo_dentro_cerca_Pitanga_Palmeirinha.xlsx"


# ============================================================
# 1 - LER ARQUIVOS
# ============================================================

print("Lendo relatório de posições...")

pos = pd.read_excel(ARQUIVO_POSICOES)

print("Lendo arquivo da cerca...")

cerca = pd.read_csv(
    ARQUIVO_CERCA,
    sep=None,
    engine="python"
)

print("\nColunas das posições:")
print(pos.columns.tolist())

print("\nColunas da cerca:")
print(cerca.columns.tolist())


# ============================================================
# 2 - LER A GEOMETRIA DA CERCA
# ============================================================

# O arquivo da cerca precisa possuir uma coluna chamada WKT
if "WKT" not in cerca.columns:
    raise Exception(
        "Não encontrei a coluna 'WKT' no arquivo da cerca."
    )

poligono = wkt.loads(cerca.iloc[0]["WKT"])


# ============================================================
# 3 - TRATAR DATA/HORA
# ============================================================

pos["Data/Hora"] = pd.to_datetime(
    pos["Data/Hora"],
    dayfirst=True,
    errors="coerce"
)

# Remove registros sem data
pos = pos.dropna(subset=["Data/Hora"])


# ============================================================
# 4 - VERIFICAR CADA POSIÇÃO DENTRO DA CERCA
# ============================================================

print("\nVerificando posições dentro da cerca...")

def verificar_cerca(row):

    try:
        ponto = Point(
            float(row["Longitude"]),
            float(row["Latitude"])
        )

        return poligono.covers(ponto)

    except:
        return False


pos["Dentro da cerca"] = pos.apply(
    verificar_cerca,
    axis=1
)


# ============================================================
# 5 - ORDENAR POR PREFIXO E HORÁRIO
# ============================================================

pos = pos.sort_values(
    ["Frota", "Data/Hora"]
).reset_index(drop=True)


# ============================================================
# 6 - PEGAR SOMENTE AS POSIÇÕES DENTRO DA CERCA
# ============================================================

dentro = pos[
    pos["Dentro da cerca"] == True
].copy()


print(
    f"\nTotal de posições analisadas: {len(pos)}"
)

print(
    f"Posições dentro da cerca: {len(dentro)}"
)


# ============================================================
# 7 - IDENTIFICAR CADA PERMANÊNCIA
# ============================================================

# O caminhão pode ter várias posições dentro da cerca.
#
# Quando existe um intervalo grande entre duas posições,
# consideramos que ele saiu e entrou novamente.
#
# Neste exemplo:
# até 2 minutos = mesma permanência
# mais de 2 minutos = nova permanência


resultado = []

for prefixo, grupo in dentro.groupby("Frota"):

    grupo = grupo.sort_values("Data/Hora").copy()

    grupo["Intervalo"] = (
        grupo["Data/Hora"]
        .diff()
        .dt.total_seconds()
        / 60
    )

    # Nova permanência quando o intervalo for maior que 2 minutos
    grupo["Nova permanencia"] = (
        grupo["Intervalo"] > 2
    ).fillna(True)

    grupo["ID_Permanencia"] = (
        grupo["Nova permanencia"]
        .cumsum()
    )

    # --------------------------------------------------------
    # CONSOLIDAR CADA PERMANÊNCIA
    # --------------------------------------------------------

    for id_perm, bloco in grupo.groupby("ID_Permanencia"):

        entrada = bloco["Data/Hora"].min()

        saida = bloco["Data/Hora"].max()

        tempo = saida - entrada

        resultado.append({

            "Prefixo": prefixo,

            "Data": entrada.date(),

            "Entrada na cerca": entrada,

            "Saída da cerca": saida,

            "Tempo dentro da cerca": tempo,

            "Minutos dentro": tempo.total_seconds() / 60,

            "Quantidade de posições": len(bloco),

            "Latitude entrada": bloco.iloc[0]["Latitude"],

            "Longitude entrada": bloco.iloc[0]["Longitude"],

            "Latitude saída": bloco.iloc[-1]["Latitude"],

            "Longitude saída": bloco.iloc[-1]["Longitude"]

        })


# ============================================================
# 8 - CRIAR RESULTADO
# ============================================================

resultado = pd.DataFrame(resultado)

resultado = resultado.sort_values(
    ["Prefixo", "Entrada na cerca"]
).reset_index(drop=True)


# ============================================================
# 9 - CRIAR COLUNA DE TEMPO EM HORAS/MINUTOS
# ============================================================

def formatar_tempo(td):

    segundos = int(td.total_seconds())

    horas = segundos // 3600

    minutos = (segundos % 3600) // 60

    segundos_restantes = segundos % 60

    return f"{horas:02d}:{minutos:02d}:{segundos_restantes:02d}"


resultado["Tempo HH:MM:SS"] = resultado[
    "Tempo dentro da cerca"
].apply(formatar_tempo)


# ============================================================
# 10 - CRIAR UMA ABA RESUMIDA POR CAMINHÃO
# ============================================================

resumo = (
    resultado
    .groupby("Prefixo")
    .agg(
        Entradas=("Entrada na cerca", "count"),
        Primeira_entrada=("Entrada na cerca", "min"),
        Última_saida=("Saída da cerca", "max"),
        Tempo_total_minutos=("Minutos dentro", "sum")
    )
    .reset_index()
)


resumo["Tempo total"] = pd.to_timedelta(
    resumo["Tempo_total_minutos"],
    unit="m"
)


# ============================================================
# 11 - SALVAR EXCEL
# ============================================================

print("\nGerando Excel...")

with pd.ExcelWriter(
    ARQUIVO_SAIDA,
    engine="openpyxl"
) as writer:

    # Aba principal
    resultado.to_excel(
        writer,
        sheet_name="Permanencias",
        index=False
    )

    # Resumo por caminhão
    resumo.to_excel(
        writer,
        sheet_name="Resumo por Prefixo",
        index=False
    )

    # Todas as posições classificadas
    pos.to_excel(
        writer,
        sheet_name="Posições",
        index=False
    )

    # Dados da cerca
    cerca.to_excel(
        writer,
        sheet_name="Cerca",
        index=False
    )


# ============================================================
# 12 - FINAL
# ============================================================

print("\n==========================================")
print("ANÁLISE CONCLUÍDA")
print("==========================================")

print(
    f"Arquivo gerado: {ARQUIVO_SAIDA}"
)

print(
    f"Total de permanências encontradas: {len(resultado)}"
)

print("\nPrimeiros resultados:")

print(
    resultado[
        [
            "Prefixo",
            "Entrada na cerca",
            "Saída da cerca",
            "Tempo HH:MM:SS"
        ]
    ].head(20)
)

print("\n==========================================")
