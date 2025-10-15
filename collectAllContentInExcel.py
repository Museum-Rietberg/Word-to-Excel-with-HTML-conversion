import pandas as pd

# Datei laden
file_path = "WIP/WIP--Japan_de_Luxe_Audio-Cult_INHALTE.xlsx"
xls = pd.ExcelFile(file_path, engine="openpyxl")

# Arbeitsblätter laden
wip = xls.parse("WIP Tracks")
de = xls.parse("DE Content")
fr = xls.parse("FR Content")
en = xls.parse("EN Content")

# Indexe setzen (nur eindeutige Zeilen behalten)
de_index = de.drop_duplicates("Text-Kennnummer").set_index("Text-Kennnummer")
fr_index = fr.drop_duplicates("Text-Kennnummer").set_index("Text-Kennnummer")
en_index = en.drop_duplicates("Text-Kennnummer").set_index("Text-Kennnummer")

# Zielspalten definieren
columns = [
    "descr1_de", "descr2_de", "text_de",
    "descr1_fr", "descr2_fr", "text_fr",
    "descr1_en", "descr2_en", "text_en"
]

# Spalten vorbereiten und als object casten
for col in columns:
    if col not in wip.columns:
        wip[col] = ""
    wip[col] = wip[col].astype("object")

# Mapping Hauptkennnummer → Liste untergeordneter Kennnummern
struktur = {
    "A_1": ["C_1_A_01_2018.1102", "C_1_A_02_2019.389"],
    "B_1": ["C_1_A_03_2019.404", "C_1_A_04_2019.255"],
    "A_2": ["C_2_A_01_J SURI 45", "C_2_A_02_2018.868", "C_2_A_03_2019.328", "C_2_A_04_2019.199", "C_2_A_05_2019.220", "C_2_A_06_2019.67"],
    "B_2": ["C_2_A_07_2019.290", "C_2_A_08_J SURI 63", "C_2_A_09_2019.186", "C_2_A_10_2019.157"],
    "C_4_A_01": ["C_4_A_01_2019.308", "C_4_A_02_Aa 132", "C_4_A_03_2019.13"],
    "B_4": ["C_4_A_08_2019.247a-e", "C_4_A_06_2019.132", "C_4_A_07_2018.1086", "C_4_A_04_2019.82", "C_4_A_05_2019.326"],
    "B_5": ["C_5_A_03_Muschel-Serie", "C_5_A_03_2019. 336", "C_5_A_03_2019.215", "C_5_A_08_ Juwel-Flüsse", "C_5_A_08_2019.277", "C_5_A_08_2019.86"],
    "ZT1": ["C_1_A_05_2019.288", "C_1_A_06_2019.70", "C_1_A_07_2018.1072", "C_1_A_08_2018.889", "C_1_A_09_2019.190", "C_1_A_10_2019.381"],
    "ZT2": ["C_2_A_07_2019.290", "C_2_A_08_J SURI 63", "C_2_A_09_2019.186", "C_2_A_10_2019.157"],
    "ZT3": ["C_2_A_11_2019.343", "C_2_A_12_2019.177", "C_2_A_13_2019.57", "C_2_A_14_2019.287"],
    "ZT4": ["C_3_A_09_2019.356", "C_3_A_08_2019.342", "C_3_A_07_2019.346", "C_3_A_06_2019.159", "C_3_A_05_RJP 3600"],
    "ZT5": ["C_3_A_04_2019.106", "C_3_A_03_2019.230", "C_3_A_02_2019.145", "C_3_A_01_2109.139"],
    "C_5_45 a-j_24 Generäle": ["C_5_A_07_1_2019.331", "C_5_A_07_2_2019.323", "C_5_A_07_4_J SURI 79-2", "C_5_A_07_3_J SURI 79-1", "C_5_AB_06_RJP 2804.1", "C_5_AB_06_RJP 2804.2", "C_5_AB_06_RJP 2804.3", "C_5_AB_06_RJP 2804.4", "C_5_AB_06_RJP 2804.5", "C_5_AB_06_RJP 2804.6"],
    "C_5_A_02_2018.905/.906/.907": ["C_5_A_02_2018.905", "C_5_A_02_2018.906", "C_5_A_02_2018.907"],
    "C_5_A_04_Pferde-Serie": ["C_5_A_04_2019.46", "C_5_A_04_J SURI 232-2", "C_5_A_04_J SURI 155", "C_5_A_04_J SURI 165", "C_5_A_04_RJP 3424", "C_5_A_04_J SURI 232-1", "C_5_A_04_2019.202"],
}

# HTML-Block-Erzeugung (escaped)
def make_html(row):
    titel = row.get("Titel", "")
    untertitel = row.get("Untertitel", "")
    laufnummer = row.get("Laufnummer", "")
    fliesstext = row.get("Fliesstext", "")
    urheber = row.get("Urheber*in", "")
    datierung = row.get("Datierung", "")
    technik = row.get("Material/Technik", "")
    credit = row.get("Creditline", "")
    return f'<h1>{laufnummer} | {titel}</h1><h2>{untertitel}</h2><small>{urheber}<br>{datierung}<br>{technik}<br>{credit}</small><br>{fliesstext}<br>'

# Zeilenweise befüllen und ggf. Inhalte anhängen
for i, row in wip.iterrows():
    key = row.get("Text-Kennnummer")
    if pd.notna(key):
        if key in de_index.index:
            wip.at[i, "descr1_de"] = de_index.at[key, "Titel"]
            wip.at[i, "descr2_de"] = de_index.at[key, "Untertitel"]
            wip.at[i, "text_de"] = de_index.at[key, "Fliesstext"]
        if key in fr_index.index:
            wip.at[i, "descr1_fr"] = fr_index.at[key, "Titel"]
            wip.at[i, "descr2_fr"] = fr_index.at[key, "Untertitel"]
            wip.at[i, "text_fr"] = fr_index.at[key, "Fliesstext"]
        if key in en_index.index:
            wip.at[i, "descr1_en"] = en_index.at[key, "Titel"]
            wip.at[i, "descr2_en"] = en_index.at[key, "Untertitel"]
            wip.at[i, "text_en"] = en_index.at[key, "Fliesstext"]
        # Kindinhalte anhängen
        if key in struktur:
            for kid in struktur[key]:
                for lang, content, col in [("de", de_index, "text_de"), ("fr", fr_index, "text_fr"), ("en", en_index, "text_en")]:
                    if kid in content.index:
                        html = make_html(content.loc[kid])
                        if not isinstance(wip.at[i, col], str):
                            wip.at[i, col] = ""
                        wip.at[i, col] += "<br><br>" + html

# Neue Datei speichern
output_path = "WIP/WIP--Japan_de_Luxe_Audio-Cult_INHALTE_UPDATED.xlsx"
wip.to_excel(output_path, index=False)
print(f"Datei gespeichert als: {output_path}")