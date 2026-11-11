import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule
import sys

OUT = sys.argv[1]
D = "2026-11-11"
wb = openpyxl.Workbook()
H = Font(bold=True, color="FFFFFF")
HF = PatternFill("solid", fgColor="1F4E78")
INP = PatternFill("solid", fgColor="FFF2CC")  # da inserire
PCT = "0.0%"; PCT2 = "0.00%"; EUR = "#,##0 €"


def sheet(name, headers, widths=None):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for c in ws[1]:
        c.font = H; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    for i, w in enumerate(widths or [], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    return ws


wb.remove(wb.active)

# ---------- Dashboard ----------
ws = sheet("Dashboard", ["KPI", "Valore", "Sett. prec. (04/11)", "Var.", "Tipo dato", "Data dato", "Fonte"],
           [46, 16, 18, 12, 26, 14, 70])
rows = [
    ("Prezzo richiesto Venezia comune €/mq (idealista)", "=Prezzi!G2", "", "", "Prezzo richiesto", "ott 2026", "https://www.edilizia.com/?p=102973"),
    ("Prezzo richiesto Venezia comune €/mq (Immobiliare.it)", "=Prezzi!G3", "", "", "Prezzo richiesto", "set 2026", "https://www.immobiliare.it/mercato-immobiliare/veneto/venezia/"),
    ("Prezzo richiesto Mestre-Chirignago-Marghera-Catene €/mq", "=Prezzi!G4", "", "", "Prezzo richiesto", "set 2026", "https://www.immobiliare.it/mercato-immobiliare/veneto/venezia/"),
    ("Stima automatica Venezia €/mq (RealAdvisor)", "=Prezzi!G5", "", "", "Valutazione automatica", "Q3 2026", "https://realadvisor.it/it/blog/barometro-immobiliare-italiano-t3-2026"),
    ("Var. mensile prezzo richiesto Venezia", "=Prezzi!M2", "", "", "Prezzo richiesto", "ott 2026", "idealista via Edilizia.com"),
    ("IPAB Italia a/a", "=Prezzi!N9", 0.04, "=IFERROR(B7-C7,\"\")", "Indice ufficiale", "Q2 2026", "https://www.istat.it/wp-content/uploads/2026/09/CS-abitazioni-Q22026.pdf"),
    ("Stock annunci Italia, var. trimestrale", "=Offerta!K5", "", "", "Inventario portale", "Q3 2026", "https://www.idealista.it/news/node/490575"),
    ("Annunci attivi Mestre (settimanale)", '=IF(Offerta!B2="","n.d.",Offerta!B2)', "", "", "–", "n.d.", "Portali non accessibili da questa sessione"),
    ("% annunci con ribasso Mestre", "=Offerta!F2", "", "", "–", "n.d.", "Portali non accessibili"),
    ("TAN fisso indicativo (miglior offerta standard)", "=Mutui!B2", "=Mutui!C2", "=Mutui!D2", "Offerta bancaria", "nov 2026", "Facile.it / Altroconsumo"),
    ("TAN variabile indicativo", "=Mutui!B3", "=Mutui!C3", "=Mutui!D3", "Offerta bancaria", "nov 2026", "Facile.it"),
    ("Euribor 3M", "=Mutui!B5", "=Mutui!C5", "=Mutui!D5", "Tasso di mercato", "06/11/2026", "Facile.it"),
    ("IRS 25 anni", "=Mutui!B8", "=Mutui!C8", "=Mutui!D8", "Tasso di mercato", "06/11/2026", "Facile.it"),
    ("Rata mutuo tipo fisso (200k, 25 a)", "=Mutui!B15", "=Mutui!C15", "=IFERROR(B15-C15,\"\")", "Calcolo", D, "Foglio Mutui"),
    ("Rata mutuo tipo variabile (200k, 25 a)", "=Mutui!B16", "=Mutui!C16", "=IFERROR(B16-C16,\"\")", "Calcolo", D, "Foglio Mutui"),
    ("Canone richiesto Venezia €/mq/mese", "=Affitti!D2", "", "", "Canone richiesto", "ott 2026", "idealista (fonte secondaria)"),
    ("Rendimento lordo Mestre", "=Affitti!F3", "", "", "Calcolo su richiesti", "set 2026", "Foglio Affitti"),
    ("Leverage compratori/venditori (-1 compratori … +1 venditori)", "=Leverage!B10", "", "", "Indicatore sintetico", D, "Foglio Leverage"),
    ("Semaforo", '=IF(B19>0.33,"Favorevole venditori",IF(B19<-0.33,"Favorevole compratori","Equilibrato"))', "Equilibrato", "", "", D, ""),
    ("Decisione area Mestre/Venezia", "=Decisioni!B2", "ATTENDERE", "", "", D, "Foglio Decisioni"),
]
for r in rows:
    ws.append(list(r))
for r in (6, 7, 8, 10, 11, 12, 13, 14, 18):
    for col in "BCD":
        ws[f"{col}{r}"].number_format = PCT2
for r in (2, 3, 4, 5, 15, 16):
    for col in "BCD":
        ws[f"{col}{r}"].number_format = EUR
ws["B17"].number_format = "0.0"
ws["B19"].number_format = "0.00"
ws.insert_rows(1)
ws["A1"] = f"Dashboard mercato immobiliare – Mestre / Venezia – report {D}"
ws["A1"].font = Font(bold=True, size=14)
# after insert, formulas referencing own sheet shift: fix them
for r in range(2, ws.max_row + 1):
    for col in "BCD":
        v = ws[f"{col}{r}"].value
        if isinstance(v, str) and v.startswith("=") and "!" not in v:
            import re
            v = re.sub(r"\b([BCD])(\d+)\b", lambda m: f"{m.group(1)}{int(m.group(2))+1}", v)
            ws[f"{col}{r}"].value = v

# ---------- Aree_CAP ----------
ws = sheet("Aree_CAP", ["Area", "CAP", "Tipo", "Zona OMI", "Priorità", "Nota"], [34, 16, 14, 14, 14, 60])
for r in [
    ("Mestre centro", "30171-30174", "Quartiere", "da verificare", "Domia Group", "Zona OMI da confermare sulla banca dati OMI"),
    ("Mestre – Chirignago, Marghera, Catene", "30174/30175", "Quartiere", "", "Domia Group", "Zona aggregata Immobiliare.it"),
    ("Venezia centro storico", "30121-30135", "Quartiere", "", "Domia Group", ""),
    ("Venezia (comune)", "30100-30176", "Comune", "", "Domia Group", ""),
    ("Provincia di Venezia", "", "Provincia", "", "Domia Group", "Nessun dato locale raccolto anche questa settimana"),
    ("[CAP/aree da inserire]", "", "", "", "", "Segnaposto del prompt ancora non compilati"),
]:
    ws.append(list(r))
for c in ws[7]: c.fill = INP

# ---------- Prezzi ----------
ws = sheet("Prezzi", ["Area", "Fonte", "Tipo dato", "Data dato", "Min €/mq", "Max €/mq", "Medio €/mq", "Mediano €/mq",
                      "Sett. prec. €/mq", "Mese prec. €/mq", "Anno prec. €/mq", "Var. sett. %", "Var. mese %", "Var. anno %", "URL"],
           [32, 18, 22, 10, 9, 9, 11, 11, 11, 11, 11, 10, 10, 10, 60])
data = [
    ("Venezia comune", "idealista", "Prezzo richiesto", "2026-10", None, None, 4970, None, None, "=G2/(1+0.003)", None, None, None, "https://www.edilizia.com/?p=102973"),
    ("Venezia comune", "Immobiliare.it", "Prezzo richiesto", "2026-09", None, None, 3423, None, None, None, None, None, None, "https://www.immobiliare.it/mercato-immobiliare/veneto/venezia/"),
    ("Mestre-Chirignago-Marghera-Catene", "Immobiliare.it", "Prezzo richiesto", "2026-09", None, None, 2109, None, None, None, "=G4/(1+0.102)", None, None, "idem – var. +10,2% riportata, periodo di confronto non esplicito (ipotesi: a/a)"),
    ("Venezia comune", "RealAdvisor", "Valutazione automatica", "2026-Q3", None, None, 2906, None, None, None, None, None, None, "https://realadvisor.it/wp-content/uploads/2026/10/barometer-it-it-q3-2026.pdf (modello ricalibrato ago 2026: non confrontabile con trimestri prec.)"),
    ("Mestre centro (OMI)", "Caasa (annunci)", "Prezzo richiesto", "2026-07", 2000, 3000, "=AVERAGE(E6:F6)", "=MEDIAN(E6:F6)", 2500, None, None, None, None, "https://www.caasa.it/veneto/venezia/venezia/vendita-65a7a250-opinione.html (dato della sett. prec., non aggiornato)"),
    ("Mestre 30172", "RealAdvisor", "Valutazione automatica", "2026-07", None, None, 2260, None, 2260, None, None, None, None, "https://realadvisor.it/it/mercato-immobiliare/30172-venezia/via-mestrina (non aggiornato)"),
    ("Italia (idealista, usato)", "idealista", "Prezzo richiesto", "2026-10", None, None, 2012, None, None, "=G8/(1-0.009)", "=G8/(1+0.015)", None, None, "https://www.idealista.it/news/node/497734"),
    ("Italia (IPAB tutte)", "ISTAT", "Indice ufficiale (transazioni)", "2026-Q2", None, None, None, None, None, None, None, None, None, "https://www.istat.it/wp-content/uploads/2026/09/CS-abitazioni-Q22026.pdf"),
    ("Italia (IPAB nuove)", "ISTAT", "Indice ufficiale", "2026-Q2", None, None, None, None, None, None, None, None, None, "idem"),
    ("Italia (IPAB esistenti)", "ISTAT", "Indice ufficiale", "2026-Q2", None, None, None, None, None, None, None, None, None, "idem"),
    ("Il mio immobile", "[da inserire]", "Stima", "", None, None, None, None, None, None, None, None, None, "Inserire indirizzo, mq, piano, stato, classe energetica"),
]
for i, r in enumerate(data, 2):
    r = list(r)
    r[11] = f'=IFERROR(G{i}/I{i}-1,"")'
    r[12] = f'=IFERROR(G{i}/J{i}-1,"")'
    r[13] = f'=IFERROR(G{i}/K{i}-1,"")'
    ws.append(r)
ws["N9"] = 0.040; ws["N10"] = 0.050; ws["N11"] = 0.037
for row in ws.iter_rows(min_row=2, min_col=12, max_col=14):
    for c in row: c.number_format = PCT
for row in ws.iter_rows(min_row=2, min_col=5, max_col=11):
    for c in row: c.number_format = "#,##0"
for c in ws[12]: c.fill = INP
ws.append([])
ws.append(["Stima mio immobile", "Superficie mq", "", "", "", "", "Valore centrale €", "Min (-8%)", "Max (+5%)"])
ws.append(["", None, "€/mq di riferimento →", "=G4", "", "", '=IFERROR(B15*D15,"")', '=IFERROR(G15*0.92,"")', '=IFERROR(G15*1.05,"")'])
ws["B15"].fill = INP
ws.append(["Nota: range su prezzo richiesto di zona; lo sconto di chiusura va calibrato su OMI/rogiti (Bankitalia: sconto medio nazionale da aggiornare)."])

# ---------- Offerta ----------
ws = sheet("Offerta", ["Area", "Annunci attivi", "Nuovi annunci sett.", "Rimossi sett.", "Annunci con ribasso", "% ribassi",
                       "Ribasso medio %", "Giorni online medi", "Saldo offerta", "Var. stock %", "Var. stock trim. %", "Fonte/Nota"],
           [26, 12, 12, 12, 12, 10, 10, 10, 10, 10, 12, 80])
offer = [
    ("Mestre", None, None, None, None, None, None, None, None, None, None, "Dato settimanale non ottenibile: portali non raggiungibili dalla sessione"),
    ("Venezia comune", None, None, None, None, None, None, None, None, None, 0.046, "idealista Q2 2026: invenduto +4,6% nel trimestre https://m.dellevittorie.it/mercato-residenziale-compravendite-q2-2026-omi/"),
    ("Provincia di Venezia", None, None, None, None, None, None, None, None, None, None, "n.d."),
    ("Italia (idealista)", None, None, None, None, None, None, None, None, None, -0.056, "Stock in vendita -5,6% nel Q3 2026 https://www.idealista.it/news/node/490575"),
    ("Italia (Bankitalia Q2 2026)", None, None, None, None, None, None, None, None, None, None, "Domanda contenuta, offerta in riduzione; saldo prezzi 7 (da 12). Q3 atteso fine nov. https://www.bancaditalia.it/pubblicazioni/sondaggio-abitazioni/2026-sondaggio-abitazioni/02/"),
    ("Italia (OMI compravendite Q2 2026)", None, None, None, None, None, None, None, None, None, None, "201.596 compravendite (+0,1% a/a); nuove +30,8%, esistenti -0,8%"),
]
for i, r in enumerate(offer, 2):
    r = list(r)
    r[5] = f'=IFERROR(E{i}/B{i},"")'
    r[8] = f'=IFERROR(C{i}-D{i},"")'
    ws.append(r)
ws["K6"].value = None
ws["K5"] = -0.056; ws["K3"] = 0.046
for row in ws.iter_rows(min_row=2, min_col=6, max_col=11):
    for c in row:
        if c.column in (6, 7, 10, 11): c.number_format = PCT
for col in "BCDEGH":
    ws[f"{col}2"].fill = INP
ws.append([])
ws.append(["Segmenti (Mestre)", "Annunci", "Prezzo medio €/mq", "% ribassi", "Giorni online"])
for s in ["Bilocali", "Trilocali", "Pregio", "Nuove costruzioni", "Da ristrutturare"]:
    ws.append([s]); [setattr(c, "fill", INP) for c in ws[ws.max_row][1:5]]

# ---------- Mutui ----------
ws = sheet("Mutui", ["Parametro", "Valore", "Sett. prec.", "Var. (pp)", "Fonte / data"], [34, 12, 12, 12, 90])
m = [
    ("TAN fisso indicativo", 0.0319, 0.0300, "Miglior fisso standard (non green) nov 2026: Facile.it https://www.facile.it/mutui/news/mutui-online-offerte-novembre-2026.html – sett. prec. = media offerte fine ott (serie diversa)"),
    ("TAN variabile indicativo", 0.0283, 0.0265, "Facile.it nov 2026 (\"da\" 2,83%) – sett. prec. media offerte fine ott (serie diversa)"),
    ("Euribor 1M", 0.02473, None, "Facile.it 06/11/2026 https://www.facile.it/mutui/news/aggiornamento-tassi-eurirs-euribor-6-novembre-2026.html"),
    ("Euribor 3M", 0.02794, 0.02765, "06/11 vs 03/11/2026 (Facile.it). NB il report prec. riportava 2,477% del 6 ago: non era aggiornato"),
    ("Euribor 6M", 0.03016, None, "Facile.it 06/11/2026"),
    ("Euribor 12M", 0.03248, None, "Facile.it 06/11/2026"),
    ("IRS 25 anni", 0.0331, 0.0336, "06/11 vs 04/11/2026 (Facile.it)"),
    ("IRS 30 anni", 0.0324, 0.0329, "06/11 vs 03/11/2026 (Facile.it)"),
    ("BCE tasso depositi", 0.025, 0.025, "Invariato il 29/10/2026; mercato prezza ~86% rialzo a dicembre (Altroconsumo) https://www.altroconsumo.it/investi/investire/mercati-e-valute/news/2026/10/bce-riunione-29-ottobre"),
    ("TAN medio fisso effettivo (MutuiOnline)", 0.0353, None, "Settembre 2026, massimo da nov 2023 (citato da Facile.it/stampa)"),
]
for i, r in enumerate(m, 2):
    ws.append([r[0], r[1], r[2], f'=IFERROR(B{i}-C{i},"")', r[3]])
for row in ws.iter_rows(min_row=2, max_row=11, min_col=2, max_col=4):
    for c in row: c.number_format = "0.000%"
ws.append(["Prezzo acquisto", 250000]); ws.append(["Anticipo %", 0.2]); ws.append(["Durata anni", 25])
# rows 12,13,14
ws.append(["Rata fisso", "=-PMT(B2/12,B14*12,B17)", "=-PMT(C2/12,B14*12,B17)", "=B15-C15"])
ws.append(["Rata variabile", "=-PMT(B3/12,B14*12,B17)", "=-PMT(C3/12,B14*12,B17)", "=B16-C16"])
ws.append(["Capitale", "=B12*(1-B13)"])
ws.append(["Differenza fisso-variabile", "=B15-B16"])
ws.append(["Rata fisso +0,5pp (stress)", "=-PMT((B2+0.005)/12,B14*12,B17)"])
ws.append(["Rata al TAN medio effettivo (3,53%)", "=-PMT(B11/12,B14*12,B17)"])
ws.append(["Rata/reddito netto mensile", '=IFERROR(B15/B22,"")', "", "", "Inserire reddito netto in B22 (soglia prudenziale ~30-35%)"])
ws.append(["Reddito netto mensile famiglia", None]); ws["B22"].fill = INP
ws["B13"].number_format = PCT
for r in range(15, 21):
    for col in "BCD": ws[f"{col}{r}"].number_format = "#,##0.00 €"
ws["B12"].number_format = EUR; ws["B17"].number_format = EUR; ws["B21"].number_format = PCT
ws.append([])
ws.append(["Mutuo in essere (per valutare surroga)", "Valore"])
for p in ["Debito residuo €", "Anni residui", "Tasso attuale (TAN)", "Tipo (fisso/variabile)"]:
    ws.append([p]); ws[f"B{ws.max_row}"].fill = INP
ws.append(["Rata attuale", '=IFERROR(-PMT(B27/12,B26*12,B25),"")'])
ws.append(["Rata con nuovo fisso indicativo", '=IFERROR(-PMT(B2/12,B26*12,B25),"")'])
ws.append(["Risparmio mensile stimato", '=IFERROR(B29-B30,"")', "", "", "Solo indicativo: verificare con banca/mediatore creditizio; surroga senza costi per legge, ma valutare condizioni e polizze"])

# ---------- Affitti ----------
ws = sheet("Affitti", ["Area", "Canone mensile €", "Superficie mq", "Canone €/mq/mese", "Prezzo €/mq", "Rend. lordo %", "Var. canone", "Fonte"],
           [30, 14, 12, 14, 12, 12, 12, 90])
ws.append(["Venezia comune", None, None, 20.1, "=Prezzi!G2", '=IFERROR(D2*12/E2,"")', None, "idealista ott 2026 (via articolo stampa, da verificare). Q2 2026: 21,7 €/mq https://www.idealista.it/news/node/496729"])
ws.append(["Mestre-Chirignago-Marghera-Catene", None, None, 12.42, "=Prezzi!G4", '=IFERROR(D3*12/E3,"")', None, "Immobiliare.it set 2026 https://www.immobiliare.it/mercato-immobiliare/veneto/venezia/"])
ws.append(["Mestre 30172 (appartamento medio)", 1368, None, '=IFERROR(B4/C4,"")', "=Prezzi!G7", '=IFERROR(D4*12/E4,"")', None, "RealAdvisor lug 2026 – superficie non indicata"])
ws.append(["Italia (idealista)", None, None, 14.6, "=Prezzi!G8", '=IFERROR(D5*12/E5,"")', -0.003, "idealista ott 2026: -0,3% m/m, +2,3% a/a"])
ws.append(["Il mio immobile", None, None, '=IFERROR(B6/C6,"")', "=Prezzi!G12", '=IFERROR(B6*12/(Prezzi!B15*Prezzi!D15),"")', None, "Inserire canone atteso e superficie"])
for c in ws[6][1:3]: c.fill = INP
for r in range(2, 7):
    ws[f"F{r}"].number_format = PCT; ws[f"G{r}"].number_format = PCT; ws[f"D{r}"].number_format = "0.00"
ws.append([])
ws.append(["Confronto formule (indicativo)", "Note"])
ws.append(["Tradizionale 4+4 / concordato 3+2", "Concordato: cedolare 10% nei comuni ad alta tensione abitativa (Venezia inclusa) – verificare accordo territoriale"])
ws.append(["Transitorio (1-18 mesi)", "Domanda in crescita (idealista: offerta transitori +41,6% a/a nazionale)"])
ws.append(["Breve (<30 gg)", "Venezia centro storico: limiti comunali / ticket d'accesso e CIN obbligatorio; verificare regolamento comunale vigente prima di ogni scelta"])

# ---------- Storico ----------
ws = sheet("Storico", ["Settimana", "Data", "Prezzo rich. Venezia €/mq (idealista)", "Prezzo rich. Mestre €/mq (Imm.it)", "TAN fisso", "Euribor 3M",
                       "IRS 25", "Annunci attivi Mestre", "Var. 4 sett. prezzo Venezia", "Var. 12 sett.", "Var. 52 sett.",
                       "Trend 4 sett. (pendenza €/sett)", "Var. 4 sett. TAN (pp)"], [10, 12, 16, 16, 10, 10, 10, 12, 12, 12, 12, 16, 14])
ws.append([1, "2026-11-04", None, None, 0.03, 0.02477, None, None])
ws.append([2, D, "=Prezzi!G2", "=Prezzi!G4", "=Mutui!B2", "=Mutui!B5", "=Mutui!B8", "=IF(Offerta!B2=\"\",\"\",Offerta!B2)"])
for n in range(3, 55):
    ws.append([n])
for r in range(2, 56):
    if r >= 6: ws[f"I{r}"] = f'=IFERROR(C{r}/C{r-4}-1,"")'; ws[f"M{r}"] = f'=IFERROR(E{r}-E{r-4},"")'
    if r >= 14: ws[f"J{r}"] = f'=IFERROR(C{r}/C{r-12}-1,"")'
    if r >= 54: ws[f"K{r}"] = f'=IFERROR(C{r}/C{r-52}-1,"")'
    if r >= 5: ws[f"L{r}"] = f'=IFERROR(SLOPE(C{r-3}:C{r},A{r-3}:A{r}),"")'
    for col in "EFG": ws[f"{col}{r}"].number_format = "0.000%"
    for col in "IJKM": ws[f"{col}{r}"].number_format = PCT2
ws["A57"] = "Nota sett. 1: Euribor 3M era il dato del 6 ago (non aggiornato); prezzi sett. 1 su fonti diverse (non confrontabili)."

# ---------- Leverage ----------
ws = sheet("Leverage", ["Componente", "Valore", "Punteggio (-1 compratori … +1 venditori)", "Peso", "Regola"], [36, 12, 22, 8, 70])
L = [
    ("Var. mensile prezzo richiesto Venezia", "=Prezzi!M2", '=IF(B2="",0,IF(B2>0.005,1,IF(B2<-0.005,-1,0)))', 0.25, ">+0,5% m/m = +1; <-0,5% = -1"),
    ("% annunci con ribasso (Mestre)", "=Offerta!F2", '=IF(B3="",0,IF(B3<0.15,1,IF(B3>0.3,-1,0)))', 0.25, "<15% = +1; >30% = -1"),
    ("Var. stock annunci Venezia (trim.)", "=Offerta!K3", '=IF(B4="",0,IF(B4<-0.03,1,IF(B4>0.03,-1,0)))', 0.2, "stock in calo >3% = +1; in aumento >3% = -1"),
    ("Saldo Bankitalia prezzi (Italia)", 7, '=IF(B5="",0,IF(B5>10,1,IF(B5<-10,-1,0)))', 0.1, "saldo > +10 = +1; < -10 = -1"),
    ("Var. sett. IRS 25 anni (pp)", "=Mutui!D8", '=IF(B6="",0,IF(B6>0.001,-1,IF(B6<-0.001,1,0)))', 0.2, "IRS su >10 pb = -1 (credito più caro, meno domanda); giù >10 pb = +1. Il 97% delle richieste è a tasso fisso"),
]
for r in L: ws.append(list(r))
ws["B2"].number_format = PCT; ws["B3"].number_format = PCT; ws["B4"].number_format = PCT; ws["B6"].number_format = PCT2
ws.append([]); ws.append([]); ws.append([])
ws.append(["Indicatore sintetico (media ponderata)", "=SUMPRODUCT(C2:C6,D2:D6)/SUM(D2:D6)", "", "", ">+0,33 favorevole venditori; <-0,33 favorevole compratori; altrimenti equilibrato"])
ws["B10"].number_format = "0.00"
ws.append(["Componenti con dato disponibile", '=COUNT(B2:B6)&" su 5"'])

# ---------- Decisioni ----------
ws = sheet("Decisioni", ["Area", "Giudizio", "Confidenza", "Evidenza 1", "Evidenza 2", "Evidenza 3", "Rischio principale", "Azione entro 7 gg"],
           [24, 16, 11, 40, 40, 40, 40, 50])
dec = [
    ("Venezia comune (Domia)", "MANTENERE", "Media-bassa",
     "Prezzo richiesto 4.970 €/mq, +0,3% m/m a ott. mentre Italia -0,9% (idealista)",
     "Canone richiesto ~20,1 €/mq/mese, 3° capoluogo più caro (idealista ott.)",
     "Euribor 3M 2,794% (+2,9 pb in 3 gg), BCE: rialzo a dicembre prezzato all'86%",
     "Stretta sul credito che raffredda la domanda nel 2027; invenduto Venezia +4,6% in Q2",
     "Scaricare report OMI 1° sem. 2026 per le zone del comune e confrontarlo con i 4.970 €/mq richiesti"),
    ("Mestre (Domia)", "MANTENERE", "Bassa",
     "Mestre-Chirignago-Marghera-Catene 2.109 €/mq richiesto (+10,2%, Immobiliare.it set.)",
     "Rendimento lordo potenziale ~7,1% (12,42 €/mq × 12 / 2.109)",
     "Compravendite Italia Q2 piatte (+0,1% a/a, OMI); usato -0,8%",
     "Dati di inventario/ribassi locali assenti; prezzo +10,2% su base non esplicita",
     "Rilevare a mano annunci attivi, ribassi e giorni online su 30172-30174 (Immobiliare.it/idealista)"),
    ("Provincia di Venezia", "ATTENDERE", "Bassa", "Nessun dato locale", "IPAB Italia +4,0% a/a (Q2)", "Stock nazionale -5,6% Q3 (idealista)",
     "Decidere senza dati locali", "Indicare i comuni della provincia da monitorare"),
    ("Mutuo (generico, se variabile)", "RINEGOZIARE O SURROGARE (da valutare)", "Bassa",
     "Euribor 3M 2,794% vs 2,477% di agosto (+0,32 pp)",
     "IRS 25A 3,31%, in calo di 5 pb in settimana: finestra per fissare",
     "Rata variabile tipo 931 € vs fisso 968 €: differenza solo 37 €/mese",
     "Rialzo BCE già scontato: il vantaggio del fisso può svanire o invertirsi",
     "Solo se si ha un variabile: chiedere 2-3 preventivi di surroga a tasso fisso a banca/mediatore creditizio abilitato"),
    ("Il mio immobile", "ATTENDERE", "Bassa", "Dati immobile non forniti", "Riferimento zona Mestre 2.109 €/mq richiesto", "Range stima calcolabile in Prezzi!G15:I15 inserendo i mq",
     "Nessuna stima possibile", "Inserire indirizzo, mq, stato, classe energetica e mutuo in essere"),
]
for r in dec: ws.append(list(r))
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")

wb.save(OUT)
print("ok", OUT)
