import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
out=sys.argv[1]
wb=Workbook()
H=Font(bold=True,color="FFFFFF"); HF=PatternFill("solid",fgColor="1F3A5F")
INP=PatternFill("solid",fgColor="FFF4CC"); T=Font(bold=True,size=14)
def hdr(ws,row,vals):
    for i,v in enumerate(vals,1):
        c=ws.cell(row=row,column=i,value=v); c.font=H; c.fill=HF; c.alignment=Alignment(wrap_text=True,vertical="center")
def widths(ws,w):
    for i,x in enumerate(w,1): ws.column_dimensions[get_column_letter(i)].width=x
D="2026-11-25"
S={ # sources
 "IMM":"https://www.immobiliare.it/mercato-immobiliare/veneto/venezia/",
 "IDE":"https://www.idealista.it/valutazione-di-immobili/mestre-venezia",
 "MIT":"https://immobiliare.mitula.it/casa-mestre",
 "CAA":"https://www.caasa.it/veneto/venezia/venezia/vendita-65f71407-opinione.html",
 "MII":"https://www.mercato-immobiliare.info/veneto/venezia/venezia.html",
 "RAP":"https://realadvisor.it/it/mercato-immobiliare/provincia-venezia",
 "IMP":"https://www.immobiliare.it/pt/mercato-immobiliare/veneto/venezia-provincia/",
 "BDI":"https://www.bancaditalia.it/pubblicazioni/sondaggio-abitazioni/2026-sondaggio-abitazioni/03/index.html",
 "REQ":"https://requadro.com/residenziale-sottotono-e-attese-degli-agenti-indebolite-secondo-il-sondaggio-congiunturale-q3-26/",
 "IST":"https://www.istat.it/comunicato-stampa/prezzi-delle-abitazioni-dati-provvisori-ii-trimestre-2026/",
 "OMI":"https://www.idealista.it/news/node/448146",
 "OMIQ":"https://www.facile.it/mutui/news/quotazioni-omi-primo-semestre-2026.html",
 "FAC18":"https://www.facile.it/mutui/news/aggiornamento-tassi-eurirs-euribor-18-novembre-2026.html",
 "FAC17":"https://www.facile.it/mutui/news/aggiornamento-tassi-eurirs-euribor-17-novembre-2026.html",
 "FAC10":"https://www.facile.it/mutui/news/aggiornamento-tassi-eurirs-euribor-10-novembre-2026.html",
 "TAEG":"https://www.simplybiz.eu/bancaditalia-mutui-tassi-erogazioni-settembre-2026/",
 "MUT":"https://www.mutui.it/news/mutui-online-novembre-2026.html",
 "BCE":"https://www.bancaditalia.it/media/bce-comunicati/documenti/2026/ecb.mp261029.it.pdf",
 "STR":"https://www.holidu.it/magazine/regole-affitti-brevi-venezia-per-host",
 "IDA":"https://www.idealista.it/affitto-case/mestre-venezia/con-affitto-lungo-termine/",
}

# ---------- Mio immobile (inputs) ----------
mi=wb.active; mi.title="Mio_Immobile"
mi["A1"]="Il mio immobile – INPUT (celle gialle da compilare)"; mi["A1"].font=T
rows=[("Indirizzo / zona","DA INSERIRE – stima illustrativa su Mestre","testo"),
("CAP","DA INSERIRE",""),
("Superficie commerciale (mq)",80,"ESEMPIO: sostituire con il dato reale"),
("Prezzo richiesto di riferimento zona (€/mq)","=Prezzi_Quotazioni!C12","mediana tra portali (prezzo RICHIESTO)"),
("Sconto medio richiesta→chiusura","=Offerta_Ribassi!C5","Banca d'Italia Q3-2026 (nazionale)"),
("Coefficiente stato/piano/classe energetica",1.00,"1,00 = medio; 0,85 da ristrutturare; 1,10 ristrutturato/classe A-B"),
("Valore stimato centrale (€)","=B5*B6*(1-B7)*B8","STIMA (non transato)"),
("Ampiezza intervallo (±)",0.07,"incertezza ±7% (campione ridotto)"),
("Valore minimo realistico (€)","=B9*(1-B10)",""),
("Valore massimo realistico (€)","=B9*(1+B10)",""),
("Valore stimato settimana precedente (€)","","compilare dalla prossima settimana"),
("Variazione stimata vs settimana prec.","=IF(B13=\"\",\"n.d.\",B9/B13-1)",""),
("Debito residuo mutuo (€)","","DA INSERIRE se presente"),
("Tasso attuale mutuo","","DA INSERIRE"),
("Tipo tasso (fisso/variabile)","","DA INSERIRE"),
("Canone stimato mensile (€)","=B5*Affitti_Rendimenti!C5","canone RICHIESTO di zona × mq"),
("Rendimento lordo su valore stimato","=B18*12/B9",""),
]
hdr(mi,2,["Voce","Valore","Note / natura del dato"])
for i,(a,b,c) in enumerate(rows,3):
    mi.cell(row=i,column=1,value=a); mi.cell(row=i,column=2,value=b); mi.cell(row=i,column=3,value=c)
for r in (3,4,5,8,10,13,15,16,17): mi.cell(row=r,column=2).fill=INP
for r in (9,11,12,13,15,18): mi.cell(row=r,column=2).number_format='#,##0 €'
for r in (7,10,14,16,19): mi.cell(row=r,column=2).number_format='0.0%'
mi["B6"].number_format='#,##0 €'
widths(mi,[44,22,60])

# ---------- Prezzi e quotazioni ----------
p=wb.create_sheet("Prezzi_Quotazioni")
p["A1"]="Prezzi e quotazioni – Mestre / Venezia / provincia (aggiornato "+D+")"; p["A1"].font=T
hdr(p,3,["Area","Fonte","€/mq vendita","Natura del dato","Periodo dato","Var. a/a","Frequenza aggiornamento","Link"])
data=[
("Mestre-Chirignago-Marghera-Catene","Immobiliare.it",2109,"Prezzo richiesto (media)","set-2026",None,"mensile",S["IMM"]),
("Mestre","idealista",1983,"Prezzo richiesto (media)","ago-2026",None,"mensile",S["IDE"]),
("Mestre","Mitula (aggreg. annunci)",2190,"Prezzo richiesto (media)","n.d.",None,"n.d.",S["MIT"]),
("Mestre – Bissuola","Caasa",2280,"Prezzo richiesto (stima)","n.d.",None,"n.d.",S["CAA"]),
("Mestre","mercato-immobiliare.info",2250,"Prezzo richiesto (media)","n.d.",None,"n.d.",S["MII"]),
]
for i,r in enumerate(data,4):
    for j,v in enumerate(r,1): p.cell(row=i,column=j,value=v)
p["A10"]="Media Mestre (tra fonti)"; p["C10"]="=AVERAGE(C4:C8)"
p["A11"]="Dev. std tra fonti"; p["C11"]="=STDEV(C4:C8)"
p["A12"]="Mediana Mestre (tra fonti)"; p["C12"]="=MEDIAN(C4:C8)"
p["D12"]="NB: mediana delle medie di portale, NON mediana degli annunci"
for r in (10,11,12): p.cell(row=r,column=1).font=Font(bold=True)
hdr(p,14,["Area","Fonte","€/mq vendita","Natura del dato","Periodo dato","Var. a/a","Frequenza aggiornamento","Link"])
d2=[
("Comune di Venezia (totale)","Immobiliare.it",3423,"Prezzo richiesto (media)","set-2026",None,"mensile",S["IMM"]),
("Venezia – San Marco/Rialto","Immobiliare.it",6032,"Prezzo richiesto (media)","set-2026",None,"mensile",S["IMM"]),
("Provincia di Venezia","Immobiliare.it",2815,"Prezzo richiesto (media)","giu-2026",0.0399,"mensile",S["IMP"]),
("Provincia di Venezia","RealAdvisor",2058,"Valutazione automatica (AVM)","ago-2026",0.042,"mensile",S["RAP"]),
("Mestre – zone OMI","Agenzia Entrate OMI",None,"Quotazione ufficiale (min-max)","1° sem 2026","","semestrale – DA INSERIRE da Geopoi",S["OMIQ"]),
("Italia – IPAB totale","ISTAT",None,"Indice prezzi transati","Q2-2026",0.040,"trimestrale (prox. 17-dic-2026)",S["IST"]),
("Italia – IPAB nuove","ISTAT",None,"Indice prezzi transati","Q2-2026",0.050,"trimestrale",S["IST"]),
("Italia – IPAB esistenti","ISTAT",None,"Indice prezzi transati","Q2-2026",0.037,"trimestrale",S["IST"]),
]
for i,r in enumerate(d2,15):
    for j,v in enumerate(r,1): p.cell(row=i,column=j,value=v)
p["A24"]="Gap richiesto vs OMI (Mestre)"; p["C24"]='=IF(C19="","n.d. – inserire OMI medio in C19",C12/C19-1)'
p["A25"]="Prezzo implicito di chiusura Mestre (richiesto × (1-sconto))"; p["C25"]="=C12*(1-Offerta_Ribassi!C5)"
p["D25"]="STIMA di prezzo transato (sconto nazionale BdI)"
for r in list(range(4,9))+list(range(15,23))+[10,12,25]: p.cell(row=r,column=3).number_format='#,##0'
for r in range(4,23): p.cell(row=r,column=6).number_format='0.0%'
p["C19"].fill=INP
widths(p,[40,26,14,32,13,10,30,70])

# ---------- Offerta e ribassi ----------
o=wb.create_sheet("Offerta_Ribassi")
o["A1"]="Offerta, domanda, ribassi e trattabilità"; o["A1"].font=T
hdr(o,3,["Indicatore","Area","Valore","Periodo","Valore periodo prec.","Variazione","Natura / note","Link"])
od=[
("Saldo agenti: prezzi in aumento − diminuzione (p.p.)","Italia",0,"Q3-2026",7,None,"Sondaggio BdI – ~1.500 agenzie",S["BDI"]),
("Sconto medio prezzo richiesto → chiusura","Italia",0.077,"Q3-2026",0.071,None,"BdI (7,7% da tavole riportate; arrotondato 8%)",S["REQ"]),
("Tempo medio di vendita (mesi)","Italia",5.5,"Q3-2026",5.5,None,"BdI – stabile",S["REQ"]),
("Quota acquisti con mutuo","Italia",0.69,"Q3-2026",0.63,None,"BdI – massimo dal Q2-2022",S["BDI"]),
("Saldo attese prezzi trimestre corrente (p.p.)","Italia",-17,"Q3-2026",-15,None,"BdI",S["BDI"]),
("Saldo attese mercato biennio (p.p.)","Italia",-21,"Q3-2026",-10,None,"BdI",S["BDI"]),
("Compravendite residenziali – var. a/a","Italia",0.001,"Q2-2026",None,None,"OMI (Q3-2026 atteso inizio dic.)",S["OMI"]),
("Quota acquisti PF con mutuo (OMI)","Italia",0.466,"Q2-2026",0.479,None,"OMI",S["OMI"]),
("Annunci attivi vendita","Mestre",None,D,None,None,"LACUNA: portali non accessibili da questa sessione","—"),
("Nuovi annunci settimana","Mestre",None,D,None,None,"LACUNA","—"),
("Annunci rimossi/venduti settimana","Mestre",None,D,None,None,"LACUNA","—"),
("% annunci con ribasso","Mestre",None,D,None,None,"LACUNA","—"),
("Ribasso medio","Mestre",None,D,None,None,"LACUNA","—"),
("Permanenza media online (giorni)","Mestre",None,D,None,None,"LACUNA","—"),
]
for i,r in enumerate(od,4):
    for j,v in enumerate(r,1): o.cell(row=i,column=j,value=v)
    o.cell(row=i,column=6,value=f'=IF(OR(C{i}="",E{i}=""),"n.d.",IF(ABS(E{i})<1,C{i}-E{i},C{i}-E{i}))')
    if i>=12: 
        o.cell(row=i,column=3).fill=INP; o.cell(row=i,column=5).fill=INP
for r in (5,7,10,11,15,16): 
    o.cell(row=r,column=3).number_format='0.0%'; o.cell(row=r,column=5).number_format='0.0%'; o.cell(row=r,column=6).number_format='+0.0%;-0.0%'
o["A19"]="Inventario: trend (da Storico)"; o["C19"]='=IF(COUNT(Storico_Settimanale!E3:E200)<2,"n.d. (serve 2ª settimana)",IF(Storico_Settimanale!F3>0.02,"in aumento",IF(Storico_Settimanale!F3<-0.02,"in diminuzione","stabile")))'
o["A21"]="Segmentazione (bilocali, trilocali, pregio, nuove costruzioni, da ristrutturare): da compilare con export annunci."
widths(o,[50,10,12,10,14,12,48,70])

# ---------- Mutui e tassi ----------
m=wb.create_sheet("Mutui_Tassi")
m["A1"]="Mutui e tassi – mutuo tipo €250.000, anticipo 20%, 25 anni"; m["A1"].font=T
hdr(m,3,["Indicatore","Valore attuale","Data","Valore precedente","Data prec.","Variazione (p.p.)","Link"])
md=[
("Euribor 1M",0.02519,"18/11/2026",None,"",S["FAC18"]),
("Euribor 3M",0.02825,"18/11/2026",0.02784,"17/11/2026",S["FAC18"]),
("Euribor 6M",0.03016,"18/11/2026",None,"",S["FAC18"]),
("Euribor 12M",0.03351,"18/11/2026",None,"",S["FAC18"]),
("IRS 10A",0.0344,"18/11/2026",None,"",S["FAC18"]),
("IRS 20A",0.0347,"18/11/2026",None,"",S["FAC18"]),
("IRS 25A",0.0341,"18/11/2026",0.0332,"10/11/2026",S["FAC10"]),
("IRS 30A",0.0333,"18/11/2026",None,"",S["FAC18"]),
("BCE tasso depositi",0.025,"29/10/2026",0.025,"10/09/2026",S["BCE"]),
("TAEG medio nuovi mutui casa (BdI)",0.0405,"set-2026",0.0399,"ago-2026",S["TAEG"]),
("Tasso iniziale medio FISSO (BdI)",0.0367,"set-2026",None,"",S["TAEG"]),
("Tasso iniziale medio VARIABILE (BdI)",0.0329,"set-2026",None,"",S["TAEG"]),
("Miglior offerta online fisso (TAN da)",0.0299,"nov-2026",None,"",S["MUT"]),
("Miglior offerta online variabile (TAN da)",0.0255,"nov-2026",None,"",S["MUT"]),
]
for i,r in enumerate(md,4):
    m.cell(row=i,column=1,value=r[0]); m.cell(row=i,column=2,value=r[1]); m.cell(row=i,column=3,value=r[2])
    m.cell(row=i,column=4,value=r[3]); m.cell(row=i,column=5,value=r[4]); m.cell(row=i,column=7,value=r[5])
    m.cell(row=i,column=6,value=f'=IF(D{i}="","n.d.",(B{i}-D{i})*100)')
    for c in (2,4): m.cell(row=i,column=c).number_format='0.000%'
    m.cell(row=i,column=6).number_format='+0.00;-0.00'
m["A20"]="Calcolo rata"; m["A20"].font=Font(bold=True)
calc=[("Prezzo acquisto (€)",250000),("Anticipo %",0.20),("Importo mutuo (€)","=B21*(1-B22)"),("Durata (anni)",25),
("Spread ipotetico su IRS (fisso)",0.009),("Tasso fisso indicativo","=B14"),("Tasso variabile indicativo","=B15"),
("Rata mensile FISSO (€)","=PMT(B26/12,B24*12,-B23)"),("Rata mensile VARIABILE (€)","=PMT(B27/12,B24*12,-B23)"),
("Tasso fisso costruito (IRS25+spread)","=B10+B25"),("Rata con IRS25+spread (€)","=PMT(B30/12,B24*12,-B23)"),
("Rata stessa base settimana prec. (IRS25 prec.+spread)","=PMT((D10+B25)/12,B24*12,-B23)"),("Δ rata vs periodo prec. (€/mese)","=B31-B32"),
("Sensitività: +0,25 p.p. sul fisso (€/mese)","=PMT((B26+0.0025)/12,B24*12,-B23)-B28"),
("Reddito netto mensile minimo (rata ≤ 33%)","=B28/0.33")]
for i,(a,b) in enumerate(calc,21):
    m.cell(row=i,column=1,value=a); m.cell(row=i,column=2,value=b)
for r in (21,23,28,29,31,32,33,34,35): m.cell(row=r,column=2).number_format='#,##0.00 €'
for r in (22,25,26,27,30): m.cell(row=r,column=2).number_format='0.00%'
for r in (21,22,24,25): m.cell(row=r,column=2).fill=INP
m["C26"]="BdI: media nuovi contratti; NON è un'offerta personalizzata"
m["A37"]="Accessibilità credito"; m["B37"]='=IF(B13>D13,"in peggioramento",IF(B13<D13,"in miglioramento","stabile"))'
widths(m,[50,16,13,15,13,15,70])

# ---------- Affitti e rendimenti ----------
a=wb.create_sheet("Affitti_Rendimenti")
a["A1"]="Affitti e rendimento lordo"; a["A1"].font=T
hdr(a,3,["Area","Fonte","Canone richiesto €/mq/mese","Periodo","Prezzo richiesto €/mq","Rendimento lordo","Var. canone","Link"])
ad=[("Mestre-Chirignago-Marghera-Catene","Immobiliare.it",12.42,"set-2026",2109,-0.022,S["IMM"]),
("Mestre (media)","=\"Media Immobiliare+idealista\"","=AVERAGE(C4,C6)","","=Prezzi_Quotazioni!C12",None,""),
("Mestre","idealista",13.3,"ago-2026",1983,None,S["IDE"]),
("Mestre – lungo termine (snapshot annunci)","idealista",18.11,"nov-2026",None,None,S["IDA"]),
("Comune di Venezia (totale)","Immobiliare.it",15.51,"set-2026",3423,None,S["IMM"]),
]
for i,r in enumerate(ad,4):
    a.cell(row=i,column=1,value=r[0]); a.cell(row=i,column=2,value=r[1]); a.cell(row=i,column=3,value=r[2]); a.cell(row=i,column=4,value=r[3])
    a.cell(row=i,column=5,value=r[4]); a.cell(row=i,column=7,value=r[5]); a.cell(row=i,column=8,value=r[6])
    a.cell(row=i,column=6,value=f'=IF(E{i}="","n.d.",C{i}*12/E{i})'); a.cell(row=i,column=6).number_format='0.00%'; a.cell(row=i,column=7).number_format='+0.0%;-0.0%'
a["A10"]="Rendimento lordo su prezzo di CHIUSURA stimato Mestre"; a["F10"]="=C5*12/Prezzi_Quotazioni!C25"; a["F10"].number_format='0.00%'
a["A12"]="Note normative (verificare sul sito del Comune di Venezia):"; a["A12"].font=Font(bold=True)
notes=["Locazione tradizionale 4+4 / concordato 3+2: canale prevalente in terraferma; verificare accordi territoriali per cedolare 10%.",
"Transitoria (1-18 mesi): richiede esigenza documentata; utile per studenti/lavoratori in trasferta.",
"Breve/turistica: Centro storico soggetto a regolamento comunale (limite base 120 gg/anno; nuove iscrizioni 'premium' bloccate fino al 31/12/2026). Secondo fonti di settore il limite non si applica a Mestre/Marghera. CIN obbligatorio.",
"Il 'snapshot' idealista lungo termine (18,11 €/mq) è un campione di annunci correnti con unità piccole sovrarappresentate: non confrontabile con le medie mensili."]
for i,n in enumerate(notes,13): a.cell(row=i,column=1,value="• "+n)
a["H15"]=S["STR"]
widths(a,[50,18,16,10,16,14,11,70])

# ---------- Storico ----------
s=wb.create_sheet("Storico_Settimanale")
s["A1"]="Storico settimanale (aggiungere una riga ogni mercoledì; le formule si estendono copiando la riga)"; s["A1"].font=Font(bold=True)
hdr(s,2,["Data report","Prezzo rich. Mestre €/mq","Var. sett.","Var. 4 sett.","Annunci attivi Mestre","Var. sett. annunci","Canone Mestre €/mq","IRS 25A","Euribor 3M","Sconto medio","Tempi vendita (mesi)","Var. 12 sett. prezzo","Var. 52 sett. prezzo","Leverage compratori (0-100)","Lettura"])
# newest on top? Use chronological downward; formulas reference rows above.
r=3
s.cell(row=r,column=1,value=D)
s.cell(row=r,column=2,value="=Prezzi_Quotazioni!C12")
s.cell(row=r,column=5,value=None); s.cell(row=r,column=5).fill=INP
s.cell(row=r,column=7,value="=Affitti_Rendimenti!C5")
s.cell(row=r,column=8,value="=Mutui_Tassi!B10"); s.cell(row=r,column=9,value="=Mutui_Tassi!B5")
s.cell(row=r,column=10,value="=Offerta_Ribassi!C5"); s.cell(row=r,column=11,value="=Offerta_Ribassi!C6")
for rr in range(3,60):
    s.cell(row=rr,column=3,value=f'=IF(OR(B{rr}="",ROW()-1<3),"n.d.",IFERROR(B{rr}/B{rr-1}-1,"n.d."))' if rr>3 else "n.d.")
    s.cell(row=rr,column=4,value=f'=IFERROR(IF(B{rr}="","",B{rr}/INDEX(B:B,ROW()-4)-1),"n.d.")' if rr>=7 else "n.d.")
    s.cell(row=rr,column=12,value=f'=IFERROR(IF(B{rr}="","",B{rr}/INDEX(B:B,ROW()-12)-1),"n.d.")' if rr>=15 else "n.d.")
    s.cell(row=rr,column=13,value=f'=IFERROR(IF(B{rr}="","",B{rr}/INDEX(B:B,ROW()-52)-1),"n.d.")' if rr>=55 else "n.d.")
    s.cell(row=rr,column=6,value=f'=IFERROR(IF(OR(E{rr}="",E{rr-1}=""),"n.d.",E{rr}/E{rr-1}-1),"n.d.")' if rr>3 else "n.d.")
    s.cell(row=rr,column=14,value=f'=IF(A{rr}="","",Decisioni!$C$20)' if rr==3 else f'=IF(A{rr}="","",Decisioni!$C$20)')
    s.cell(row=rr,column=15,value=f'=IF(N{rr}="","",IF(N{rr}>=55,"Favorevole compratori",IF(N{rr}<=45,"Favorevole venditori","Equilibrato")))')
    for c in (3,4,6,12,13): s.cell(row=rr,column=c).number_format='+0.0%;-0.0%'
    s.cell(row=rr,column=8).number_format='0.00%'; s.cell(row=rr,column=9).number_format='0.000%'; s.cell(row=rr,column=10).number_format='0.0%'
s["A62"]="NB: nelle settimane successive inserire i valori (non formule) nelle righe nuove e trasformare in valori la riga precedente (Incolla speciale → Valori), così lo storico non si sovrascrive. Colonna N: congelare il valore a fine settimana."
widths(s,[12,14,9,9,12,11,12,9,10,10,10,11,11,14,22])

# ---------- Decisioni ----------
d=wb.create_sheet("Decisioni")
d["A1"]="Decisioni operative e azioni (non è consulenza finanziaria personalizzata)"; d["A1"].font=T
hdr(d,3,["Area / oggetto","Giudizio","Confidenza","Evidenza 1","Evidenza 2","Evidenza 3","Rischio principale","Azione entro 7 giorni"])
dd=[
("Mestre (area prioritaria Domia)","MANTENERE","Media","Prezzi richiesti ~2.100 €/mq, stabili; saldo prezzi BdI 0 p.p. (da +7)","Sconto medio 7,7% (da 7,1%), tempi 5,5 mesi","Rendimento lordo potenziale ~7,1% (12,42 €/mq canone)","Rialzo BCE a dicembre (prob. mercato ~86%) frena domanda a mutuo (69% acquisti)","Estrarre da Geopoi le quotazioni OMI 1° sem 2026 delle zone di Mestre e inserirle in Prezzi_Quotazioni!C19"),
("Venezia centro storico","MANTENERE","Bassa","Comune ~3.423 €/mq, San Marco-Rialto 6.032 €/mq (richiesti)","Canone comune 15,51 €/mq","Regolamento affitti brevi 120 gg; premium bloccato a fine 2026","Cambi normativi su locazioni turistiche","Verificare sul sito del Comune lo stato del regolamento per il 2027"),
("Provincia di Venezia","MANTENERE","Bassa","Immobiliare.it 2.815 €/mq, +4,0% a/a (giu-2026)","RealAdvisor 2.058 €/mq, +4,2% a/a (AVM, ago-2026)","Compravendite Italia Q2 +0,1% a/a","Dati provinciali vecchi di 3-5 mesi e metodologie non confrontabili","Definire i comuni da monitorare (CAP) per sostituire il dato provinciale"),
("Mutuo (se in essere a variabile)","RINEGOZIARE O SURROGARE – da valutare","Media","Euribor 3M 2,825% in salita (+4,1 bp in un giorno)","IRS 25A 3,41% (+9 bp vs 10/11)","Mercato prezza rialzo BCE a dicembre","Bloccare un fisso proprio prima di un'eventuale inversione","Chiedere 2 preventivi surroga a banca/mediatore creditizio abilitato; confrontare TAEG"),
("Il mio immobile","ATTENDERE","Bassa","Superficie/indirizzo/stato non forniti: stima solo illustrativa","Valore illustrativo 80 mq Mestre ≈ 150-173 k€ (centrale ~162 k€) (vedi Mio_Immobile)","Domanda e offerta in calo (BdI)","Decisione su dati incompleti","Compilare le celle gialle del foglio Mio_Immobile (mq, CAP, stato, mutuo)"),
]
for i,r in enumerate(dd,4):
    for j,v in enumerate(r,1):
        c=d.cell(row=i,column=j,value=v); c.alignment=Alignment(wrap_text=True,vertical="top")
d["A11"]="Indicatore sintetico di leverage compratori/venditori (0 = venditori, 100 = compratori)"; d["A11"].font=Font(bold=True)
hdr(d,12,["Componente","Valore","Benchmark neutro","Punti per unità","Contributo"])
comp=[("Sconto medio (p.p.)","=Offerta_Ribassi!C5*100",9,4),
("Tempi di vendita (mesi)","=Offerta_Ribassi!C6",7,4),
("Saldo prezzi agenti (p.p.)","=Offerta_Ribassi!C4",5,-0.8),
("Saldo attese trimestre (p.p.)","=Offerta_Ribassi!C8",-5,-0.5),
("TAEG medio mutui (%)","=Mutui_Tassi!B13*100",3.5,6),
("Var. a/a prezzi IPAB (%)","=Prezzi_Quotazioni!F20*100",3,-1.5)]
for i,(a1,b,c,e) in enumerate(comp,13):
    d.cell(row=i,column=1,value=a1); d.cell(row=i,column=2,value=b); d.cell(row=i,column=3,value=c); d.cell(row=i,column=4,value=e)
    d.cell(row=i,column=5,value=f"=(B{i}-C{i})*D{i}")
    for cc in (3,4): d.cell(row=i,column=cc).fill=INP
d["A20"]="Leverage compratori (0-100)"; d["C20"]="=MAX(0,MIN(100,50+SUM(E13:E18)))"; d["A20"].font=Font(bold=True)
d["A21"]="Semaforo"; d["C21"]='=IF(C20>=55,"FAVOREVOLE COMPRATORI",IF(C20<=45,"FAVOREVOLE VENDITORI","EQUILIBRATO"))'
d["A22"]="Benchmark neutri ≈ medie storiche BdI (sconto ~9%, tempi ~7 mesi); pesi e benchmark sono ipotesi modificabili (celle gialle)."
widths(d,[32,22,14,40,40,40,40,48])
for i in range(4,9): d.row_dimensions[i].height=75

# ---------- Dati area ----------
ar=wb.create_sheet("Dati_Area")
ar["A1"]="Anagrafica aree monitorate"; ar["A1"].font=T
hdr(ar,3,["Area","Tipo","CAP","Zona OMI","Priorità","Prezzo rich. €/mq","Canone €/mq","Rend. lordo","Giudizio","Note"])
arr=[("Mestre","Quartiere/terraferma","30170-30175","DA INSERIRE","Alta (Domia)","=Prezzi_Quotazioni!C12","=Affitti_Rendimenti!C5","=G4*12/F4","=Decisioni!B4",""),
("Venezia centro storico","Città","30121-30135","DA INSERIRE","Alta (Domia)","=Prezzi_Quotazioni!C15","=Affitti_Rendimenti!C8","=G5*12/F5","=Decisioni!B5","dato comune intero incl. terraferma"),
("Provincia di Venezia","Provincia","—","—","Alta (Domia)","=Prezzi_Quotazioni!C17","","","=Decisioni!B6",""),
("[INSERISCI CAP]","CAP","","","","","","","","placeholder non compilato nel prompt"),
("[INSERISCI ZONE]","Comune/quartiere","","","","","","","",""),
("[INSERISCI CITTÀ]","Area metropolitana","","","","","","","","")]
for i,r in enumerate(arr,4):
    for j,v in enumerate(r,1): ar.cell(row=i,column=j,value=v)
    ar.cell(row=i,column=6).number_format='#,##0'; ar.cell(row=i,column=8).number_format='0.0%'
ar["H6"]=None
widths(ar,[26,20,14,14,14,14,12,11,26,40])

# ---------- Dashboard ----------
db=wb.create_sheet("Dashboard",0)
db["A1"]="Dashboard mercato immobiliare – Mestre / Venezia"; db["A1"].font=Font(bold=True,size=16)
db["A2"]="Data report"; db["B2"]=D
db["A3"]="Semaforo mercato"; db["B3"]="=Decisioni!C21"; db["C3"]="=Decisioni!C20"; db["D3"]="indice leverage 0-100"
hdr(db,5,["KPI","Valore","Natura","Var. vs sett. prec.","Fonte/periodo"])
k=[("Prezzo richiesto Mestre (mediana fonti) €/mq","=Prezzi_Quotazioni!C12","richiesto","=Storico_Settimanale!C3","portali, ago-set 2026"),
("Prezzo chiusura stimato Mestre €/mq","=Prezzi_Quotazioni!C25","stimato","n.d. (1ª settimana)","richiesto × (1-sconto BdI)"),
("IPAB Italia a/a","=Prezzi_Quotazioni!F20","transato (indice)","invariato (trimestrale)","ISTAT Q2-2026"),
("Annunci attivi Mestre","=IF(Offerta_Ribassi!C12=\"\",\"n.d.\",Offerta_Ribassi!C12)","annunci","n.d.","LACUNA"),
("Nuovi annunci settimana","=IF(Offerta_Ribassi!C13=\"\",\"n.d.\",Offerta_Ribassi!C13)","annunci","n.d.","LACUNA"),
("Sconto medio richiesta→chiusura","=Offerta_Ribassi!C5","sondaggio","+0,6 p.p. (vs trim. prec.)","BdI Q3-2026"),
("Tempi di vendita (mesi)","=Offerta_Ribassi!C6","sondaggio","stabile (vs trim. prec.)","BdI Q3-2026"),
("Saldo attese prezzi (p.p.)","=Offerta_Ribassi!C8","sondaggio","-2 p.p. (vs trim. prec.)","BdI Q3-2026"),
("IRS 25A","=Mutui_Tassi!B10","mercato","=TEXT(Mutui_Tassi!F10,\"+0.00;-0.00\")&\" p.p. vs 10/11\"","Facile.it 18/11"),
("Euribor 3M","=Mutui_Tassi!B5","mercato","=TEXT(Mutui_Tassi!F5,\"+0.000;-0.000\")&\" p.p. vs 17/11\"","Facile.it 18/11"),
("TAEG medio nuovi mutui","=Mutui_Tassi!B13","media BdI","+0,06 p.p. (vs ago)","BdI set-2026"),
("Rata mutuo tipo (fisso) €/mese","=Mutui_Tassi!B28","calcolo","n.d.","200k€, 25 anni"),
("Canone richiesto Mestre €/mq/mese","=Affitti_Rendimenti!C5","richiesto","n.d.","portali"),
("Rendimento lordo Mestre","=Affitti_Rendimenti!F5","calcolo","n.d.","canone/prezzo richiesto")]
for i,r in enumerate(k,6):
    for j,v in enumerate(r,1): db.cell(row=i,column=j,value=v)
fmt={6:'#,##0',7:'#,##0',8:'0.0%',11:'0.0%',12:'0.0',13:'0',14:'0.00%',15:'0.000%',16:'0.00%',17:'#,##0 €',18:'0.00',19:'0.0%'}
for r,f in fmt.items(): db.cell(row=r,column=2).number_format=f
db["A21"]="Sintesi"; db["A21"].font=Font(bold=True)
summ=["1) Prezzi: crescita ferma (saldo agenti BdI da +7 a 0); IPAB Italia +4,0% a/a ma in rallentamento.",
"2) Trattabilità in aumento: sconto medio 7,7% (+0,6 p.p.), tempi stabili ~5,5 mesi; attese a 2 anni in netto peggioramento (-21).",
"3) Credito più caro: TAEG 4,05% (set), Euribor 3M 2,83%, IRS 25A 3,41%; BCE ferma al 2,50% ma rialzo a dicembre prezzato.",
"4) Mestre: ~2.100 €/mq richiesti, canoni ~12-13 €/mq, rendimento lordo ~7%: tra i più alti del comune.",
"5) Lacune: inventario/annunci/ribassi non rilevabili questa settimana (portali non raggiungibili); prima settimana di storico."]
for i,t in enumerate(summ,22): db.cell(row=i,column=1,value=t)
db["A28"]="Decisione consigliata"; db["A28"].font=Font(bold=True)
db["B28"]="Mestre: MANTENERE (media) · Mio immobile: ATTENDERE (dati mancanti) · Mutuo variabile: valutare surroga/rinegoziazione"
widths(db,[44,22,18,32,30])
# Fonti
f=wb.create_sheet("Fonti")
hdr(f,1,["Codice","URL"])
for i,(kk,v) in enumerate(S.items(),2): f.cell(row=i,column=1,value=kk); f.cell(row=i,column=2,value=v)
widths(f,[10,110])
wb.save(out)
