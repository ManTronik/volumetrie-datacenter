#!/usr/bin/env python3
"""Genere le classeur volumetrie_datacenter.xlsx (formules vivantes)."""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "volumetrie_datacenter.xlsx"

TITLE_FONT = Font(bold=True, size=14)
HDR_FONT = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="4472C4")
SEC_FONT = Font(bold=True, size=12)
SEC_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Border(*[Side(style="thin", color="B0B0B0")] * 4)
PCT = "0%"
NUM1 = "#,##0.0"
NUM0 = "#,##0"

wb = Workbook()


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.border = THIN
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------- Mode d'emploi
ws = wb.active
ws.title = "Mode d'emploi"
set_widths(ws, [110])
lines = [
    ("VOLUMETRIE RESEAU DATACENTER - MODE D'EMPLOI", TITLE_FONT),
    ("", None),
    ("Principe : calcul BOTTOM-UP par POD, calibre sur les mesures disponibles (releves FAI, SNMP coeur de reseau, compteurs de stockage).", None),
    ("Multi-datacenter : la colonne 'Datacenter' de l'inventaire permet de consolider par site (onglet Synthese DC).", None),
    ("", None),
    ("1. Onglet Parametres : facteurs globaux (simultaneite, crete/moyenne, croissance, marge, horizon) et noms des datacenters.", None),
    ("2. Onglet Profils : un profil par categorie applicative = debit moyen par serveur + repartition du flux", None),
    ("   (sortant internet / entrant internet / interne LAN / replication inter-DC). Les parts doivent totaliser 100 %.", None),
    ("3. Onglet Inventaire : une ligne par POD. Renseignez datacenter, salle, POD, nb de baies, serveurs/baie et categorie.", None),
    ("   Le nb de serveurs et les debits se calculent seuls. Ajoutez autant de lignes que necessaire", None),
    ("   (les formules sont pre-posees jusqu'a la ligne 100).", None),
    ("4. Onglet Resultats : volumetrie globale par direction (moyenne, crete, volumes mensuels), projection a l'horizon", None),
    ("   choisi et capacite de dimensionnement recommandee.", None),
    ("5. Onglet Synthese DC : volumetrie par datacenter, trafic WAN inter-DC (decompte simple vs double comptage)", None),
    ("   et scenario de bascule N+1 (choix du DC defaillant et taux de bascule).", None),
    ("", None),
    ("Regles de conversion utilisees :", SEC_FONT),
    ("   1 Mbps soutenu pendant 1 mois de 30 jours = 0,324 To       (1 Mbps = 0,125 Mo/s)", None),
    ("   1 Gb/s soutenu pendant 1 mois de 30 jours = 324 To", None),
    ("", None),
    ("Points de vigilance :", SEC_FONT),
    ("   - Ne jamais sommer les debits nominaux des NIC : utiliser les profils mesures (5-15 % d'utilisation moyenne).", None),
    ("   - La crete agregee est calculee POD par POD puis sommee : c'est le 95e centile approche, base de dimensionnement.", None),
    ("   - Ne pas oublier les flux transverses : sauvegardes, replication, mises a jour, monitoring (souvent 30-60 % du total).", None),
    ("   - Le trafic Est-Ouest interne ne sort jamais du datacenter : bien le distinguer du trafic internet (N-S).", None),
    ("   - La replication entre datacenters est comptee aux DEUX extremites : le total global double-compte ce flux,", None),
    ("     l'onglet Synthese DC fournit le decompte simple (volume unique) et le decompte double (WAN a provisionner).", None),
    ("   - Si un site n'a pas de sortie internet locale (backhaul vers un site principal), classer son trafic en WAN", None),
    ("     (reduction de la part 'sortant internet' de ses profils) plutot qu'en internet.", None),
    ("", None),
    ("Calibration conseillee : si vous relevez X Mbps sortants internet chez le FAI et Y Mbps modelises ici,", None),
    ("appliquez Y/X comme coefficient correcteur sur la colonne 'Debit moyen / serveur' des profils.", None),
]
for i, (txt, f) in enumerate(lines, start=1):
    c = ws.cell(row=i, column=1, value=txt)
    if f:
        c.font = f
    c.alignment = Alignment(wrap_text=False)

# ---------------------------------------------------------------- Parametres
ws = wb.create_sheet("Parametres")
set_widths(ws, [46, 14, 52])
ws["A1"] = "PARAMETRES GLOBAUX"
ws["A1"].font = TITLE_FONT
params = [
    ("Jours par mois", 30, "", NUM0),
    ("Facteur de simultaneite (part des serveurs actifs en meme temps)", 0.5, "0,3 a 0,7 - evite de sommer les cretes de tous les serveurs", PCT),
    ("Ratio crete / moyenne", 4, "typiquement 3 a 10 selon activite", NUM1),
    ("Croissance annuelle", 0.15, "croissance du parc et du trafic par serveur", PCT),
    ("Horizon de projection (annees)", 3, "", NUM0),
    ("Marge de securite", 0.30, "marge ajoutee au dimensionnement final", PCT),
    ("Sur-souscription cible access -> coeur", 4, "4 = 4 Gb/s d'access pour 1 Gb/s d'uplink (typique 3:1 a 5:1)", NUM0),
]
for r, (label, val, note, fmt) in enumerate(params, start=3):
    ws.cell(row=r, column=1, value=label)
    c = ws.cell(row=r, column=2, value=val)
    c.number_format = fmt
    c.border = THIN
    c.fill = PatternFill("solid", fgColor="FFF2CC")
    ws.cell(row=r, column=3, value=note).font = Font(italic=True, color="808080")

ws.cell(row=11, column=1, value="Noms des datacenters (utilises par l'inventaire et la synthese)").font = SEC_FONT
for c in range(1, 4):
    ws.cell(row=11, column=c).fill = SEC_FILL
for i in range(3):
    r = 12 + i
    ws.cell(row=r, column=1, value=f"Nom Datacenter {i+1}")
    c = ws.cell(row=r, column=2, value=f"DC{i+1}")
    c.border = THIN
    c.fill = PatternFill("solid", fgColor="FFF2CC")

P_JOURS = "Parametres!$B$3"
P_SIM = "Parametres!$B$4"
P_CRETE = "Parametres!$B$5"
P_CROIS = "Parametres!$B$6"
P_HORIZON = "Parametres!$B$7"
P_MARGE = "Parametres!$B$8"
P_SURSUB = "Parametres!$B$9"

# ---------------------------------------------------------------- Profils
ws = wb.create_sheet("Profils")
set_widths(ws, [28, 16, 15, 15, 15, 17, 40])
ws["A1"] = "PROFILS PAR CATEGORIE APPLICATIVE"
ws["A1"].font = TITLE_FONT
headers = ["Categorie", "Debit moyen / serveur (Mbps)", "Part sortant internet", "Part entrant internet",
           "Part interne LAN", "Part replication inter-DC", "Commentaire"]
for c, h in enumerate(headers, start=1):
    ws.cell(row=2, column=c, value=h)
style_header(ws, 2, len(headers))
profils = [
    ("Web / exposition publique", 30, 0.70, 0.15, 0.10, 0.05, "fortement asymetrique, sortant dominant"),
    ("Applicatif metier", 20, 0.05, 0.10, 0.70, 0.15, "flux clients internes dominants"),
    ("Base de donnees", 25, 0.02, 0.03, 0.80, 0.15, "Est-Ouest dominant"),
    ("Sauvegarde / stockage", 60, 0.00, 0.00, 0.40, 0.60, "pics nocturnes, replication lourde"),
    ("Administration / outils", 5, 0.10, 0.10, 0.75, 0.05, "faible volumetrie"),
]
for r, row in enumerate(profils, start=3):
    for c, val in enumerate(row, start=1):
        cell = ws.cell(row=r, column=c, value=val)
        cell.border = THIN
        if c == 2:
            cell.number_format = NUM0
        elif 3 <= c <= 6:
            cell.number_format = PCT
for r in range(len(profils) + 3, 21):
    for c in range(1, 7):
        ws.cell(row=r, column=c).border = THIN
ws.cell(row=21, column=1, value="Les 4 parts de chaque ligne doivent totaliser 100 % (controle colonne H ci-dessous).").font = Font(italic=True, color="808080")
ws["H2"] = "Controle"
ws["H2"].font = HDR_FONT
ws["H2"].fill = PatternFill("solid", fgColor="C00000")
ws["H2"].border = THIN
for r in range(3, 21):
    ws.cell(row=r, column=8, value=f'=IF(COUNTA(A{r})=0,"",IF(ROUND(SUM(C{r}:F{r}),4)=1,"OK","ERREUR : somme <> 100 %"))')

# ---------------------------------------------------------------- Inventaire
ws = wb.create_sheet("Inventaire")
set_widths(ws, [13, 12, 10, 10, 13, 12, 26, 14, 14, 14, 14, 14, 14])
ws["A1"] = "INVENTAIRE - une ligne par POD"
ws["A1"].font = TITLE_FONT
inv_headers = ["Datacenter", "Salle", "POD", "Nb baies", "Serveurs / baie", "Nb serveurs", "Categorie",
               "Trafic moyen (Mbps)", "Trafic crete (Mbps)",
               "Sortant internet (Mbps)", "Entrant internet (Mbps)", "Interne LAN (Mbps)", "Replication inter-DC (Mbps)"]
for c, h in enumerate(inv_headers, start=1):
    ws.cell(row=2, column=c, value=h)
style_header(ws, 2, len(inv_headers))
exemple = [
    ("DC1", "Salle A", "A1", 24, 16, "Web / exposition publique"),
    ("DC1", "Salle A", "A2", 24, 16, "Applicatif metier"),
    ("DC1", "Salle A", "A3", 24, 12, "Base de donnees"),
    ("DC1", "Salle A", "A4", 24, 8, "Sauvegarde / stockage"),
    ("DC2", "Salle B", "B1", 24, 16, "Web / exposition publique"),
    ("DC2", "Salle B", "B2", 24, 16, "Applicatif metier"),
    ("DC2", "Salle B", "B3", 24, 10, "Base de donnees"),
    ("DC2", "Salle B", "B4", 24, 6, "Administration / outils"),
    ("DC3", "Salle C", "C1", 24, 14, "Web / exposition publique"),
    ("DC3", "Salle C", "C2", 24, 12, "Applicatif metier"),
]
for i, (dc, salle, pod, baies, spb, cat) in enumerate(exemple):
    r = 3 + i
    ws.cell(row=r, column=1, value=dc)
    ws.cell(row=r, column=2, value=salle)
    ws.cell(row=r, column=3, value=pod)
    ws.cell(row=r, column=4, value=baies)
    ws.cell(row=r, column=5, value=spb)
    ws.cell(row=r, column=7, value=cat)
for r in range(3, 101):
    ws.cell(row=r, column=6, value=f'=IF(COUNTA(B{r}:E{r})=0,"",D{r}*E{r})')
    ws.cell(row=r, column=8,
            value=f'=IF($F{r}="","",ROUND($F{r}*VLOOKUP($G{r},Profils!$A$3:$B$20,2,FALSE)*{P_SIM},1))')
    ws.cell(row=r, column=9, value=f'=IF($H{r}="","",ROUND($H{r}*{P_CRETE},1))')
    for col, pcol in ((10, 3), (11, 4), (12, 5), (13, 6)):
        ws.cell(row=r, column=col,
                value=f'=IF($H{r}="","",ROUND($H{r}*VLOOKUP($G{r},Profils!$A$3:$F$20,{pcol},FALSE),1))')
dv_dc = DataValidation(type="list", formula1="Parametres!$B$12:$B$14", allow_blank=True)
ws.add_data_validation(dv_dc)
dv_dc.add("A3:A100")
dv_cat = DataValidation(type="list", formula1="Profils!$A$3:$A$20", allow_blank=True)
ws.add_data_validation(dv_cat)
dv_cat.add("G3:G100")

# ---------------------------------------------------------------- Resultats
ws = wb.create_sheet("Resultats")
set_widths(ws, [34, 16, 16, 20, 22, 40])
ws["A1"] = "VOLUMETRIE GLOBALE - RESULTATS"
ws["A1"].font = TITLE_FONT

DIRECTIONS = [("Sortant internet", "J"), ("Entrant internet", "K"),
              ("Interne LAN (Est-Ouest)", "L"), ("Replication inter-DC", "M")]

ws["A3"] = "1. VOLUMETRIE ACTUELLE (tous datacenters)"
ws["A3"].font = SEC_FONT
for c in range(1, 7):
    ws.cell(row=3, column=c).fill = SEC_FILL
hdr = ["Flux", "Moyen (Mbps)", "Crete (Mbps)", "Volume mensuel moyen (To)", "Volume mensuel si crete permanente (To)", "Commentaire"]
for c, h in enumerate(hdr, start=1):
    ws.cell(row=4, column=c, value=h)
style_header(ws, 4, len(hdr))
r = 5
for label, col in DIRECTIONS:
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=2, value=f"=SUM(Inventaire!{col}3:{col}100)").number_format = NUM0
    ws.cell(row=r, column=3, value=f"=B{r}*{P_CRETE}").number_format = NUM0
    ws.cell(row=r, column=4, value=f"=B{r}*0.0108*{P_JOURS}").number_format = NUM1
    ws.cell(row=r, column=5, value=f"=C{r}*0.0108*{P_JOURS}").number_format = NUM1
    r += 1
ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
for c, col in ((2, "B"), (3, "C"), (4, "D"), (5, "E")):
    ws.cell(row=r, column=c, value=f"=SUM({col}5:{col}{r-1})").number_format = NUM1
    ws.cell(row=r, column=c).font = Font(bold=True)
total_row = r

r += 2
ws.cell(row=r, column=1, value="2. PROJECTION A L'HORIZON CHOISI").font = SEC_FONT
for c in range(1, 7):
    ws.cell(row=r, column=c).fill = SEC_FILL
r += 1
ws.cell(row=r, column=1, value="Facteur de croissance")
ws.cell(row=r, column=2, value=f"=(1+{P_CROIS})^{P_HORIZON}").number_format = "0.00"
fact_row = r
r += 1
ws.cell(row=r, column=1, value="Facteur de marge de securite")
ws.cell(row=r, column=2, value=f"=1+{P_MARGE}").number_format = "0.00"
r += 1
hdr2 = ["Flux", "Crete actuelle (Mbps)", "Crete projetee (Mbps)", "Volume mensuel projete (To)", "", ""]
for c, h in enumerate(hdr2, start=1):
    ws.cell(row=r, column=c, value=h)
style_header(ws, r, 4)
r += 1
proj_start = r
for label, col in DIRECTIONS:
    src = 5 + DIRECTIONS.index((label, col))
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=2, value=f"=C{src}").number_format = NUM0
    ws.cell(row=r, column=3, value=f"=B{r}*$B${fact_row}*$B${fact_row+1}").number_format = NUM0
    ws.cell(row=r, column=4, value=f"=B{r}*0.0108*{P_JOURS}*$B${fact_row}").number_format = NUM1
    r += 1
ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
for c, col in ((2, "B"), (3, "C"), (4, "D")):
    ws.cell(row=r, column=c, value=f"=SUM({col}{proj_start}:{col}{r-1})").number_format = NUM1
    ws.cell(row=r, column=c).font = Font(bold=True)
proj_total = r

r += 2
ws.cell(row=r, column=1, value="3. DIMENSIONNEMENT RECOMMANDE").font = SEC_FONT
for c in range(1, 7):
    ws.cell(row=r, column=c).fill = SEC_FILL
r += 1
dim = [
    ("Capacite coeur requise - internet (Gb/s)", f"=(C{proj_start}+C{proj_start+1})/1000", "crete projetee sortant + entrant"),
    ("Capacite coeur requise - interne (Gb/s)", f"=(C{proj_start+2}+C{proj_start+3})/1000", "crete projetee LAN + replication"),
    ("Capacite coeur requise - totale (Gb/s)", f"=C{proj_total}/1000", "somme des cretes projettees"),
    ("Uplink par POD conseille (Gb/s)", f"=ROUNDUP(C{proj_total}/1000/COUNTA(Inventaire!C3:C100),0)", "moyenne par POD, arrondie au Gb/s superieur"),
]
for label, formula, note in dim:
    ws.cell(row=r, column=1, value=label).font = Font(bold=True)
    ws.cell(row=r, column=2, value=formula).number_format = NUM0
    ws.cell(row=r, column=3, value=note).font = Font(italic=True, color="808080")
    r += 1

# ---------------------------------------------------------------- Synthese DC
ws = wb.create_sheet("Synthese DC")
set_widths(ws, [18, 16, 16, 17, 17, 16, 19, 20])
ws["A1"] = "SYNTHESE PAR DATACENTER"
ws["A1"].font = TITLE_FONT

ws["A3"] = "1. VOLUMETRIE PAR DATACENTER"
ws["A3"].font = SEC_FONT
for c in range(1, 9):
    ws.cell(row=3, column=c).fill = SEC_FILL
hdr = ["Datacenter", "Moyen total (Mbps)", "Crete totale (Mbps)", "Sortant internet (Mbps)",
       "Entrant internet (Mbps)", "Interne LAN (Mbps)", "Replication inter-DC (Mbps)", "Volume mensuel moyen (To)"]
for c, h in enumerate(hdr, start=1):
    ws.cell(row=4, column=c, value=h)
style_header(ws, 4, len(hdr))
for i in range(3):
    r = 5 + i
    prow = 12 + i
    ws.cell(row=r, column=1, value=f"=Parametres!$B${prow}").font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=SUMIF(Inventaire!$A$3:$A$100,$A{r},Inventaire!H$3:H$100)").number_format = NUM0
    ws.cell(row=r, column=3, value=f"=SUMIF(Inventaire!$A$3:$A$100,$A{r},Inventaire!I$3:I$100)").number_format = NUM0
    ws.cell(row=r, column=4, value=f"=SUMIF(Inventaire!$A$3:$A$100,$A{r},Inventaire!J$3:J$100)").number_format = NUM0
    ws.cell(row=r, column=5, value=f"=SUMIF(Inventaire!$A$3:$A$100,$A{r},Inventaire!K$3:K$100)").number_format = NUM0
    ws.cell(row=r, column=6, value=f"=SUMIF(Inventaire!$A$3:$A$100,$A{r},Inventaire!L$3:L$100)").number_format = NUM0
    ws.cell(row=r, column=7, value=f"=SUMIF(Inventaire!$A$3:$A$100,$A{r},Inventaire!M$3:M$100)").number_format = NUM0
    ws.cell(row=r, column=8, value=f"=B{r}*0.0108*{P_JOURS}").number_format = NUM1
ws.cell(row=8, column=1, value="GLOBAL").font = Font(bold=True)
for c in range(2, 9):
    col = get_column_letter(c)
    ws.cell(row=8, column=c, value=f"=SUM({col}5:{col}7)").number_format = NUM1
    ws.cell(row=8, column=c).font = Font(bold=True)
ws.cell(row=9, column=1, value="NB : la replication inter-DC est comptee aux deux extremites dans le GLOBAL (voir section 2).").font = Font(italic=True, color="808080")

ws["A11"] = "2. TRAFIC WAN INTER-DC"
ws["A11"].font = SEC_FONT
for c in range(1, 9):
    ws.cell(row=11, column=c).fill = SEC_FILL
wan = [
    ("Replication totale - moyenne (Mbps)", "=G8", "somme des trois sites (double comptage)", NUM0),
    ("Replication totale - crete (Mbps)", f"=G8*{P_CRETE}", "base de dimensionnement WAN", NUM0),
    ("Volume unique - replication (To/mois)", f"=G8*0.0108*{P_JOURS}/2", "decompte simple : flux compté une seule fois", NUM1),
    ("WAN inter-DC a provisionner (Gb/s)", "=ROUNDUP(G8*Parametres!$B$5/1000,0)", "les deux extremites transportent le flux", NUM0),
]
for r, (label, formula, note, fmt) in enumerate(wan, start=12):
    ws.cell(row=r, column=1, value=label).font = Font(bold=True)
    ws.cell(row=r, column=2, value=formula).number_format = fmt
    ws.cell(row=r, column=3, value=note).font = Font(italic=True, color="808080")

ws["A17"] = "3. SCENARIO DE BASCULE (N+1)"
ws["A17"].font = SEC_FONT
for c in range(1, 9):
    ws.cell(row=17, column=c).fill = SEC_FILL
ws["A18"] = "Datacenter defaillant"
ws["B18"] = "DC1"
ws["B18"].border = THIN
ws["B18"].fill = PatternFill("solid", fgColor="FFF2CC")
ws["C18"] = "menu deroulant : noms definis dans Parametres"
ws["C18"].font = Font(italic=True, color="808080")
ws["A19"] = "Taux de bascule"
c = ws["B19"]
c.value = 1.0
c.number_format = PCT
c.border = THIN
c.fill = PatternFill("solid", fgColor="FFF2CC")
ws["C19"] = "1 = PRA complet, 0,5 = bascule de la moitie des charges"
ws["C19"].font = Font(italic=True, color="808080")
hdr = ["Datacenter", "Crete propre (Mbps)", "Crete absorbee (Mbps)", "Crete apres bascule (Mbps)", "Capacite coeur requise (Gb/s)"]
for c, h in enumerate(hdr, start=1):
    ws.cell(row=21, column=c, value=h)
style_header(ws, 21, len(hdr))
for i in range(3):
    r = 22 + i
    src = 5 + i
    prow = 12 + i
    ws.cell(row=r, column=1, value=f"=Parametres!$B${prow}")
    ws.cell(row=r, column=2, value=f"=C{src}").number_format = NUM0
    ws.cell(row=r, column=3,
            value=f"=IF($A{r}=$B$18,0,SUMIF(Inventaire!$A$3:$A$100,$B$18,Inventaire!$I$3:$I$100)*$B$19)").number_format = NUM0
    ws.cell(row=r, column=4, value=f"=B{r}+C{r}").number_format = NUM0
    ws.cell(row=r, column=5, value=f"=D{r}/1000").number_format = NUM1
ws.cell(row=26, column=1, value="Dimensionner chaque site survivant sur sa ligne ; le site defaillant affiche 0.").font = Font(italic=True, color="808080")
dv_fail = DataValidation(type="list", formula1="Parametres!$B$12:$B$14", allow_blank=True)
ws.add_data_validation(dv_fail)
dv_fail.add("B18")

wb.save(OUT)
print(f"OK -> {OUT}")
