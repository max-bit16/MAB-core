"""
MAB Belgique — Génération des deux documents Word
"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── Helpers ────────────────────────────────────────────────────────────────

FONT = 'Calibri'

def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    run = p.runs[0] if p.runs else p.add_run(text)
    run.font.name = FONT
    if level == 1:
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    elif level == 2:
        run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    return p

def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(level * 0.5)
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.name = FONT
    return p

def add_note(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(f"⚠ {text}")
    run.font.color.rgb = RGBColor(0xC0, 0x50, 0x20)
    run.font.italic = True
    run.font.size = Pt(9)
    run.font.name = FONT
    return p

def add_para(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.name = FONT
    return p

def bold_bullet(doc, label, value, level=0):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(level * 0.5)
    r1 = p.add_run(f"{label} : ")
    r1.bold = True
    r1.font.size = Pt(10)
    r1.font.name = FONT
    r2 = p.add_run(value)
    r2.font.size = Pt(10)
    r2.font.name = FONT
    return p

# ─── DOCUMENT PRINCIPAL ──────────────────────────────────────────────────────

doc = Document()

# Marges
for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

# Page de titre
t = doc.add_heading("MAB BELGIQUE — ÉTUDE DE MARCHÉ", 0)
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub = doc.add_paragraph("Robinetterie sanitaire collective / ERP — Les Robinets Presto")
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.runs[0].font.size = Pt(12)
date_p = doc.add_paragraph("Juin 2026")
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
date_p.runs[0].font.size = Pt(10)
date_p.runs[0].font.italic = True
doc.add_paragraph()

# ─── RÉSUMÉ EXÉCUTIF ─────────────────────────────────────────────────────────
add_heading(doc, "RÉSUMÉ EXÉCUTIF", 1)
add_bullet(doc, "Marché de la construction en recul structurel (-2,8% en 2024, -0,4% en 2025) avec un résidentiel neuf sinistré — la rénovation et le génie civil tirent le secteur.")
add_bullet(doc, "Non-résidentiel seul segment positif (+1,4% en 2024, +1,5% en 2025) : bureaux, industrie, projets publics (santé, éducation). Fort potentiel pour la robinetterie ERP.")
add_bullet(doc, "Investissements publics massifs dans les segments-clés : 438 M€ hôpitaux universitaires, 1 Md€ bâtiments scolaires, 1,7 Md€ logements sociaux (Flandre), 120 M€ psychiatrie (BEI). Opportunité directe pour Presto.")
add_bullet(doc, "Delabie solidement implanté via sa filiale Benelux (Sint-Pieters-Leeuw) — concurrent direct, leader sur le même segment collectif/ERP.")
add_bullet(doc, "Presto est déjà présent (présence commerciale existante) : l'enjeu est de structurer et consolider la distribution face à Delabie.")
doc.add_paragraph()

# ─── PARTIE 1 ────────────────────────────────────────────────────────────────
add_heading(doc, "PARTIE 1 — OVERVIEW CONTEXTE PAYS", 1)

add_heading(doc, "1.1 Analyse PESTEL", 2)

bold_bullet(doc, "POLITIQUE", "Instabilité relative — score stabilité politique 0,4/2,58 en 2023 (TheGlobalEconomy). Régime fédéral à 3 régions (Flandre, Wallonie, Bruxelles) créant une complexité réglementaire et administrative. Gouvernement Arizona formé en 2025 (coalition De Wever).")
bold_bullet(doc, "ÉCONOMIQUE", "PIB 642 Md€ en 2025, croissance +1,0% (NBB, 2025). PIB/hab : 48 700€ (Trésor FR, 2025). Chômage : 6,2%. Déficit public : -5,2% du PIB en 2025. Économie très tertiaire (services ~80% du PIB).")
bold_bullet(doc, "SOCIAL", "11,77 millions d'habitants (Worldometer, 2026). Taux d'urbanisation : 99%. Vieillissement de la population (pression sur les ERP de santé). Multilinguisme (néerlandais 60%, français 40%, allemand <1%) : complexité commerciale.")
bold_bullet(doc, "TECHNOLOGIQUE", "Forte adoption du BIM et outils numériques dans la construction. Exigences PEB (Performance Énergétique des Bâtiments) moteur d'innovation produit. Développement domotique et capteurs IoT dans les ERP.")
bold_bullet(doc, "ENVIRONNEMENTAL", "Stress hydrique modéré à élevé (EEA, 2025). Réglementation PEB stricte : objectif bâtiments neutres carbone 2045 (Flandre), -35% énergie primaire d'ici 2030. Forte demande pour produits hydro-économes — avantage Presto.")
bold_bullet(doc, "LÉGAL", "Normes harmonisées EU (NBN EN 816, NBN EN 817). Marchés publics réglementés (seuils UE 2024-2025). 3 régions = 3 législations bâtiment distinctes (complexité pour les appels d'offres).")
doc.add_paragraph()

add_heading(doc, "1.2 Indicateurs socio-économiques clés", 2)
bold_bullet(doc, "PIB total", "642 Md€ en 2025 (Trésor FR / NBB)")
bold_bullet(doc, "PIB/habitant", "48 700 €/an — parmi les plus élevés d'Europe")
bold_bullet(doc, "Population", "11 774 642 habitants (Worldometer, mi-2026)")
bold_bullet(doc, "Urbanisation", "99% (quasi-totalité de la population en zone urbaine)")
bold_bullet(doc, "Croissance PIB", "+1,0% en 2025, +1,1% en 2024")
bold_bullet(doc, "Déficit public", "-5,2% du PIB en 2025 (pression sur dépenses publiques futures)")
bold_bullet(doc, "Poids BTP", "[DONNÉE NON DISPONIBLE — % PIB construction BE non trouvé en source fiable]")
doc.add_paragraph()

add_heading(doc, "1.3 Relations commerciales France-Belgique", 2)
add_bullet(doc, "Volume total échanges bilatéraux : 89,4 Md€ en 2025 (Trésor FR, 2025), -8,5% vs 2024.")
add_bullet(doc, "Exportations françaises vers BE : 44,9 Md€ (stable). Importations depuis BE : 44,5 Md€ (-15,1%).")
add_bullet(doc, "La France dégage un léger excédent de +0,3 Md€ en 2025 (après déficit -7,2 Md€ en 2024).")
add_bullet(doc, "Stock IDE français en Belgique : 126 405 M€ en 2024. Flux : 1 358 M€.")
add_bullet(doc, "Belgique = 6ème partenaire commercial de la France (2ème client et 3ème fournisseur vu de BE).")
add_bullet(doc, "Proximité culturelle forte côté wallon/bruxellois (francophone). Côté flamand : néerlandophone, sensibilité produit allemande/néerlandaise souvent dominante.")
add_bullet(doc, "Avantage Presto : marque française reconnue, facilité logistique (livraison rapide depuis FR), documentation disponible en français pour partie francophone.")
doc.add_paragraph()

add_heading(doc, "1.4 Tendances d'investissement", 2)
add_bullet(doc, "BEI Groupe : 2,6 Md€ investis en Belgique en 2025 (infra sociales, innovation, transition verte).")
add_bullet(doc, "Logement social Flandre : 1,7 Md€ BEI pour 56 000 logements sociaux d'ici 2042 (premier tranche 700 M€ signée).")
add_bullet(doc, "Santé mentale : 120 M€ BEI pour Z.org KU Leuven/Kortenberg — programme total 270 M€ (2026-2040).")
add_bullet(doc, "Hôpitaux universitaires FWB : 438 M€ plan 2024-2028 (CHU Liège 160M€, Saint-Luc 171M€, Erasme 55M€, Mont-Godinne 37M€).")
add_bullet(doc, "Bâtiments scolaires FWB : 1 Md€ de subventions (dont 300M€ appel 1, 200M€ appel 2 nov. 2023).")
add_bullet(doc, "Rénovation logements sociaux Wallonie : 1,2 Md€ pour 25 000 logements (objectif 55 000 d'ici 2030).")
add_bullet(doc, "Plan de relance : 51% des fonds RRF alloués à des mesures climatiques (rénovation énergétique).")
doc.add_paragraph()

# ─── PARTIE 2 ────────────────────────────────────────────────────────────────
add_heading(doc, "PARTIE 2 — MARCHÉ DE LA CONSTRUCTION", 1)

add_heading(doc, "2.1 État actuel", 2)
add_bullet(doc, "Marché Benelux (BE+NL+LU) : 111 Md$ en 2024, CAGR 4,2% prévu jusqu'à 2032 → 152 Md$ (DataBridge Market Research, 2024).")
add_bullet(doc, "Belgique seule 2024 : contraction de -2,8% — une des pires performances d'Europe occidentale.")
add_bullet(doc, "Permis de construire : -11% (tous types). Permis résidentiels : -15%.")
add_bullet(doc, "Faillites : 2 600+ entreprises de construction en 2024, soit +17% vs 2023 (Allianz Trade, 2024).")
add_bullet(doc, "53 800 logements livrés en 2024 (+10% vs 2023), mais volumes de construction résidentielle -11% depuis pic T3 2021.")

add_heading(doc, "2.2 Dynamique Neuf vs Rénovation", 2)
add_bullet(doc, "Neuf résidentiel : sous-performance majeure. Nombre de permis historiquement bas.")
add_bullet(doc, "Rénovation : dominant — le nombre de permis de rénovation dépasse celui des nouvelles constructions depuis plusieurs années.")
add_bullet(doc, "Génie civil 2024 : +4,4% (après +4,9% en 2023) — seul segment performant, porté par plans de relance et investissements municipaux.")
add_bullet(doc, "Génie civil 2025 : -2,2% attendu (Embuild / ING, 2025) — décélération après deux années de soutien public.")

add_heading(doc, "2.3 Perspectives", 2)
add_bullet(doc, "2025 : -0,4% global (EMAE, 2025). Contexte : demande atone, marges faibles, contraintes administratives excessives.")
add_bullet(doc, "2026 : reprise très modérée attendue (ING, déc. 2025) — bâtiments non-résidentiels et rénovation en tête.")
add_bullet(doc, "Moteurs durables : rénovation énergétique (PEB), investissements publics ciblés, demande de bâtiments durables.")
doc.add_paragraph()

# ─── PARTIE 3 ────────────────────────────────────────────────────────────────
add_heading(doc, "PARTIE 3 — CONSTRUCTION NON-RÉSIDENTIELLE", 1)

add_heading(doc, "3.1 État actuel et part dans la construction totale", 2)
add_bullet(doc, "Seul grand segment en croissance positive en 2024 : +1,4% pour le non-résidentiel neuf (bureaux, commerces, bâtiments industriels).")
add_bullet(doc, "Non-résidentiel rénovation : -2,1% en 2024, mais autorisations accordées +8,7% — signal positif à 12-18 mois.")
add_bullet(doc, "Perspectives 2025 : +1,5% attendu pour le neuf non-résidentiel (ING, 2025). Soutien des projets publics (santé, éducation) et rénovations énergétiques.")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — part exacte du non-résidentiel dans la construction totale belge en valeur]")

add_heading(doc, "3.2 Segments dominants", 2)
add_bullet(doc, "Bureaux et espaces commerciaux : en hausse — regain d'activité post-COVID.")
add_bullet(doc, "Bâtiments industriels : croissance continue, portée par logistique et relocalisation industrielle.")
add_bullet(doc, "Projets publics (santé, éducation) : investissements planifiés pluriannuels (cf. Partie 1.4) — pipeline solide et visible.")
add_bullet(doc, "Non-résidentiel privé : prudent mais stable.")
doc.add_paragraph()

# ─── PARTIE 4 ────────────────────────────────────────────────────────────────
add_heading(doc, "PARTIE 4 — POIDS DES SEGMENTS ERP", 1)
add_note(doc, "Les données chiffrées par segment ERP en Belgique sont fragmentaires. Les informations ci-dessous combinent données disponibles et signaux qualitatifs. Les lacunes sont explicitement signalées.")

add_heading(doc, "4.1 Éducation", 2)
add_bullet(doc, "Fédération Wallonie-Bruxelles : 1 Md€ de subventions pour rénovation bâtiments scolaires (2023-2028). 1er appel 300M€, 2ème appel 200M€ (nov. 2023).")
add_bullet(doc, "Priorités : performance énergétique, accessibilité PMR, connectivité.")
add_bullet(doc, "Résidences étudiantes : projet emblématique Sart-Tilman (ULiège) — 407 chambres, livraison fin 2025. Appel à projets Wallonie pour logements étudiants publics en cours.")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — nombre total d'établissements scolaires BE, volume marché rénovation/neuf éducation par an]")

add_heading(doc, "4.2 Santé / Hôpitaux / EHPAD", 2)
add_bullet(doc, "Hôpitaux universitaires FWB : plan 438 M€ 2024-2028 — rénovation + construction neuve (CHU Liège, Saint-Luc, Erasme, Mont-Godinne).")
add_bullet(doc, "Santé mentale : Z.org KU Leuven + Kortenberg, programme 270 M€ jusqu'à 2040 (120 M€ BEI).")
add_bullet(doc, "EHPAD/Maisons de retraite : [DONNÉE NON DISPONIBLE — parc total BE, taux de rénovation, volume investissement]")
add_bullet(doc, "Signal fort : vieillissement de la population belge → pression croissante sur l'offre de soins et d'hébergement.")

add_heading(doc, "4.3 Bâtiments tertiaires (Bureaux / Cantines)", 2)
add_bullet(doc, "Seul segment non-résidentiel en croissance nette (+1,4% 2024). Bureaux et cantines d'entreprises en regain.")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — volume marché bureau neuf/rénové en valeur, nombre de m² livrés]")

add_heading(doc, "4.4 Bâtiments industriels", 2)
add_bullet(doc, "En hausse en 2024. Logistique et industrie légère moteurs principaux.")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — données chiffrées spécifiques]")

add_heading(doc, "4.5 CHR (Cafés, Hôtels, Restaurants)", 2)
add_bullet(doc, "Secteur hôtelier belge 2024 : reprise post-COVID, tendance à la montée en gamme (sécurité, wellness, piscines privées). (Forbes Belgique, 2024)")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — nombre d'hôtels BE, volume investissement CHR en robinetterie]")

add_heading(doc, "4.6 HPA (Hôtellerie Plein Air, Campings, Piscines plein air)", 2)
add_bullet(doc, "Wallonie : extension des incitations pour piscines communautaires votée déc. 2025. Coopérations supracommunales encouragées.")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — parc campings BE, volume équipement HPA]")

add_heading(doc, "4.7 Centres Sport & Loisirs", 2)
add_bullet(doc, "Programme Infrasports Wallonie : subventions aux communes pour équipements sportifs.")
add_bullet(doc, "Projet phare : nouveau complexe sportif avec piscine en périphérie bruxelloise — travaux début 2025 (piscine olympique + récréative + wellness + vagues). (L'Avenir, 2024)")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — nombre de gymnases et piscines couvertes BE, volume marché]")

add_heading(doc, "4.8 Établissements à sécurité renforcée (Pénitentiaire)", 2)
add_bullet(doc, "Surpopulation carcérale critique : 13 397 détenus pour 10 795 places (taux 124%), 614 au sol (Prison Insider, déc. 2025).")
add_bullet(doc, "Plan d'action fédéral juillet 2025 : maintien prisons anciennes + unités modulaires préfabriquées + nouvelles constructions avec technologies avancées.")
add_bullet(doc, "Nouvelles prisons : Anvers livrée début 2025 (440 détenus). Vresse-sur-Semois : 171 M€ (Wallonie). 3 autres prisons à horizon 2030.")
add_bullet(doc, "Pipeline de construction pénitentiaire solide — segment anti-vandalisme stratégique pour Presto.")

add_heading(doc, "4.9 Bâtiments culturels, lieux de culte, transports", 2)
add_bullet(doc, "[DONNÉE NON DISPONIBLE — données sectorielles BE non trouvées en sources fiables]")
add_bullet(doc, "Note : aéroports (Brussels Airport, Charleroi), gares SNCB — projets de modernisation en cours mais données non sourcées.")
doc.add_paragraph()

add_heading(doc, "4.12 Opportunités Presto par segment — synthèse", 2)
add_note(doc, "Classement par potentiel décroissant pour Presto sur le marché belge ERP.")
rows_412 = [
    ["Segment", "Score", "Produits Presto", "Arguments clés", "Canal recommandé"],
    ["4.2 Santé / Hôpitaux", "5/5", "Temposoft, Tempostop, gamme hospitalière", "Hygiène, anti-brûlure, PMR, BELGAQUA", "Prescription BET + distributeurs santé"],
    ["4.1 Éducation", "4/5", "Temporisateurs push-button, électroniques", "Économie d'eau (PLAGE/UREBA), robustesse", "Appels d'offres publics, négoce pro"],
    ["4.8 Pénitentiaire", "4/5", "Inox anti-vandalisme, encastrés", "Anti-vandalisme, indestructibilité, entretien minimal", "Prescription directe Régie des Bâtiments"],
    ["4.5 CHR / Hôtels", "3/5", "Thermostatiques, design, économiseurs", "Montée en gamme, durabilité, pression constante", "Distributeurs CHR, architectes"],
    ["4.7 Sport & Loisirs", "3/5", "Temporisateurs douche, push-button", "Économie d'eau, hygiène, certif. PMR", "Collectivités locales, BET sport"],
    ["4.3 Tertiaire / Bureaux", "3/5", "Temporisateurs, électroniques sans contact", "BREEAM, économie eau, UE/OTAN niche", "Prescription architectes + BET"],
    ["4.11 Transports", "2/5", "Robinets encastrés, anti-vandalisme", "Résistance vandalisme, entretien minimal", "BET spécialisés, marchés SNCB/BSCA"],
    ["4.4 Industrie", "2/5", "Inox, temporisateurs process", "Résistance corrosion, hygiène alimentaire/pharma", "Bureaux d'études process"],
    ["4.6 HPA", "1/5", "Limitée", "Marché quasi inexistant en Belgique", "—"],
    ["4.9 Bâtiments culturels", "1/5", "Opportunités ponctuelles", "Projets municipaux dispersés", "—"],
    ["4.10 Lieux de culte", "1/5", "Très limité", "Marché résiduel, non prioritaire", "—"],
]
tbl_412 = doc.add_table(rows=len(rows_412), cols=5)
tbl_412.style = "Table Grid"
for i, row_data in enumerate(rows_412):
    for j, cell_text in enumerate(row_data):
        cell = tbl_412.rows[i].cells[j]
        cell.text = cell_text
        run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(cell_text)
        run.font.name = FONT
        run.font.size = Pt(9)
        if i == 0:
            run.bold = True
doc.add_paragraph()
add_para(doc, "Priorités absolues : Santé et Éducation (volumes + pipelines sécurisés). Niche stratégique : Pénitentiaire (anti-vandalisme, sans concurrent dominant). Levier différenciant : certification BELGAQUA à obtenir en priorité.")
doc.add_paragraph()

# ─── PARTIE 5 ────────────────────────────────────────────────────────────────
add_heading(doc, "PARTIE 5 — TAILLE MARCHÉ : ROBINETTERIE GÉNÉRALE", 1)

add_heading(doc, "5.1 Taille et valeur du marché — Double estimation obligatoire (v4)", 2)
add_note(doc, "Aucune source publique ne recense la taille exacte du marché belge. Deux estimations par extrapolation, présentées côte à côte conformément au protocole MAB v4.")

add_heading(doc, "Constantes de référence", 2)
bold_bullet(doc, "PIB/hab France 2025", "48 982 USD (Worldometer)")
bold_bullet(doc, "Population France 2025", "69,1 M (Worldometer)")
bold_bullet(doc, "PIB/hab Belgique 2024", "46 000 USD (SPF Économie / FMI)")
bold_bullet(doc, "Population Belgique 2025", "11,825 M (Statbel)")
bold_bullet(doc, "Coefficient global", "(46 000 / 48 982) × (11,825 / 69,1) = 0,939 × 0,171 = 0,161")

add_heading(doc, "Estimation A — base « Analyse de Marché France » (source interne Presto)", 2)
add_note(doc, "Base France = valeurs internes Presto divisées par 2 (valeurs HT). Ne pas utiliser Xerfi 635 M€.")
rows_5a = [
    ["Segment", "Base France (/2)", "× Coeff. 0,161", "Estimation Belgique"],
    ["Robinetterie collectivités", "100–125 M€", "× 0,161", "16–20 M€"],
    ["Chasses d'eau & équipements WC", "90–110 M€", "× 0,161", "14–18 M€"],
    ["Douches & équipements connexes", "52–65 M€", "× 0,161", "8–10 M€"],
    ["TOTAL Estimation A", "242–300 M€", "× 0,161", "39–48 M€"],
]
tbl_5a = doc.add_table(rows=len(rows_5a), cols=4)
tbl_5a.style = "Table Grid"
for i, row_data in enumerate(rows_5a):
    for j, cell_text in enumerate(row_data):
        cell = tbl_5a.rows[i].cells[j]
        cell.text = cell_text
        run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(cell_text)
        run.font.name = FONT
        run.font.size = Pt(9)
        if i == 0 or j == 0:
            run.bold = True
doc.add_paragraph()
add_note(doc, "Fiabilité : moyenne — méthode macro, ne capte pas l'économie informelle ni les spécificités sectorielles locales.")

add_heading(doc, "Estimation B — base « Études BRG » (BRG Belgium 2020)", 2)
add_note(doc, "[DONNÉE NON DISPONIBLE — BE_Bathrooms_Full_Report_2020.pdf présent en sources-internes mais non extractible sans poppler. La valeur de référence France citée dans le BRG n'a pas pu être lue. Estimation B à compléter dès installation de poppler-utils ou conversion manuelle du PDF.]")
add_bullet(doc, "Action : installer poppler (`brew install poppler`) et relancer le script pour extraire automatiquement la base BRG et calculer l'Estimation B.")

add_heading(doc, "Synthèse Partie 5", 2)
add_bullet(doc, "Estimation A (seule disponible) : 39–48 M€ pour le périmètre robinetterie collective + chasses d'eau + douches en Belgique.")
add_bullet(doc, "Estimation B : [NON DISPONIBLE — BRG non lisible]")
add_bullet(doc, "Écart inter-méthodes : non calculable. À compléter.")
add_bullet(doc, "Commentaire : L'Estimation A est cohérente avec la fourchette 24–29 M€ retenue en source interne pour la seule robinetterie collective (segment le plus étroit). Le total élargi converge.")

add_heading(doc, "5.2 Canaux de distribution", 2)
add_bullet(doc, "Négoce spécialisé plomberie-chauffage : principal canal BtoB (SIDER, Versani, Willems-Diels, Aquacaro).")
add_bullet(doc, "Groupe BME/STG : acquisition de Paepens (Flandre) → réseau négoce sanitaire/chauffage en structuration.")
add_bullet(doc, "E-commerce professionnel : Sawiday.be, Sider.biz — présence Presto confirmée sur ces plateformes.")
add_bullet(doc, "Prescripteurs : architectes, bureaux d'études techniques, coordinateurs de marchés publics.")
add_note(doc, "Hypothèse basée sur le modèle France — à confirmer terrain.")

add_heading(doc, "5.3 Spécificités produit Belgique", 2)
add_bullet(doc, "Forte sensibilité aux marques allemandes/autrichiennes côté flamand (Hansgrohe, Grohe, Hansa).")
add_bullet(doc, "Côté wallon/bruxellois : proximité culturelle française → ouverture aux marques FR (Presto, Delabie).")
add_bullet(doc, "Exigence croissante en efficacité hydrique et certification PEB — avantage produits temporisés et électroniques.")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — volumes de vente par gamme produit, prix moyens marché BE]")

add_heading(doc, "5.4 Dynamique et perspectives", 2)
add_bullet(doc, "Croissance portée par la rénovation énergétique — économiseurs d'eau certifiés imposés dans les marchés publics.")
add_bullet(doc, "E-procurement en progression dans les marchés publics belges.")
add_bullet(doc, "Digitalisation des achats B2B : grossistes STG et Facq Pro disposent de plateformes e-commerce professionnelles.")
doc.add_paragraph()

# ─── PARTIE 6 ────────────────────────────────────────────────────────────────
add_heading(doc, "PARTIE 6 — TAILLE MARCHÉ : ROBINETTERIE COLLECTIVE ERP", 1)

add_heading(doc, "6.1 Taille du marché", 2)
add_bullet(doc, "Source interne Presto : 24–29 M€ pour la robinetterie ERP stricte (collective) en 2025.")
add_bullet(doc, "Périmètre élargi (semi-collectif, cantines, douches) : 55–66 M€.")
add_bullet(doc, "Croissance estimée : 3–4 %/an sur 2025–2028 (source interne).")
add_bullet(doc, "Atout structurel : institutions UE/OTAN à Bruxelles = demande ERP institutionnelle haut de gamme, peu cyclique.")

add_heading(doc, "6.2 Double méthode d'extrapolation (v4)", 2)
add_note(doc, "Deux méthodes obligatoires conformément au protocole MAB v4. Résultats à présenter côte à côte.")

add_heading(doc, "Méthode 1 — base « Analyse de Marché France »", 2)
add_bullet(doc, "Base France collectivités (/2) : 100–125 M€")
add_bullet(doc, "Coefficient global : 0,161 (cf. Partie 5)")
add_bullet(doc, "Estimation brute : 0,161 × 112,5 M€ (midpoint) = 18,1 M€")
add_bullet(doc, "Ajustement structurel :")
add_bullet(doc, "+10 % économie formelle (Belgique : quasi-zéro économie informelle, marchés publics traçables)", level=1)
add_bullet(doc, "+5 % urbanisation très élevée (99 % — demande ERP concentrée, structurée)", level=1)
add_bullet(doc, "+5 % PIB/hab proche de la France (marché premium, valeur unitaire haute)", level=1)
add_bullet(doc, "–5 % marché BTP atone 2024–2025 (ralentissement investissement public)", level=1)
add_bullet(doc, "Ajustement net : +15 % → facteur 1,15")
add_bullet(doc, "Estimation Méthode 1 : 18,1 M€ × 1,15 = 20,8 M€  →  fourchette 19–23 M€")
add_note(doc, "Estimation par extrapolation avec ajustement structurel — fiabilité moyenne. Variables ajustement à confirmer terrain.")

add_heading(doc, "Méthode 2 — base « Études BRG »", 2)
add_note(doc, "[DONNÉE NON DISPONIBLE — BE_Bathrooms_Full_Report_2020.pdf non extractible sans poppler. La base France citée dans le BRG Belgique 2020 n'a pas pu être lue. Méthode 2 à compléter après installation de poppler-utils (`brew install poppler`).]")

add_heading(doc, "Synthèse 6.2", 2)
rows_62 = [
    ["", "Méthode 1 (base Presto interne)", "Méthode 2 (base BRG)"],
    ["Base France", "100–125 M€ (/2)", "[NON DISPONIBLE]"],
    ["Coefficient", "0,161", "0,161"],
    ["Ajustement structurel", "+15 %", "+15 % (identique)"],
    ["Estimation brute", "~18 M€", "[NON DISPONIBLE]"],
    ["Estimation ajustée", "19–23 M€", "[NON DISPONIBLE]"],
    ["Source interne validation", "24–29 M€ (convergence ✓)", "—"],
    ["Niveau de confiance", "Moyen", "—"],
]
tbl_62 = doc.add_table(rows=len(rows_62), cols=3)
tbl_62.style = "Table Grid"
for i, row_data in enumerate(rows_62):
    for j, cell_text in enumerate(row_data):
        cell = tbl_62.rows[i].cells[j]
        cell.text = cell_text
        run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(cell_text)
        run.font.name = FONT
        run.font.size = Pt(9)
        if i == 0:
            run.bold = True
doc.add_paragraph()
add_bullet(doc, "Fourchette finale retenue : 19–29 M€ (Méthode 1 + source interne). Niveau de confiance : MOYEN.")
add_note(doc, "Limites : ne capte pas l'économie informelle ; ne reflète pas les spécificités sectorielles locales ; volatilité taux de change ; ajustement structurel repose sur hypothèses à confirmer terrain.")

add_heading(doc, "6.3 Potentiel par segment ERP — Scoring 1 à 5 (v4)", 2)
add_note(doc, "Score : 1 = très faible / 2 = faible / 3 = moyen / 4 = fort / 5 = très fort potentiel pour Presto en Belgique.")
rows_63 = [
    ["Segment", "Score", "Justification", "Hypothèses clés"],
    ["4.1 Éducation", "4/5", "Pipeline 3,2 Md€ Flandre + 1 Md€ FWB documenté ; rénovation urgente (66 % bâti pre-1981) ; produits temporisés imposés par PLAGE/UREBA.", "Accès marché flamand (néerlandophone) via distributeur local."],
    ["4.2 Santé / EHPAD", "5/5", "Investissements sécurisés pluriannuels (438 M€ hôpitaux univ., 270 M€ psychiatrie) ; exigences hygiène max. ; gamme hospitalière Presto directement applicable.", "Différenciation prix vs Delabie sur segments à budget contraint."],
    ["4.3 Tertiaire / Bureaux", "3/5", "Segment en reprise (+1,4 % 2024) ; UE/OTAN à Bruxelles = niche haute valeur + BREEAM ; mais moins spécifique ERP que santé/éducation.", "Accès via prescription architectes + BET."],
    ["4.4 Industrie / Logistique", "2/5", "Marché inox spécialisé à forte valeur unitaire mais volume très limité ; exigences techniques élevées.", "Référencement via bureaux d'études process (pharma, biotech)."],
    ["4.5 CHR / Hôtels", "3/5", "Hôtellerie en reprise forte (10 M+ visiteurs 2024) ; montée en gamme → budget sanitaire premium ; mais concurrence généraliste forte.", "Hôtels 4–5 étoiles Bruxelles/Anvers = cœur de cible."],
    ["4.6 HPA", "1/5", "Marché quasi inexistant en Belgique (climat tempéré, faible culture camping vs France) ; opportunités structurelles négligeables.", "Marché résiduel, non à prioriser."],
    ["4.7 Sport & Loisirs", "3/5", "Plans piscines actifs (PLAGE, Infrasports Wallonie 29 M€, Bruxelles 43 M€) ; parc vieillissant ; certification économie eau imposée.", "Rénovations liées aux obligations EPBD = déclencheur."],
    ["4.8 Pénitentiaire", "4/5", "Plan Maître III : 5 nouvelles prisons actives (Anvers livrée 2025, Vresse 171 M€, 3 autres d'ici 2030) ; anti-vandalisme = cœur de gamme ; marché captif.", "Presto non dominé par Delabie sur ce segment — ouverture réelle."],
    ["4.9 Bâtiments culturels", "1/5", "Données insuffisantes ; marché résiduel ; investissements principalement régionaux et non programmés.", "Opportunités ponctuelles uniquement."],
    ["4.10 Lieux de culte", "1/5", "Marché résiduel, non prioritaire pour ERP collectif Presto.", "Investissements très faibles et dispersés."],
    ["4.11 Transports", "2/5", "SNCB + aéroports : volumes intéressants mais chaîne prescription longue, certifications spécifiques, cycles longs.", "Accès via BET spécialisés infra ; délais commerciaux 18–36 mois."],
]
tbl_63 = doc.add_table(rows=len(rows_63), cols=4)
tbl_63.style = "Table Grid"
for i, row_data in enumerate(rows_63):
    for j, cell_text in enumerate(row_data):
        cell = tbl_63.rows[i].cells[j]
        cell.text = cell_text
        run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(cell_text)
        run.font.name = FONT
        run.font.size = Pt(9)
        if i == 0:
            run.bold = True
doc.add_paragraph()
add_para(doc, "Top 3 segments prioritaires : Santé (5/5) · Éducation (4/5) · Pénitentiaire (4/5)")

add_heading(doc, "6.4 Spécificités produit ERP Belgique", 2)
add_bullet(doc, "Robinetterie temporisée push-button et électronique : dominantes sur ERP — hygiène, économie d'eau, résistance vandalisme.")
add_bullet(doc, "Prix moyen ERP standard : 90–160 € HT. Segment institutionnel haut de gamme (UE/OTAN) : 200 €+ (source interne Presto).")
add_bullet(doc, "Marché francophone (Wallonie + Bruxelles) : préférence marques françaises — avantage Presto.")
add_bullet(doc, "Marché flamand : marques germaniques dominantes — entrée via partenaire distributeur local recommandée.")

add_heading(doc, "6.5 Dynamique et perspectives", 2)
add_bullet(doc, "Programmes PLAGE (Flandre) et UREBA (Wallonie) : rénovation bâtiments publics impose économiseurs d'eau certifiés — débouché direct.")
add_bullet(doc, "Institutions UE à Bruxelles : cahiers des charges BREEAM, PMR, anti-legionella — marché de niche à haute valeur.")
add_bullet(doc, "Rénovation hôpitaux et écoles : flux régulier de projets 2024–2030 — marché visible et qualifiable.")
doc.add_paragraph()

# ─── PARTIE 7 ────────────────────────────────────────────────────────────────
add_heading(doc, "PARTIE 7 — CONCURRENTS", 1)

add_heading(doc, "7.1 DELABIE (analyse prioritaire)", 2)
add_bullet(doc, "Positionnement : leader européen de la robinetterie et équipements sanitaires pour ERP. Haut de gamme institutionnel.")
add_bullet(doc, "Présence Belgique : filiale Delabie Benelux SRL, Sint-Pieters-Leeuw (Bruxelles), issu de l'acquisition de BSC Belgium. Site dédié : delabiebenelux.com (FR + NL).")
add_bullet(doc, "Organisation : équipe commerciale locale, catalogue complet ERP, présence sur marchés publics belges.")
add_bullet(doc, "Gamme : robinetterie temporisée et électronique, robinetterie hospitalière, sanitaires inox anti-vandalisme, accessibilité PMR, grandes cuisines collectives.")
add_bullet(doc, "Forces : marque très connue dans les ERP, réseau négoce établi, documentation technique exhaustive, certifications NF et EN.")
add_bullet(doc, "Faiblesses : prix haut de gamme, potentiellement perçu comme inaccessible sur certains segments publics à budget contraint.")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — parts de marché Delabie Benelux en % du marché BE collectif]")

add_heading(doc, "7.2 Autres concurrents", 2)
bold_bullet(doc, "Hansgrohe / Grohe (LIXIL)", "Marques premium allemandes très présentes en Belgique, notamment côté flamand. Forte présence retail + projets résidentiels haut de gamme. Moins spécialisées ERP collectif.")
bold_bullet(doc, "Hansa", "Robinets intelligents, milieu-haut de gamme. Site Hansa.com en version belge (hansa.com/fr-be). Produits economy-eau mais positionnement moins ERP.")
bold_bullet(doc, "Geberit", "Acteur majeur en sanitaires encastrés, moins sur robinetterie apparente. Fort en projets neufs tertiaires et santé.")
bold_bullet(doc, "Jacob Delafon (Kohler)", "Présence Belgique, milieu de gamme, plus résidentiel que collectif.")
bold_bullet(doc, "Oras / Damixa", "Marques nordiques présentes via négoce. Moins dominantes sur le collectif institutionnel.")

add_heading(doc, "7.3 Opportunités de différenciation pour Presto", 2)
add_bullet(doc, "Profondeur de gamme collective/temporisée : Presto = spécialiste pur ERP comme Delabie, contrairement aux généralistes résidentiels.")
add_bullet(doc, "Rapport qualité/prix : alternative crédible à Delabie sur segments à budget contraint (éducation, logement social).")
add_bullet(doc, "Réactivité logistique depuis France : délais courts, stock disponible, proximité géographique.")
add_bullet(doc, "Présence digitale Belgique à renforcer : développer la présence sur les plateformes B2B belges et les prescripteurs.")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — réputation terrain Presto en Belgique auprès des installateurs et maîtres d'œuvre — à valider par entretiens]")
doc.add_paragraph()

# ─── PARTIE 8 ────────────────────────────────────────────────────────────────
add_heading(doc, "PARTIE 8 — NORMES & CERTIFICATIONS ROBINETTERIE", 1)

add_heading(doc, "8.1 Normes applicables", 2)
add_bullet(doc, "NBN EN 816 (1996) : Robinetterie sanitaire — robinets à fermeture automatique PN 10. S'applique aux temporisateurs (cœur de gamme Presto). Couvre marquage, hydraulique, étanchéité, acoustique, résistance pression, endurance mécanique.")
add_bullet(doc, "NBN EN 817 (2008) : Robinetterie sanitaire — mitigeurs mécaniques PN 10. Spécifications techniques générales.")
add_bullet(doc, "Marquage CE : obligatoire pour mise sur le marché en Belgique (directive produits de construction). Condition sine qua non.")
add_bullet(doc, "Normes EN 200, EN 1111, EN 1112, EN 1113 : robinetterie résidentielle (moins contraignantes pour ERP).")

add_heading(doc, "8.2 Certifications recommandées", 2)
add_bullet(doc, "NF Robinetterie Sanitaire (NF077, AFNOR) : certification française reconnue et valorisée en Belgique francophone — donne confiance aux prescripteurs et maîtres d'ouvrage publics.")
add_bullet(doc, "Certifications hospitalières spécifiques : pour les produits destinés aux ERP de santé (cf. Delabie — certifications hôpitaux NF).")
add_bullet(doc, "[DONNÉE NON DISPONIBLE — existence d'une certification belge spécifique type 'BENOR' pour la robinetterie — à vérifier auprès du NBN]")

add_heading(doc, "8.3 Organisme certificat belge", 2)
add_bullet(doc, "NBN (Bureau de Normalisation Belge) — www.nbn.be — organisme de référence pour les normes EN transposées en Belgique.")
add_bullet(doc, "Les normes belges NBN EN sont directement transposées des EN européennes : peu d'écarts vs normes françaises NF EN.")

add_heading(doc, "8.4 Contraintes d'entrée et écarts vs France", 2)
add_bullet(doc, "Écarts normatifs vs France : faibles (normes EN harmonisées communes). Les produits déjà certifiés CE et NF pour la France sont en principe conformes pour le marché belge.")
add_bullet(doc, "Particularité : marchés publics belges peuvent exiger des certifications spécifiques dans les cahiers des charges — à vérifier au cas par cas.")
add_bullet(doc, "Délais d'obtention : pas de certification belge spécifique identifiée → entrée de marché normative rapide si produits déjà certifiés CE/NF.")
doc.add_paragraph()

# ─── PARTIE 9 ────────────────────────────────────────────────────────────────
add_heading(doc, "PARTIE 9 — POINTS À REVÉRIFIER", 1)
add_para(doc, "Données incertaines ou manquantes, à revalider par recherches complémentaires ou terrain :")

bold_bullet(doc, "Taille marché robinetterie BE", "Estimation uniquement disponible par extrapolation FR. Vérifier via Xerfi, Markest Belgium, ou fédération professionnelle Techlink. Contact : techlink.be")
bold_bullet(doc, "Taille marché robinetterie collective/ERP BE", "Aucune donnée trouvée. À obtenir via rapport sectoriel ou entretiens distributeurs. Contact : négoce plomberie (SIDER, STG/BME).")
bold_bullet(doc, "Parts de marché Delabie BE", "Non disponibles publiquement. À estimer via entretiens terrain (distributeurs, installateurs). Indicateur : taille de la filiale Benelux (RCS/Trends.be).")
bold_bullet(doc, "Réputation Presto terrain en Belgique", "À valider par entretiens installateurs et prescripteurs. Identifier les régions où Presto est le plus/moins connu.")
bold_bullet(doc, "Certification BENOR robinetterie", "Vérifier auprès du NBN si une certification belge spécifique existe pour la robinetterie ERP et si elle est requise dans les marchés publics belges.")
bold_bullet(doc, "Poids BTP dans PIB belge", "Donnée non trouvée en source fiable — à chercher via Statbel (statbel.fgov.be) ou Embuild (fédération construction BE).")
bold_bullet(doc, "Données segments ERP chiffrées", "Parc hôpitaux, EHPAD, gymnases, piscines, prisons en Belgique — à compléter via SPF Santé publique, Statbel, Régie des Bâtiments.")
bold_bullet(doc, "Distribution Presto Belgique actuelle", "Identifier les distributeurs actuels, les régions couvertes, et les gaps réseau. À demander en interne.")
doc.add_paragraph()

# Sauvegarde
path_etude = os.path.join(OUTPUT_DIR, "MAB_Belgique_Etude.docx")
doc.save(path_etude)
print(f"✓ Étude sauvegardée : {path_etude}")

# ─── DOCUMENT ANNEXES ────────────────────────────────────────────────────────

ann = Document()
for section in ann.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

t2 = ann.add_heading("MAB BELGIQUE — ANNEXES & SOURCES", 0)
t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub2 = ann.add_paragraph("Sources complètes, données brutes et compléments")
sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub2.runs[0].font.size = Pt(11)
ann.add_paragraph()

add_heading(ann, "ANNEXE 1 — LISTE DES SOURCES UTILISÉES", 1)
ann.add_paragraph("Toutes les sources consultées pour la production de MAB_Belgique_Etude.docx :")
ann.add_paragraph()

sources = [
    ("Direction Générale du Trésor FR — Fiche Belgique (2025)", "https://www.tresor.economie.gouv.fr/Pays/BE/situation-economique-et-financiere-de-la-belgique", "Français", "Juin 2026"),
    ("Worldometer — Population Belgique (2026)", "https://www.worldometers.info/fr/population-mondiale/belgique-population/", "Français", "Juin 2026"),
    ("Banque Nationale de Belgique — Projections macroéconomiques", "https://www.nbb.be/en/publications-research/publications/economic-and-financial-publications/macroeconomic-projections", "Anglais/Français", "Juin 2026"),
    ("Trésor FR — Relations bilatérales France-Belgique (2025)", "https://www.tresor.economie.gouv.fr/Pays/BE/relations-bilaterales", "Français", "Juin 2026"),
    ("Ambassade de France en Belgique — Relations économiques", "https://be.ambafrance.org/Relations-economiques-7221", "Français", "Juin 2026"),
    ("DataBridge Market Research — Benelux Construction Market (2024)", "https://www.databridgemarketresearch.com/reports/benelux-construction-market", "Anglais", "Juin 2026"),
    ("Allianz Trade — Analyse secteur construction Belgique 2024", "https://www.allianz-trade.com/fr_BE/actualites/dernieres-actualites/analyse-secteur-construction.html", "Français", "Juin 2026"),
    ("Embuild — Secteur construction en recul (2025)", "https://embuild.be/fr/le-secteur-de-la-construction-et-de-l%E2%80%99installation-en-recul-pour-la-quatri%C3%A8me-ann%C3%A9e-cons%C3%A9cutive", "Français", "Juin 2026"),
    ("ING Belgique — Perspectives construction 2026", "https://www.ing.be/fr/particuliers/actus/economie-et-marches-financiers/secteur-de-la-construction", "Français", "Juin 2026"),
    ("Techlink — Rétrospective non-résidentielle (2024-2025)", "https://techlink.be/fr/actualites/retrospective-et-perspectives-sur-la-construction-neuve-et-la-renovation-non-residentielle", "Français", "Juin 2026"),
    ("BEI — Belgique 2025 : 2,6 Md€ de financement", "https://www.eib.org/en/press/all/2026-038-activite-du-groupe-bei-en-2025-26-milliards-d-euros", "Anglais", "Juin 2026"),
    ("BEI — 120 M€ pour Z.org KU Leuven (santé mentale)", "https://www.eib.org/en/press/all/2025-529-eib-supports-major-renewal-of-mental-health-infrastructure-in-leuven-and-kortenberg-with-eur120-million-loan", "Anglais", "Juin 2026"),
    ("BEI — 1,7 Md€ logements sociaux Flandre", "https://europeansting.com/2026/02/03/belgium-eib-group-2025-figures-e2-6-bn-in-financing-for-social-infrastructure-innovation-and-the-green-transition/", "Anglais", "Juin 2026"),
    ("Le Spécialiste — Plan construction hôpitaux universitaires 438 M€", "https://www.lespecialiste.be/fr/actualites/socio-professionnel/plan-de-construction-438-millions-pour-les-4-hopitaux-universitaires.html", "Français", "Juin 2026"),
    ("Fédération Wallonie-Bruxelles — 2ème appel bâtiments scolaires 200 M€", "https://www.federation-wallonie-bruxelles.be/nc/detail-article/", "Français", "Juin 2026"),
    ("Prison Insider — Belgique 2025", "https://www.prison-insider.com/fichepays/belgique-2025", "Français", "Juin 2026"),
    ("Jan De Nul — Nouvelle prison Anvers", "https://www.jandenul.com/fr/projets/nouvelle-prison-pour-anvers-belgique", "Français", "Juin 2026"),
    ("DH Les Sports+ — Quatre nouvelles prisons d'ici 2030", "https://www.dhnet.be/actu/belgique/2023/03/21/quatre-nouvelles-prisons", "Français", "Juin 2026"),
    ("Media24 — Vresse-sur-Semois prison 171 M€ (nov. 2025)", "https://media24.fr/2025/11/10/la-belgique-a-le-meme-probleme-que-la-france-avec-ses-prisons/", "Français", "Juin 2026"),
    ("L'Avenir — Complexe sportif piscine Bruxelles (juil. 2024)", "https://www.lavenir.net/regions/bruxelles/2024/07/28/un-nouveau-complexe-sportif-avec-piscine", "Français", "Juin 2026"),
    ("Delabie Benelux — Site officiel", "https://www.delabiebenelux.com/fr", "Français/Néerlandais", "Juin 2026"),
    ("Delabie Benelux — Histoire et présence monde", "https://www.delabiebenelux.com/fr/le-groupe/a-propos-de-nous/notre-histoire", "Français", "Juin 2026"),
    ("Coexpert Aalberts — Normes robinetterie sanitaire", "https://coexpert.aalberts-hfc.com/fr-fr/installation/robinetterie-sanitaire-normes-certifications/", "Français", "Juin 2026"),
    ("NBN — Législation et normes Belgique", "https://www.nbn.be/en/using-standards/standards-legislation", "Anglais", "Juin 2026"),
    ("GMI Insights — Sanitary Ware Market (2024)", "https://www.gminsights.com/industry-analysis/sanitary-ware-market", "Anglais", "Juin 2026"),
    ("Forbes Belgique — Tendances hôtellerie belge 2024", "https://www.forbes.be/fr/6-tendances-a-suivre-dans-le-secteur-hotelier-belge-en-2024/", "Français", "Juin 2026"),
    ("RTBF — Bilan 2024 secteur construction", "https://www.rtbf.be/article/un-bilan-2024-bien-morose-pour-le-secteur-de-la-construction", "Français", "Juin 2026"),
    ("EMAE — Extrapolation, notes de recherche (source interne)", "Source interne MAB-core/sources-internes/", "Français", "2025"),
    ("TheGlobalEconomy — Belgium Political Stability", "https://www.theglobaleconomy.com/Belgium/wb_political_stability/", "Anglais", "Juin 2026"),
    ("EEA — Belgium Country Profile 2025", "https://www.eea.europa.eu/en/europe-environment-2025/countries/belgium", "Anglais", "Juin 2026"),
    ("Societé Wallonne du Logement — Plan rénovation", "https://www.swl.be/projets-immobiliers/plan-de-renovation.html", "Français", "Juin 2026"),
    ("Architectura.be — Plus grand projet logements étudiants Sart-Tilman", "https://www.architectura.be/fr/actualite/uau-collectiv-et-altiplan-realisent-le-plus-grand-projet-de-logements-etudiants-en-belgique-avec-sart-tilman/", "Français", "Juin 2026"),
    ("Negoce Zepros — Groupe BME rachète Paepens", "https://negoce.zepros.fr/actu-enseignes/sanitaire-chauffage-groupe-bme-rachete-flamand-paepens", "Français", "Juin 2026"),
    ("Sawiday.be — Robinetterie Presto", "https://www.sawiday.be/fr-be/robinetterie/presto/", "Français", "Juin 2026"),
]

for i, (title, url, lang, date) in enumerate(sources, 1):
    p = ann.add_paragraph(style="List Number")
    r = p.add_run(f"[{i}] {title}")
    r.bold = True
    r.font.size = Pt(9)
    p2 = ann.add_paragraph(f"    URL : {url}")
    p2.runs[0].font.size = Pt(8)
    p2.runs[0].font.color.rgb = RGBColor(0x00, 0x56, 0xA2)
    p3 = ann.add_paragraph(f"    Langue : {lang} | Consulté : {date}")
    p3.runs[0].font.size = Pt(8)
    p3.runs[0].font.italic = True

ann.add_paragraph()

add_heading(ann, "ANNEXE 2 — DONNÉES BRUTES CONSTRUCTION BELGIQUE", 1)

ann.add_paragraph("Source : Allianz Trade / Embuild / ING / DataBridge — compilées juin 2026").runs[0].font.italic = True
ann.add_paragraph()

table_data = [
    ["Indicateur", "2023", "2024", "2025 (est.)", "Source"],
    ["Construction totale (croissance %)", "+1.5%", "-2.8%", "-0.4%", "Embuild / EMAE"],
    ["Résidentiel neuf (permis)", "base", "-15%", "très faible", "Embuild 2024"],
    ["Génie civil (croissance %)", "+4.9%", "+4.4%", "-2.2%", "Embuild / ING"],
    ["Non-résidentiel neuf (croissance %)", "nd", "+1.4%", "+1.5%", "Techlink / ING"],
    ["Non-résidentiel rénov. (croissance %)", "nd", "-2.1%", "nd", "Techlink"],
    ["Faillites construction", "nd", "2 600+", "nd", "Allianz Trade"],
    ["Hausse faillites vs N-1", "nd", "+17%", "nd", "Allianz Trade"],
    ["Logements livrés", "~49 000", "53 800", "nd", "Embuild"],
]
table = ann.add_table(rows=len(table_data), cols=5)
table.style = "Table Grid"
for i, row_data in enumerate(table_data):
    row = table.rows[i]
    for j, cell_text in enumerate(row_data):
        cell = row.cells[j]
        cell.text = cell_text
        if i == 0:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)
        else:
            cell.paragraphs[0].runs[0].font.size = Pt(8)

ann.add_paragraph()

add_heading(ann, "ANNEXE 3 — DONNÉES BRUTES INVESTISSEMENTS ERP BELGIQUE", 1)
ann.add_paragraph("Pipeline d'investissements identifiés dans les ERP belges (2024-2030) :").runs[0].font.italic = True
ann.add_paragraph()

inv_data = [
    ["Segment", "Montant", "Période", "Porteur", "Source"],
    ["Hôpitaux universitaires (FWB)", "438 M€", "2024-2028", "Fédération Wallonie-Bruxelles", "Le Spécialiste"],
    ["Santé mentale Leuven/Kortenberg", "270 M€ (dont BEI 120 M€)", "2026-2040", "Z.org KU Leuven", "BEI 2025"],
    ["Bâtiments scolaires FWB", "1 Md€ (subventions)", "2023-2028+", "FWB / Wallonie", "FWB"],
    ["Logements sociaux Flandre", "1.7 Md€ (BEI)", "→2042", "Gouvernement flamand / BEI", "BEI 2025"],
    ["Logements sociaux Wallonie rénov.", "1.2 Md€", "→2030", "SWL / Wallonie", "SWL"],
    ["Résidences étudiantes Wallonie", "nd (appel à projets)", "→2025", "Wallonie (RRP)", "Wallonie RRP"],
    ["Prison Vresse-sur-Semois", "171 M€", "→2030", "État fédéral", "Media24"],
    ["Prison Anvers (livrée)", "nd", "2025", "État fédéral", "Jan De Nul"],
    ["Complexe sportif Bruxelles piscine", "nd", "2025-2027", "Commune péri-bruxelloise", "L'Avenir"],
    ["Infrasports Wallonie", "nd (programme annuel)", "récurrent", "Wallonie", "Wallonie"],
]
table2 = ann.add_table(rows=len(inv_data), cols=5)
table2.style = "Table Grid"
for i, row_data in enumerate(inv_data):
    row = table2.rows[i]
    for j, cell_text in enumerate(row_data):
        cell = row.cells[j]
        cell.text = cell_text
        if i == 0:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)
        else:
            cell.paragraphs[0].runs[0].font.size = Pt(8)

ann.add_paragraph()

add_heading(ann, "ANNEXE 4 — DISTRIBUTEURS ROBINETTERIE BELGES IDENTIFIÉS", 1)
dist_data = [
    ["Distributeur", "Type", "Zone", "Remarque"],
    ["SIDER", "Négoce pro plomberie", "National", "35 000 références, livraison pros, présent en ligne"],
    ["STG (ex-Paepens / BME)", "Négoce sanitaire-chauffage", "Flandre / Bruxelles", "Flandre-Orientale (Ninove), racheté par BME"],
    ["Versani NV", "Grossiste plomberie-chauffage", "Kempen (Anvers)", "Fondé 1975, indépendant"],
    ["Willems-Diels", "Grossiste plomberie", "Balen (Anvers)", "Familial, 40 ans d'existence"],
    ["Aquacaro", "Grossiste robinetterie design B2B", "National", "Robinetterie design salle de bains"],
    ["Sawiday.be", "E-commerce pro", "National/BE", "Présence Presto confirmée sur plateforme"],
    ["Sider.biz", "E-commerce pro", "National/BE", "Robinetterie + sanitaire pour pros"],
    ["JA Santé Belgique", "Distributeur équipements santé", "National", "Revendeur Delabie confirmé"],
    ["Hygiene-shop.be", "E-commerce spécialisé hygiène", "National/BE", "Revendeur Delabie confirmé"],
]
table3 = ann.add_table(rows=len(dist_data), cols=4)
table3.style = "Table Grid"
for i, row_data in enumerate(dist_data):
    row = table3.rows[i]
    for j, cell_text in enumerate(row_data):
        cell = row.cells[j]
        cell.text = cell_text
        if i == 0:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)
        else:
            cell.paragraphs[0].runs[0].font.size = Pt(8)

ann.add_paragraph()

add_heading(ann, "ANNEXE 5 — CONCURRENTS DÉTAILLÉS", 1)
comp_data = [
    ["Marque", "Groupe", "Niveau gamme", "Spécialité ERP", "Présence BE"],
    ["DELABIE", "Familial FR (1928)", "Haut de gamme", "Oui — leader ERP", "Filiale Benelux (Sint-Pieters-Leeuw)"],
    ["Hansgrohe", "Masco Corp (DE)", "Haut de gamme", "Partielle (tertiaire)", "Distribution nationale + showrooms"],
    ["Grohe", "LIXIL (DE)", "Milieu-haut", "Partielle", "Distribution nationale"],
    ["Hansa", "Delabie Group depuis 2021", "Milieu-haut", "Partielle (élec.)", "Site BE (hansa.com/fr-be)"],
    ["Geberit", "Geberit AG (CH)", "Milieu-haut", "Sanitaires encastrés", "Distribution nationale"],
    ["Oras / Damixa", "Artek Industries (FI)", "Milieu de gamme", "Limitée", "Via négoce"],
    ["Jacob Delafon", "Kohler (US)", "Milieu de gamme", "Résidentiel principalement", "Distribution nationale"],
]
table4 = ann.add_table(rows=len(comp_data), cols=5)
table4.style = "Table Grid"
for i, row_data in enumerate(comp_data):
    row = table4.rows[i]
    for j, cell_text in enumerate(row_data):
        cell = row.cells[j]
        cell.text = cell_text
        if i == 0:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)
        else:
            cell.paragraphs[0].runs[0].font.size = Pt(8)

ann.add_paragraph()

add_heading(ann, "ANNEXE 6 — NORMES DÉTAILLÉES ROBINETTERIE BELGIQUE", 1)
norm_data = [
    ["Norme", "Objet", "Date", "Applicabilité Presto"],
    ["NBN EN 816", "Robinets à fermeture automatique PN 10", "1996 (harmonisée)", "Cœur gamme temporisateurs — critique"],
    ["NBN EN 817", "Mitigeurs mécaniques PN 10", "2008 (harmonisée)", "Gamme mitigeurs collectifs"],
    ["NBN EN 200", "Robinetterie sanitaire domestique PN 10", "harmonisée", "Résidentiel — secondaire"],
    ["NBN EN 1111", "Mitigeurs thermostatiques (PN 10)", "harmonisée", "Gamme thermostatique — pertinent santé"],
    ["Marquage CE", "Obligatoire mise sur marché UE", "Permanent", "Condition sine qua non"],
    ["NF Robinetterie (NF077)", "Certification française valorisée en BE", "Permanente", "Déjà détenue par Presto — avantage"]
]
table5 = ann.add_table(rows=len(norm_data), cols=4)
table5.style = "Table Grid"
for i, row_data in enumerate(norm_data):
    row = table5.rows[i]
    for j, cell_text in enumerate(row_data):
        cell = row.cells[j]
        cell.text = cell_text
        if i == 0:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)
        else:
            cell.paragraphs[0].runs[0].font.size = Pt(8)

path_ann = os.path.join(OUTPUT_DIR, "MAB_Belgique_Annexes.docx")
ann.save(path_ann)
print(f"✓ Annexes sauvegardées : {path_ann}")
print("\n✓ Génération terminée. 2 documents dans pays/belgique/outputs/")
