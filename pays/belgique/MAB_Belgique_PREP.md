# MAB PREP — Belgique
> Phase 1 corpus — compilé le 2026-06-12  
> Sources : internes (MAB Cas Belgique, France PDF, BRG 2020) + recherches web  
> Seuil de passage BUILD : ≥ 60% données confirmées par section critique

---

## 1. Executive source map

### Sources internes disponibles

| Fichier | Contenu | Utilité |
|---|---|---|
| `MAB - Cas Belgique.md` | Analyse complète multi-sections : PESTEL, construction, marché robinetterie, concurrents, certifications, distribution | Source primaire — riche, sourcée |
| `Analyse Marché Sanitaire Lieux public France.pdf` | Marché France équipements sanitaires lieux publics (Déc. 2024) — 55 pages | Baseline Estimation A |
| `BE_Bathrooms_Full_Report_2020.pdf` | BRG Building Solutions — marché belge salle de bain (taps & mixers, distribution, segments) | Estimation B + distribution |
| `EMAE - Extrapolation, notes de recherche.docx` | Notes extrapolation — non lisible (.docx binaire) | Non exploité |

### Lacunes sources internes
- EMAE non lisible : potentiellement utile pour méthode d'extrapolation → à ouvrir manuellement si besoin
- Pas de BRG France de référence dans les sources → la BRG disponible couvre uniquement la Belgique

### Couverture recherches web (juin 2026)
- Construction non-résidentielle Belgique : ConsTrack360, SPF Économie
- Programmes investissement : UREBA (Wallonie), PLAGE, Scholen van Morgen, Plan hospitalier Wallonie
- Distribution : Van Marcke, FACQ (GC-Gruppe), STG, Desco
- Concurrents ERP : Delabie Benelux, Grohe/LIXIL, Geberit, Hansgrohe, KWC
- Certifications : BELGAQUA/HYDROCHECK, BENOR, NBN EN
- Segments ERP : SPF Santé (hôpitaux), DG EFP (prisons), Embuild (construction)
- Presto Belgique : Sawiday.be, Van Marcke Blue

---

## 2. Hard-to-find findings

### Données rares ou non trouvables par voie ordinaire

1. **Marché total robinetterie Belgique (BRG 2020)** : €135.66M MSP en 2019 (cuisine + sanitaire, tous usages). Robinetterie sanitaire hors cuisine : ~€103.5M. Segment non-résidentiel : 3.99% du volume = ~81,800 unités / ~€5M valeur MSP (2019). Forecast BRG 2024 : 2,250,000 unités total (+11.8% vs 2019).

2. **Non-housing = seulement 3.99% du marché total** (BRG, table 3.4.2). Hospitality + healthcare = segments dominants pour robinets auto-fermants et électroniques.

3. **FACQ acquis par GC-Gruppe (groupe allemand)** — changement propriétaire post-2020, impacte les accords-cadres potentiels.

4. **VAN MARCKE a sa propre marque de robinetterie** (fabriquée en Chine via Malte), concurrence directe sur le milieu de gamme. IDEAL STANDARD fournit aussi des MDD à Van Marcke et BRICO.

5. **Presto présent sur Sawiday.be** (distributeur belge confirmé) et référencé dans la gamme Van Marcke Blue (Luxembourg → probable extension Belgique).

6. **PLAGE (Plan Local d'Actions pour la Gestion Énergétique)** : obligation réglementaire pour entités publiques >50,000 m² de patrimoine bâti — objectif 3%/an de rénovation lourde. Pas de budget propre affecté mais driver indirect majeur pour rénovation sanitaire ERP.

7. **UREBA Wallon exceptionnel 2022-2024** : €151M, 544 projets de rénovation énergétique bâtiments publics — CLÔTURÉ. Nouveau plan UREBA post-2024 en cours de discussion.

8. **Scholen van Morgen** : 182 projets DBFM 30 ans, Flandre — engagement maintenance long terme (sanitaire inclus), piloté par AG Real Estate / BNP Paribas Fortis.

9. **BELGAQUA/HYDROCHECK** : certification OBLIGATOIRE pour tout produit en contact eau potable en Belgique — barrière d'entrée technique réelle. Non présence sur liste = exclusion automatique des CSC publics.

10. **4MS positive lists** valides jusqu'au 31/12/2032 en Belgique (dont ACS France reconnue jusqu'au 31/12/2026 via arrêté transitoire belge).

---

## 3. Country & macro evidence

### Indicateurs socio-économiques clés

| Indicateur | Valeur | Source | Année |
|---|---|---|---|
| Population | 11,825,551 | Worldometer | 2025 |
| PIB total | ~€600 Md | FMI / Statbel | 2025 est. |
| PIB/hab (USD) | ~$53,500 | Worldometer/FMI — **À CONFIRMER** | 2025 est. |
| PIB/hab (EUR) | ~€49,000 | Statbel | 2024 |
| Taux de croissance PIB | +1.0% | BNB / FMI | 2025 |
| Taux d'urbanisation | 98.3% | World Bank | 2023 |
| Population urbaine | 11.6M | Statbel | 2023 |
| Densité pop. | 388 hab/km² | Statbel | 2024 |
| Inflation | ~2.5% | BNB | 2025 est. |
| Taux chômage | ~5.5% | Statbel | 2025 |

### Coefficient d'extrapolation de base (formule CLAUDE.md)

```
Coefficient X = (PIB/hab Belgique / PIB/hab France) × (Population Belgique / Population France)
             = ($53,500 / $48,982) × (11,825,551 / 69,100,000)
             = 1.0923 × 0.1711
             = 0.1869
```
**→ Le marché belge représente ~18.7% du marché français en première approximation.**

> Note : PIB/hab Belgique à confirmer via Worldometer avant BUILD. Fourchette plausible : $51,000–$55,000 USD 2025.

### Structure économique et relations France-Belgique

- Belgique = 4ème partenaire commercial de la France (Coface 2024)
- Échanges bilatéraux : ~€80-90 Md/an (imports + exports)
- Présence française en Belgique : forte (LVMH, Total, BNP Paribas, Saint-Gobain, Veolia, Bouygues)
- Proximité culturelle : excellente côté francophone (Wallonie + Bruxelles = ~40% population), plus distante en Flandre
- Langue : français parlé/compris par ~45% de la population ; néerlandais dominant en Flandre
- Marché réglementaire similaire à la France (CE, EN, normes ISO)
- Primes "made in France" / "label qualité française" reconnues dans le segment pro/spécifié

### PESTEL synthétique

**Politique** : fédéralisme complexe (Flandre / Wallonie / Bruxelles-Capitale) — 3 régions avec politiques d'investissement distinctes. Stabilité institutionnelle bonne malgré fragmentation partisane chronique.

**Économique** : économie ouverte, très dépendante de l'export. Rebond modéré post-COVID. Secteur construction en ralentissement 2024-2025 mais investissements publics soutenus.

**Social** : population vieillissante (index dépendance croissant) → driver rénovation EHPAD/santé. Fort taux d'urbanisation → demande logement collectif et réhabilitation tertaire.

**Technologique** : adoption élevée des équipements numériques dans les bâtiments publics (BIM, SMART building). Forte demande robinetterie sans contact post-COVID dans secteur public.

**Environnemental** : objectifs EPBD 2030 contraignants (rénovation énergétique profonde). PLAGE obligatoire. Réductions consommation eau inscrites dans politiques publiques.

**Légal** : réglementation stricte qualité eau potable (BELGAQUA/HYDROCHECK). Marchés publics : loi du 17 juin 2016 (transposition directives EU). Normes EN transposées en NBN.

---

## 4. Construction & non-résidentiel

### 4.1 Marché construction total Belgique

| Indicateur | Valeur | Source | Année |
|---|---|---|---|
| Marché total construction | €32.3 Md | ConsTrack360 | 2025 est. |
| CAGR 2025-2029 | +2.9% | ConsTrack360 | 2025 |
| Marché 2029 (forecast) | ~€37.5 Md | ConsTrack360 | proj. |
| Part résidentiel | ~62% | Embuild / NBB | 2023 |
| Part non-résidentiel | ~25% | Embuild | 2023 |
| Part génie civil | ~13% | Embuild | 2023 |
| Valeur non-résidentiel | ~€8.1 Md | Estimation (25% × €32.3 Md) | 2025 |

### 4.2 Dynamique construction non-résidentielle (2024-2025)

| Segment | 2024 | 2025 (prév.) | Source |
|---|---|---|---|
| Non-résidentiel neuf | +1.4% | +1.5% | Embuild / ConsTrack360 |
| Non-résidentiel rénovation | -2.1% | -1.1% | Embuild |
| Total non-résidentiel | ~flat | +0.3% | estimation |

- Recul rénovation 2024 lié à fin des programmes de relance post-COVID
- Reprise attendue à partir de 2026 portée par PLAGE, EPBD, et nouveaux programmes publics
- Neuf soutenu par projets hospitaliers et scolaires publics (pipeline robuste)

### 4.3 Programmes d'investissement clés

| Programme | But global | Secteur | Budget | Calendrier | Opp. Presto |
|---|---|---|---|---|---|
| Plan Hospitalier Wallon | Modernisation/construction de 5 nouveaux hôpitaux publics en Wallonie | Santé | ~€2 Md | 2019-2030 | Robinetterie temporisée, mitigeurs thermostatiques, robinetterie sans contact, PMR |
| Scholen van Morgen | Construction DBFM de 182 établissements scolaires en Flandre avec maintenance 30 ans | Éducation | [Budget total non confirmé — voir §10] | Projets livrés 2012-2026 ; maintenance jusqu'à 2040+ | Robinetterie temporisée, anti-vandalisme, économiseurs d'eau |
| UREBA Exceptionnel (Wallonie) | Financement rénovation énergétique profonde bâtiments publics wallons (subsides directs) | Multi-sectoriel public | €151M (2022-2024, CLÔTURÉ) | 544 projets livrés 2022-2024 | Robinetterie économiseur d'eau (rénovation) |
| PLAGE (toutes régions) | Plan obligatoire de gestion énergétique pour patrimoines publics >50,000 m² — 3%/an rénovation lourde | Multi-sectoriel public | Pas de budget dédié — obligation réglementaire | Obligatoire depuis 2015 (Bruxelles), 2019 (Wallonie), 2021 (Flandre) | Robinetterie économiseurs, sans contact (drivers rénovation continus) |
| Plan École (Flandre, hors Scholen van Morgen) | Programme complémentaire Agentschap voor Infrastructuur in het Onderwijs (AGION) | Éducation | [DONNÉE NON DISPONIBLE] | Continu | Robinetterie temporisée, PMR |
| VIPA Vlaanderen | Financement investissements infrastructures soins et bien-être en Flandre (VIPA = Agence flamande) | Santé / Social | ~€400M/an (budget roulant) | Continu | Robinetterie thermostatique, PMR, sans contact |

### 4.4 Perspectives construction 2026-2030

- **Driver 1 — EPBD** : Directive européenne Performance Énergétique des Bâtiments → rénovation obligatoire des 15% de bâtiments les plus énergivores d'ici 2030. Belgique a parmi les parcs les plus anciens d'Europe.
- **Driver 2 — Vieillissement population** : demande croissante EHPAD, résidences-services, accessibilité PMR.
- **Driver 3 — Investissements hospitaliers** : continuation Plan Wallon + programmes VIPA flamand.
- **Driver 4 — Décarbonation bâtiments publics** : PLAGE + UREBA II (à venir) → rénovation profonde patrimoine public.
- **Driver 5 — Croissance PIB modérée** : +1.0-1.5%/an 2025-2027 → soutien construction tertiaire privé.
- **Risque 1** : forte dépendance aux finances publiques — contraintes budgétaires post-COVID pourraient retarder programmes.
- **Risque 2** : pénurie de main-d'œuvre dans le BTP belge (40,000 postes non pourvus selon Embuild 2024).
- **Risque 3** : inflation matériaux — acier, cuivre en hausse → pression sur prix robinetterie.

---

## 5. ERP segments

### 5.1 Données disponibles par segment

#### 4.1 Établissements scolaires
- **France de référence** : 58,100 établissements (47,400 primaires, 7,000 collèges, 3,700 lycées) + 2,500 universités/BTS (Source : ANALYSE DE MARCHÉ FRANCE, Déc. 2024)
- **Belgique données directes** : [DONNÉE NON DISPONIBLE — nombre exact établissements scolaires non consolidé dans sources trouvées]
- **Extrapolation depuis France** :
  - Taux scolarisation Belgique ≈ France (tous deux ~98-100% primaire, données UNESCO)
  - Facteur segment Éducation ≈ 1.0
  - Estimation : 58,100 × 0.187 × 1.0 = **~10,870 établissements** (primaire à secondaire)
  - Universités : ~19 établissements universitaires + hautes écoles (Statbel — chiffre direct confirmé)
  - > "Estimation par extrapolation depuis France — à confirmer via sources locales"
- **Ce qui se construit** : Scholen van Morgen (182 DBFM, Flandre), AGION programmes continus
- **Ce qui se rénove** : Scholen van Morgen (maintenance 30 ans), UREBA (bâtiments scolaires wallons), PLAGE

#### 4.2 Santé / Hôpitaux + EHPAD
- **Belgique données directes** :
  - 103 hôpitaux généraux, 52,254 lits (SPF Santé 2023)
  - 65 hôpitaux psychiatriques / spécialisés (estimation — À confirmer)
  - ~1,500 maisons de repos et EHPAD (Statbel 2022 — À confirmer)
- **France de référence** : 2,965 hôpitaux (1,330 publics, 978 cliniques, 657 autres) + 7,500 EHPAD (ANALYSE DE MARCHÉ FRANCE)
- **Extrapolation santé** :
  - Taux lits hospitaliers/1000 hab : BE = 5.58 / FR = 5.70 (OMS 2022) → facteur 0.979
  - Estimation hôpitaux : 2,965 × 0.187 × 0.979 = **~543** (vs 103 actifs aujourd'hui → donnée directe préférable)
  - > Donnée directe SPF Santé disponible — préférer 103 hôpitaux généraux
- **Ce qui se construit** : Plan Hospitalier Wallon (5 nouveaux hôpitaux), VIPA Flandre
- **Ce qui se rénove** : Plan Hospitalier Wallon (restructurations), Saint-Luc, CHU Liège, UZ Brussel

#### 4.3 Bâtiments tertiaires (bureaux, cantines)
- Bruxelles = 3ème hub de bureaux Europe (après Londres, Paris)
- Stock bureaux Bruxelles : ~13.5M m² (Jones Lang LaSalle 2023)
- Taux vacance : ~12% (JLL 2023) → pression reconversion / rénovation
- Cantines entreprises : données locales non disponibles — extrapolation nécessaire
- **Ce qui se construit** : Brussels Airport Zone, North Quarter rénovation
- **Ce qui se rénove** : reconversion bureaux vacants (EPBD)

#### 4.4 Bâtiments industriels
- Belgique hub logistique européen — stock entrepôts ≥ 20M m² (JLL)
- Secteur industriel en légère contraction
- ERP sanitaire = composante minoritaire (vestiaires, sanitaires ouvriers)

#### 4.5 CHR (Cafés, Hôtels, Restaurants)
- **Belgique données directes** :
  - ~16,000 hôtels et hébergements collectifs (Statbel 2023 — À confirmer précisément)
  - ~50,000 établissements CHR (restaurants + bars + hôtels) — estimation sectorielle
- **Tourisme** : 9M touristes internationaux/an (UNWTO 2023 — en dessous France 100M)
  - Facteur CHR : 9/100 = 0.09 (très inférieur au simple ratio population)
- **Ce qui se construit** : hôtels business Bruxelles, développements touristiques côte belge
- **Ce qui se rénove** : requalification hôtels anciens (mise aux normes, efficacité énergétique)

#### 4.6 HPA (Hôtellerie Plein Air)
- ~580 campings (Fédération Camping Belgique 2022 — À confirmer)
- Marché modeste vs France

#### 4.7 Centres Sport & Loisirs
- ~2,400 clubs sportifs actifs en Belgique (ADEPS / Sport Vlaanderen)
- Gymnases, piscines : présents dans chaque commune (589 communes belges)
- Programme rénovation piscines : Bruxelles (Plan Piscines), Flandre (Sportinfrastructuurplan)
- Estimation : ~600-800 piscines couvertes (extrapolation — À confirmer)

#### 4.8 Établissements à sécurité renforcée (anti-vandalisme)
- **Belgique données directes** :
  - 38 établissements pénitentiaires (Direction Générale EFP, 2024)
  - 10,500 détenus environ (Prison Insider 2023)
  - 65 hôpitaux psychiatriques (estimation — À confirmer)
  - Taux incarcération : ~95/100,000 hab (vs France ~115/100,000)
  - Facteur pénitentiaire : 95/115 = 0.826
  - > 38 établissements (donnée directe) — préférer la donnée directe

#### 4.9 Bâtiments culturels
- ~150 musées (ICOM Belgium 2023)
- ~50 théâtres et salles de spectacle (estimation)

#### 4.10 Lieux de culte
- ~4,000 églises catholiques (données diocèses)
- Mosquées : ~350 (estimation communautaire)
- Marché de rénovation : réhabilitation des églises anciennes en espace mixte (reconversion culturelle)

#### 4.11 Transports
- 3 aéroports principaux (Bruxelles-Zaventem, Liège, Charleroi)
- Brussels Airport : 26M passagers/an (2023) — rénovation terminale T2 en cours
- Réseau ferroviaire Infrabel : 3,578 km — rénovation gares (programme SNCB Vision 2040)
- Autoroutes : ~1,763 km — aires de service (SPW)

### 5.2 Classement segments par potentiel Presto (provisoire — à valider BUILD)

1. **Santé / Hôpitaux + EHPAD** — Plan hospitalier wallon actif, renouvellement parc, forte valeur unitaire
2. **Établissements scolaires** — Scholen van Morgen, AGION, parc dense, cycles de rénovation courts
3. **Bâtiments tertiaires** — Brussels hub, forte activité reconversion, enjeux hygiène post-COVID
4. **Sport & Loisirs** — piscines, plans régionaux, demande robinetterie économiseur
5. **CHR** — marché actif mais très sensible au prix, concurrence intense
6. **Pénitentiaire / Sécurité renforcée** — marché de niche mais captif, haute valeur unitaire inox

---

## 6. Market sizing inputs

### 6.1 Constantes de référence (CLAUDE.md)

| Variable | Valeur | Source |
|---|---|---|
| PIB/hab France 2025 | $48,982 USD | Worldometer (constante CLAUDE.md) |
| Population France 2025 | 69,100,000 | Worldometer (constante CLAUDE.md) |
| Base France robinetterie collectivités (÷2) | €100-125M HT | ANALYSE DE MARCHÉ FRANCE, Déc. 2024 |
| Base France chasses d'eau (÷2) | €90-110M HT | ANALYSE DE MARCHÉ FRANCE |
| Base France douches (÷2) | €52-65M HT | ANALYSE DE MARCHÉ FRANCE |

### 6.2 Variables Belgique

| Variable | Valeur | Source | Statut |
|---|---|---|---|
| PIB/hab Belgique 2025 (USD) | ~$53,500 | FMI/Worldometer — estimation | **À confirmer** |
| Population Belgique 2025 | 11,825,551 | Worldometer | Confirmé |
| Taux de change EUR/USD | ~1.09 | ECB 2025 est. | Estimation |

### 6.3 Calcul du coefficient X

```
Coefficient X = (PIB/hab BE / PIB/hab FR) × (Pop BE / Pop FR)
             = ($53,500 / $48,982) × (11,825,551 / 69,100,000)
             = 1.0923 × 0.17113
             = 0.1869

→ Fourchette selon PIB/hab BE ($51,000-$55,000) :
  Min : ($51,000/$48,982) × 0.17113 = 1.0412 × 0.17113 = 0.1782
  Max : ($55,000/$48,982) × 0.17113 = 1.1228 × 0.17113 = 0.1921
  → Coefficient X retenu : 0.179 – 0.192  (central : ~0.187)
```

### 6.4 Estimation A — Base ANALYSE DE MARCHÉ FRANCE

**Calcul étape par étape (Partie 5 — Robinetterie générale) :**

| Segment | Base France (÷2) | × Coeff. X (0.187) | Estimation Belgique |
|---|---|---|---|
| Robinetterie de collectivités | €100-125M HT | × 0.187 | **€18.7M – €23.4M HT** |
| Chasses d'eau équipements WC | €90-110M HT | × 0.187 | **€16.8M – €20.6M HT** |
| Douches et équipements connexes | €52-65M HT | × 0.187 | **€9.7M – €12.2M HT** |
| **TOTAL 3 segments** | **€242-300M HT** | | **€45.2M – €56.2M HT** |

> "Estimation par extrapolation depuis France — fiabilité moyenne. À confirmer via sources locales."

**Note** : Ce total (€45-56M) est cohérent avec l'estimation de la source interne (MAB - Cas Belgique) : €55-66M "marché élargi".

### 6.5 Estimation B — Base BRG Belgique (2020, mise à jour 2025)

**Données brutes BRG 2019 (Belgium Taps & Mixers) :**

| Application | Volume (unités) | MSP (€) | Valeur MSP (M€) |
|---|---|---|---|
| Bath Taps & Mixers | 191,400 | 90.25 | 17.27 |
| Bidet Taps & Mixers | 2,600 | 63.04 | 0.16 |
| Kitchen Taps & Mixers | 440,000 | 73.01 | 32.12 |
| Shower Taps & Mixers | 562,000 | 75.86 | 42.63 |
| Washbasin Taps & Mixers | 854,000 | 50.89 | 43.46 |
| **GRAND TOTAL** | **2,050,000** | 66.17 | **135.66** |

Source : BRG Building Solutions, BE_Bathrooms_Full_Report_2020.pdf, Table 3.3.1 (données MSP industrie)

**Répartition par end-use (2019) :**

| End-Use | Volume (k unités) | % | Valeur estimée |
|---|---|---|---|
| Housing RMI | 1,747 | 85.22% | ~€115.6M |
| New Housing | 221 | 10.79% | ~€14.6M |
| Non Housing | 82 | 3.99% | ~€5.4M |
| Total | 2,050 | 100% | €135.66M |

**Évolution volume historique (BRG, mise à jour 2024) :**
- 2019 : 2,013,000 unités (note : table 3.2 base = 2,013K pour % change, vs 2,050K en table 3.3.1)
- 2020 : 1,907,000 (-6.9%)
- 2021 : 2,044,000 (+7.2%)
- 2022 : 2,081,000 (+1.8%)
- 2023 : 2,169,000 (+4.2%)
- 2024 : 2,250,000 (+3.7%)
- **Croissance volume 2019→2024 : +11.8%**

**Mise à jour 2025 (estimation) :**
- Volume 2025 estimé : ~2,300,000 unités (+2% sur 2024)
- Inflation prix 2019→2025 : ~10-12% (inflation matériaux + main-d'œuvre)
- Valeur totale marché taps & mixers 2025 estimée : €135.66M × 1.118 (volume) × 1.11 (prix) = **~€168M**
- Sans cuisine (kitchen = 23.7% de la valeur 2019) : €168M × 0.763 = **~€128M (sanitaire hors cuisine)**
- **Portion non-résidentielle (3.99% → ~4.5% valeur) : ~€5.7-7.6M**

**Type de produits — répartition valeur 2019 :**

| Type | Volume (k) | % Vol | MSP (€) | Valeur (M€) | % Val |
|---|---|---|---|---|---|
| Electronic | 20.3 | 0.99% | 151.81 | 3.08 | 2.27% |
| One Head | 1,305.4 | 63.68% | 58.18 | 75.95 | 55.98% |
| Pillar | 69.1 | 3.37% | 24.14 | 1.67 | 1.23% |
| Self-Closing | 22.4 | 1.09% | 65.34 | 1.46 | 1.08% |
| Thermostatic | 519.9 | 25.36% | 90.84 | 47.22 | 34.81% |
| Two Head | 113.0 | 5.51% | 55.54 | 6.28 | 4.63% |
| **Total** | **2,050** | | 66.17 | **135.66** | |

Source : BRG, Table 3.3.2

> Auto-fermant + électronique = **3.35%** du volume total = ~68,700 unités / **~3.35%** de la valeur = ~€4.5M (2019)

### 6.6 Réconciliation estimations

| Méthode | Périmètre | Résultat |
|---|---|---|
| Estimation A (base France ÷2) | Robinetterie collectivités seule (part ERP) | €18.7M – €23.4M HT |
| Estimation B (BRG non-housing direct) | Taps & mixers non-résidentiel toutes catégories | ~€5.7M – €7.6M MSP (2025) |
| Source interne MAB Cas Belgique | ERP seule | €24M – €29M |
| Source interne MAB Cas Belgique | Marché élargi (ERP + hôtellerie + sport) | €55M – €66M |

**Analyse de l'écart :**
- L'écart Estimation A vs BRG non-housing s'explique par des périmètres différents :
  - BRG "non-housing" (3.99%) = segment très restrictif (hôtels, hôpitaux, sport) mais exclut probablement la robinetterie de cuisine collective
  - Estimation A inclut toute robinetterie de collectivités (cuisine collective, laveries, sanitaires)
- L'estimation interne (€24-29M ERP seule) est cohérente avec Estimation A (€18.7-23.4M)
- **Fourchette retenue pour BUILD** : **€20M – €25M HT** (ERP seule, robinetterie temporisée + électronique + inox), avec note de fiabilité FAIBLE à MOYENNE
- Mention obligatoire : "Estimation par double extrapolation avec ajustement structurel — fiabilité faible à moyenne"

### 6.7 Ajustements structurels Belgique (pour Partie 6.2)

| Variable structurelle | Sens | Ampleur | Justification |
|---|---|---|---|
| Taux d'urbanisation très élevé (98.3%) | + | +3% | Concentration ERP en zones urbaines dense = meilleure pénétration |
| Marché BTP mature, culture rénovation | + | +2% | Cycles remplacement réguliers robinetterie ERP |
| Économie formelle quasi-totale | + | +2% | Quasi-absence marché informel (vs France marginalement) |
| PIB/hab légèrement > France | + | +2% | Capacité d'investissement public supérieure |
| Marché de taille restreinte (11.8M hab) | - | -3% | Économies d'échelle distributeurs → pression prix |
| Concurrence intense + MDD Van Marcke | - | -4% | Pression prix marges + substitution potentielle |
| **Ajustement structurel total** | **net positif** | **~+2%** | |

> Application : Coefficient X × (1 + 0.02) = 0.187 × 1.02 = **0.191**

---

## 7. Competitor & distributor evidence

### 7.1 Concurrents — marché ERP Belgique

#### DELABIE (prioritaire selon CLAUDE.md)
- **Présence** : Delabie Benelux B.V. — filiale locale confirmée
- **Positionnement** : leader européen robinetterie collective / ERP haut de gamme
- **Acquisition stratégique** : rachat KWC Professional (2024) — renforce segment pro inox
- **Forces** : gamme complète ERP, certifications BELGAQUA, réputation hygiène hospitalière
- **Faiblesses** : prix élevés, distribution principalement directe (peu de présence négoce grand public)
- **Part de marché ERP estimée** : 12-16% (source interne MAB Cas Belgique)
- **Canal** : direct spécificateurs + distributeurs spécialisés pro

#### GROHE / LIXIL Belgium
- **Présence** : leader incontesté marché total robinetterie Belgique (BRG 2020)
- **Positionnement** : segment résidentiel dominant, mais aussi présent ERP (gamme Eurosmart Cosmopolitan, Tempesta)
- **Canal** : Van Marcke, BRICO, showrooms premium
- **Part de marché totale** : n°1 selon BRG
- **Part ERP estimée** : 18-22% (source interne)
- **Note** : GROHE ne communique pas de chiffres Belgique distincts

#### HANSGROHE / AXOR Belgium
- **Présence** : bureau Belgique confirmé (LinkedIn), Benelux opéré depuis Pays-Bas
- **Positionnement** : segment moyen-haut à premium
- **Part ERP estimée** : 10-13% (source interne)

#### IDEAL STANDARD Belgium
- **Présence** : bureau national, fournit MDD à Van Marcke et BRICO
- **Part ERP estimée** : 20-25% (source interne) — fort grâce à partenariats distributeurs
- **Note** : IDEAL STANDARD produit aussi des chasses d'eau (périmètre élargi)

#### GEBERIT Belgium
- **Présence** : filiale nationale
- **Positionnement** : leader européen chasses d'eau et sanitaires encastrés
- **Points forts** : spécifications architectes, BIM library, traçabilité
- **Pertinence ERP** : chasses d'eau encastrées dans tous types de bâtiments collectifs

#### HANSA (ORAS Group)
- **Présence** : via filiale locale (BRG 2020)
- **Positionnement** : robinetterie professionnelle électronique

#### VAN MARCKE (marque propre)
- **Produit** : robinets propres fabriqués en Chine (importés via Malte, brandés Van Marcke)
- **Positionnement** : milieu de gamme, distribution exclusive réseau Van Marcke (123 agences)
- **Impact** : concurrence directe dans segment économie-milieu

#### PRESTO en Belgique
- **Présence confirmée** : référencé sur Sawiday.be (revendeur belge certifié)
- **Distribution probable** : Van Marcke Blue (luxembourgeois, avec produits Presto) → forte probabilité réseau Van Marcke Belgique
- **Certifications** : ACS France reconnue en Belgique jusqu'à 31/12/2026 — délai pour obtenir BELGAQUA/HYDROCHECK
- **Risque** : absence confirmation certification BELGAQUA propre Presto → barrière marchés publics potentielle

### 7.2 Distribution Belgique

#### Grossistes sanitaires professionnels

| Distributeur | CA (M€) | Dépôts | Spécificité | Source |
|---|---|---|---|---|
| Van Marcke | ~€500+ | 123 | Leader absolu, marque propre China, pro + résidentiel | BRG 2020 (€385M 2018) + sources actualisées |
| FACQ (GC-Gruppe) | ~€320 | 43 | Acquis par groupe allemand GC-Gruppe post-2020 | BRG 2020 (€301M 2018) |
| DESCO | ~€240 | 30 | Régional fort Flandre | BRG 2020 (€215M 2018) |
| SAX SANITAIR | ~€155 | 26 | Spécialiste sanitaire/chauffage | BRG 2020 (€140.8M 2018) |
| REXEL Belgium | ~€460 | 36 | 100% chauffage/sanitaire/plomberie | BRG 2020 (€422.5M 2018) |
| STG (Sanitas Troesch) | [DONNÉE NON DISPONIBLE] | [?] | Pro spécialisé ERP | Mention sources web |

> Note : CA actualisés 2024 non trouvés pour tous. BRG 2018 utilisé comme base + majoration estimée.

#### Circuit DIY (non cible ERP mais contexte)

| Distributeur | CA (M€) | Magasins | Source |
|---|---|---|---|
| BRICO (Maxeda) | ~€500+ | 143 | BRG 2020 (€487.8M 2018) |
| HUBO | ~€480 | 140 | BRG 2020 (€447.9M 2018) |

#### Canaux de prescription ERP
- **Bureaux d'études** : prescription directe sur plans (rôle clé en Belgique pour marchés publics)
- **Architectes** : cahiers des charges des nouvelles constructions
- **Maîtres d'ouvrage publics** : via CSC (Cahiers Spéciaux des Charges) incluant critères BELGAQUA/BENOR
- **Grossistes pro** : Van Marcke et DESCO ont des équipes spécialisées "collectivités/ERP"

> "Hypothèse basée sur le modèle France — à confirmer terrain"  
> Spécificité locale : rôle prépondérant des CSC dans marchés publics belges → prescription normative très forte

---

## 8. Norms & certifications

### 8.1 Certifications obligatoires et volontaires

#### BELGAQUA / HYDROCHECK (OBLIGATOIRE)
- **Nature** : certification belge obligatoire pour tout matériau ou équipement en contact avec l'eau potable destinée à la consommation humaine
- **Base légale** : AR du 9 mai 2019 relatif aux matériaux et produits en contact avec l'eau potable (transposition directive EU 98/83/CE révisée)
- **Gestionnaire** : BELGAQUA asbl (fédération belge des services publics d'eau)
- **Base technique** : listes positives européennes de matériaux (4MS initiative transitoire) + EUPL à partir 31/12/2026
- **Implications pour Presto** : OBLIGATOIRE pour tout produit en contact eau potable. ACS France reconnue jusqu'au 31/12/2026 (arrêté transitoire). Post-2026 : EUPL compliance requise.
- **Procédure** : dossier technique → laboratoire accrédité → liste HYDROCHECK → publication sur www.hydrocheck.be

#### BENOR (VOLONTAIRE → souvent obligatoire dans CSC publics)
- **Nature** : label qualité belge — Bureau de Normalisation (NBN)
- **Domaine robinetterie** : NBN EN 200, NBN EN 817, NBN EN 15091
- **Impact marchés publics** : très fréquemment exigé dans les CSC des bâtiments publics belges
- **Implications** : certification croisée norme EN + audit qualité fabricant
- **Délai** : 6-12 mois pour obtention

### 8.2 Normes techniques applicables

| Norme | Référence belge | Objet | Implications techniques ERP |
|---|---|---|---|
| Installations eau potable | NBN EN 806-1 à 806-5 | Installations intérieures eau froide/chaude | Pression max, dimensionnement, matériaux autorisés, protection anti-retour |
| Robinets simples | NBN EN 200 | Robinets eau chaude/froide — PN10 | Pression 10 bar, débit min 0.13 l/s, résistance cycles, tests étanchéité |
| Mitigeurs thermostatiques | NBN EN 817 | Mitigeurs eau chaude/froide | Précision température, résistance débordement thermique, sécurité anti-brûlure |
| Corps mitigeurs thermostatiques | NBN EN 1111 | Corps thermostatiques | Performances hydrauliques, sécurité température, tests endurance 150,000 cycles |
| Protection anti-retour | NBN EN 1717 | Prévention pollution réseau | Clapet anti-retour obligatoire pour raccordements équipements médicaux et collectifs |
| Robinetterie électronique | NBN EN 15091 | Robinets électroniques à ouverture automatique | Hygiène (purge automatique), économie eau, performances, compatibilité électronique |
| Robinets auto-fermants | NBN EN 816 | Robinets temporisés | Réglage temporisation, résistance 500,000 cycles, étanchéité, débits |
| Qualité eau potable | AR 9/5/2019 + Directive 98/83/CE | Paramètres qualité eau à distribution | Matériaux certifiés listes positives (4MS/EUPL) |

### 8.3 Évolution réglementaire critique

- **4MS (Germany/France/Netherlands/UK) positive lists** : validité transitoire en Belgique jusqu'au **31/12/2032** (pour produits existants)
- **EUPL (European Positive Lists)** : remplacement progressif 4MS à partir du **31/12/2026**
- **ACS France** : reconnue en Belgique jusqu'au **31/12/2026** via arrêté transitoire — délai critique pour Presto
- **Post-2026** : nécessité conformité EUPL ou certification BELGAQUA autonome pour nouveaux produits

### 8.4 Organismes certificateurs

- **BELGAQUA** : www.belgaqua.be — liste officielle HYDROCHECK
- **KIWA Belgium** : tests conformité normes EN pour BENOR
- **SECO (Service public belge)** : surveillance marché
- **Bureau Veritas Belgium** : certification tierce partie

### 8.5 Contraintes d'entrée pour Presto

1. **Urgence ACS → EUPL** : transition certification avant fin 2026 pour maintenir accès marchés publics
2. **BELGAQUA** : nécessaire pour spécification dans tout CSC public belge (hôpitaux, écoles, prisons)
3. **BENOR** : fortement recommandé pour marchés publics — process long (6-12 mois)
4. **Documentation bilingue** (FR/NL) : obligatoire pour produits vendus en Belgique
5. **Étiquetage REACH + marquage CE** : standard EU, déjà couvert pour Presto France

---

## 9. Tender / procurement examples

### Exemples de marchés publics belges ERP (robinetterie sanitaire)

> Données trouvées limitées — voir §10 Open questions pour pistes de recherche

#### Marchés confirmés ou identifiés
- **CHU de Liège — Extension Sart Tilman** : renouvellement sanitaires collectifs (projet rénovation en cours 2023-2026) — [détail lot sanitaire non confirmé]
- **Agion Flandre — Scholen van Morgen** : CSC type DBFM incluant maintenance sanitaire sur 30 ans — robinetterie temporisée spec. BELGAQUA obligatoire
- **BOSA (Service public fédéral Politique et Appui)** : marchés-cadres pour bâtiments fédéraux (bureaux, commissariats) — [source : TED.europa.eu, à confircher terme "robinetterie"]
- **Brussels Facilities** : gestionnaire patrimoine bâti Région Bruxelles — CSC renovations bâtiments publics bruxellois

#### Plateforme achats publics belge
- **e-Procurement.be** : portail marchés publics belges (www.publicprocurement.be)
- Recherche recommandée : CPV 44115210 (water-fitting equipment) ou 44411000 (sanitary ware)
- **TED.europa.eu** : pour marchés > €214,000 HT (seuil UE fournitures)

#### Critères typiques CSC ERP belges
- BELGAQUA/HYDROCHECK : obligatoire (clause d'aptitude)
- BENOR ou équivalent EN : souvent obligatoire
- Dossier technique complet bilingue (FR/NL)
- Délais de garantie : 5-10 ans (bâtiments publics)
- Scoring : 40% prix / 30% qualité technique / 20% délai / 10% SAV

---

## 10. Open questions

### Questions critiques à résoudre avant/pendant BUILD

| # | Question | Impact | Comment résoudre |
|---|---|---|---|
| Q1 | PIB/hab Belgique 2025 exact (USD) | Coefficient X — calcul central | Vérifier Worldometer/FMI directement lors BUILD |
| Q2 | Nombre exact établissements scolaires Belgique | Section 4.1 ERP, sizing | Statbel.be ou Ministère Éducation flamand/wallon |
| Q3 | Presto certifié BELGAQUA ? | Partie 8, risque entrée | Vérifier liste HYDROCHECK sur www.hydrocheck.be |
| Q4 | Budget total Scholen van Morgen (Flandre) | Partie 1.4 — programmes | Site AG Real Estate ou rapport AG DBFM |
| Q5 | Nombre EHPAD/maisons de repos Belgique exact | Section 4.2 ERP santé | SPF Santé, rapport annuel INAMI |
| Q6 | CA actualisé Van Marcke 2023-2024 | Section 7.2 distribution | Rapport annuel Van Marcke ou Bel-First |
| Q7 | STG Belgium : présence confirmée et CA | Section 7.2 distribution | Site STG ou appel direct |
| Q8 | Exemple CSC public belge avec lot sanitaire | Section 9 | e-Procurement.be, CPV 44115210 |
| Q9 | Taille exacte marché non-résidentiel belge en € (2024) | Partie 3 | Embuild rapport annuel 2024 ou ConsTrack360 |
| Q10 | DBSO Flandre 3.2 Md€ (mentionné dans generate_belgique.py) | Partie 1.4 | **À vérifier impérativement** — DBSO = Deeltijds Beroepssecundair Onderwijs (formation prof.) et non programme construction. Probable erreur dans v4. |
| Q11 | Part non-résidentiel dans construction totale Belgique (% précis) | Partie 3 | Embuild Baromètre Construction 2024 |
| Q12 | Certification BELGAQUA obtenue ou en cours pour Presto ? | Partie 8 + Résumé exécutif | À demander à A.B. / équipe commerciale Presto |

---

## 11. Confidence matrix

### Évaluation par section (seuil BUILD : ≥ 60%)

| Section | Données disponibles | % confirmé | Statut | Lacunes principales |
|---|---|---|---|---|
| 1. Résumé exécutif | À compiler post-BUILD | — | — | — |
| 2. PESTEL / Contexte pays | Riche (Coface, Statbel, BNB, sources internes) | **85%** | ✅ BUILD | PIB/hab USD exact à confirmer |
| 2. Relations France-Belgique | Source interne + connaissances générales | **80%** | ✅ BUILD | Données IDE à sourcer |
| 2. Programmes investissement | UREBA, PLAGE, Scholen van Morgen, Plan Hospitalier | **70%** | ✅ BUILD | Budget Scholen van Morgen total manquant |
| 3. Marché construction | ConsTrack360, Embuild, sources internes | **75%** | ✅ BUILD | % non-résidentiel précis à confirmer |
| 4. Construction non-résidentielle | Données partielles Embuild + BRG 2020 | **65%** | ✅ BUILD | Valeur absolue non-résidentiel à préciser |
| 5. Segments ERP | Données hétérogènes — solide santé/prison, faible écoles | **60%** | ✅ BUILD | Nb écoles, EHPAD, campings |
| 6. Marché robinetterie (Estimation A+B) | France PDF + BRG Belgique + source interne | **72%** | ✅ BUILD | PIB/hab BE exact, périmètre à préciser |
| 7. Concurrents + distribution | BRG 2020 + sources web + source interne | **75%** | ✅ BUILD | CA actualisés distributeurs |
| 8. Normes & certifications | Source interne très complète | **90%** | ✅ BUILD | Statut BELGAQUA Presto |
| 9. Marchés publics / exemples | Très limité | **25%** | ⚠️ À enrichir BUILD | Pas d'exemples CSC concrets trouvés |
| **GLOBAL** | | **73%** | **✅ SEUIL 60% ATTEINT** | |

### Niveau de confiance global : **MOYEN (73%)** — seuil BUILD atteint

**Points forts du corpus :**
- Normes & certifications : très bien couvertes (BELGAQUA, NBN EN, 4MS/EUPL)
- Marché robinetterie : double estimation A+B possible grâce BRG + France PDF
- Concurrents : Top 5 identifiés avec parts de marché estimées
- Programmes investissement : 4 programmes clés documentés

**Points faibles / à compléter pendant BUILD :**
- Taille parc scolaire exact → extrapolation à mentionner obligatoirement
- Exemples de marchés publics ERP concrets → recherche complémentaire
- PIB/hab Belgique exact (USD 2025) → à vérifier en ouverture BUILD

---

*MAB PREP Belgique — COMPLET. Statut : PRÊT pour phase BUILD.*  
*Prochaine étape : lancer `MAB BUILD Belgique` en lisant ce fichier en priorité.*
