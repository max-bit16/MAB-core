#!/usr/bin/env python3
"""
Contrôle qualité structurel pour les études MAB.
Vérifie : page de garde, ordre des sections, absence de doublon, absence d'emoji.
Usage : python3 check_mab_structure.py MAB_PAYS_Etude.docx
Exit code 0 = OK, 1 = problème détecté (bloquant pour livraison).
"""
import sys
import zipfile
import re

EXPECTED_ORDER = [
    "RÉSUMÉ EXÉCUTIF",
    "PARTIE 1",
    "PARTIE 2",
    "PARTIE 3",
    "PARTIE 4",
    "PARTIE 5",
    "PARTIE 6",
    "PARTIE 7",
    "PARTIE 8",
    "PARTIE 9",
]

def extract_text(docx_path):
    """Extract plain text from DOCX via zipfile (zero external deps)."""
    with zipfile.ZipFile(docx_path, 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
    texts = re.findall(r'<w:t[^>]*>([^<]+)</w:t>', xml)
    return ' '.join(texts)

def check(docx_path):
    text = extract_text(docx_path)
    errors = []
    warnings = []

    # 1. Page de garde : titre attendu dans les ~500 premiers caractères
    head = text[:500]
    if "ÉTUDE DE MARCHÉ" not in head.upper() and "ANNEXES" not in head.upper():
        errors.append("Page de garde manquante ou titre non trouvé en tête de document")
    if "Les Robinets Presto" not in head and "ROBINETS PRESTO" not in head.upper():
        warnings.append("Sous-titre Presto non détecté en tête de document")

    # 2. Ordre des sections + doublons
    positions = {}
    for marker in EXPECTED_ORDER:
        matches = [m.start() for m in re.finditer(re.escape(marker), text)]
        if not matches:
            errors.append(f"Section absente : {marker}")
            continue
        if len(matches) > 1:
            errors.append(f"DOUBLON détecté : '{marker}' apparaît {len(matches)} fois")
        positions[marker] = matches[0]

    # vérifier ordre croissant
    ordered_found = [(positions[m], m) for m in EXPECTED_ORDER if m in positions]
    sorted_check = sorted(ordered_found)
    if ordered_found != sorted_check:
        errors.append("Ordre des sections incorrect (ne suit pas 1→9)")

    # 3. Résumé Exécutif doit être avant Partie 1
    if "RÉSUMÉ EXÉCUTIF" in positions and "PARTIE 1" in positions:
        if positions["RÉSUMÉ EXÉCUTIF"] > positions["PARTIE 1"]:
            errors.append("Résumé Exécutif positionné APRÈS Partie 1")

    # 4. Emojis / pictogrammes interdits
    emoji_pattern = re.compile(
        "[\U0001F300-\U0001FAFF☀-➿←-⇿⬀-⯿]"
    )
    found_emojis = set(emoji_pattern.findall(text))
    found_emojis -= {"→", "—", "–"}
    if found_emojis:
        warnings.append(f"Pictogrammes potentiellement non conformes détectés : {found_emojis}")

    # Rapport
    print(f"\n=== Contrôle structure : {docx_path} ===")
    if errors:
        print(f"\n❌ {len(errors)} ERREUR(S) BLOQUANTE(S) :")
        for e in errors:
            print(f"   - {e}")
    if warnings:
        print(f"\n⚠ {len(warnings)} avertissement(s) :")
        for w in warnings:
            print(f"   - {w}")
    if not errors and not warnings:
        print("\n✅ Structure conforme.")

    return len(errors) == 0

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 scripts/check_mab_structure.py <fichier.docx>")
        sys.exit(1)
    ok = check(sys.argv[1])
    sys.exit(0 if ok else 1)
