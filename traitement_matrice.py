#!/usr/bin/env python3
import sys
import glob
import re
import datetime as dt
import shutil
import json
import zipfile
import tempfile
from pathlib import Path
import ezodf
from odf.opendocument import load as odf_load
from odf.style import Style, TableCellProperties, ParagraphProperties, TextProperties
from odf.element import Element as ODFElement
from odf.namespaces import STYLENS, FONS, TABLENS, OFFICENS

# --- IMPORTS QT (PySide6) ---
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QHBoxLayout, QLabel, QPushButton, QListWidget,
                               QFileDialog, QTabWidget, QGroupBox, QLineEdit,
                               QCheckBox, QComboBox, QTreeWidget, QTreeWidgetItem,
                               QProgressBar, QMessageBox, QTextEdit, QDialog, QDialogButtonBox)
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction
from qt_material import apply_stylesheet

# ==============================================================================
# 1. LOGIQUE MÉTIER (BACKEND)
# ==============================================================================
# Cette section contient toute la logique de manipulation des fichiers ODS,
# indépendamment de l'interface graphique.

# --- Expressions Régulières Globales ---
A1_RE = re.compile(r"^([A-Za-z]+)(\d+)$")  # Pour parser les adresses type "A1", "B2", etc.
RANGE_RE = re.compile(r"^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$")  # Pour parser les plages type "A1:C5"
VER_RE = re.compile(r'(.*?)(v)(\d+(?:\.\d+)*)(.*)$', re.IGNORECASE)  # Pour trouver un numéro de version (ex: v1.2)
INVALID_FS_CHARS = set('/\\:*?"<>|')  # Caractères invalides dans les noms de fichiers


def a1_to_rc(a1: str):
    """
    Convertit une adresse de cellule de type A1 (ex: "B3") en coordonnées (ligne, colonne)
    basées sur un index zéro (ex: (2, 1)).

    Args:
        a1 (str): L'adresse de la cellule au format A1.

    Returns:
        tuple[int, int]: Un tuple (ligne, colonne) avec index à partir de zéro.

    Raises:
        ValueError: Si le format de l'adresse est invalide.
    """
    m = A1_RE.match(a1.strip())
    if not m: raise ValueError(f"Adresse invalide : {a1}")
    col_letters, row_str = m.groups()
    col = 0
    # Convertit les lettres de la colonne en un nombre (A=0, B=1, ..., Z=25, AA=26, etc.)
    for ch in col_letters.upper(): col = col * 26 + (ord(ch) - ord('A') + 1)
    col -= 1  # Ajustement pour un index basé sur zéro
    return int(row_str) - 1, col


def parse_range(rng: str):
    """
    Analyse une chaîne représentant une plage (ex: "A1:C5") ou une seule cellule ("B2")
    et retourne les coordonnées (ligne, colonne) de début et de fin.

    Args:
        rng (str): La chaîne représentant la plage.

    Returns:
        tuple[int, int, int, int]: Un tuple (ligne_début, col_début, ligne_fin, col_fin).
                                   Pour une cellule unique, les coordonnées de début et de fin sont identiques.
    """
    rng = rng.strip()
    m = RANGE_RE.match(rng)
    if m:
        # C'est une plage (ex: "A1:C5")
        c1, r1, c2, r2 = m.groups()
        r1, c1 = a1_to_rc(f"{c1}{r1}")
        r2, c2 = a1_to_rc(f"{c2}{r2}")
        # S'assure que le coin supérieur gauche est bien le point de départ
        if r2 < r1 or c2 < c1:
            r1, r2 = min(r1, r2), max(r1, r2)
            c1, c2 = min(c1, c2), max(c1, c2)
        return r1, c1, r2, c2
    else:
        # C'est une seule cellule (ex: "B2")
        r, c = a1_to_rc(rng)
        return r, c, r, c


def get_sheet(doc, sheet_name: str):
    """
    Récupère un objet feuille de calcul à partir de son nom dans un document ODS.

    Args:
        doc: L'objet document `ezodf` ouvert.
        sheet_name (str): Le nom de la feuille à trouver.

    Returns:
        ezodf.Sheet: L'objet feuille de calcul correspondant.

    Raises:
        ValueError: Si la feuille n'est pas trouvée.
    """
    for sh in doc.sheets:
        if sh.name == sheet_name: return sh
    raise ValueError(f"Feuille introuvable: {sheet_name}")


def ensure_size(sheet, r, c):
    """
    Vérifie si une feuille est assez grande pour contenir une cellule aux coordonnées (r, c).
    Si non, ajoute les lignes et/ou colonnes nécessaires.

    Args:
        sheet (ezodf.Sheet): La feuille de calcul à vérifier/modifier.
        r (int): L'index de la ligne (base zéro).
        c (int): L'index de la colonne (base zéro).
    """
    nrows, ncols = sheet.nrows(), sheet.ncols()
    if r >= nrows: sheet.append_rows(r - nrows + 1)
    if c >= ncols: sheet.append_columns(c - ncols + 1)


def set_cell_value(sheet, a1: str, value, typ: str):
    """
    Définit la valeur d'une cellule en utilisant son adresse A1.
    La fonction gère la conversion de type (string, int, float).

    Args:
        sheet (ezodf.Sheet): La feuille de calcul à modifier.
        a1 (str): L'adresse de la cellule (ex: "A1").
        value: La valeur à insérer.
        typ (str): Le type de la valeur ("string", "int", "float").
    """
    r, c = a1_to_rc(a1)
    ensure_size(sheet, r, c)
    cell = sheet[r, c]
    # Applique la valeur avec le bon type
    if typ == "string":
        cell.set_value(str(value))
    elif typ in ("int", "integer"):
        cell.set_value(int(value))
    elif typ in ("float", "number"):
        cell.set_value(float(value))
    else:
        cell.set_value(str(value))


def insert_rows(sheet, at: int, count: int):
    """
    Insère un nombre de lignes donné à une position spécifique dans la feuille.

    Args:
        sheet (ezodf.Sheet): La feuille de calcul à modifier.
        at (int): Le numéro de la ligne (base 1) avant laquelle insérer.
        count (int): Le nombre de lignes à insérer.
    
    Raises:
        ValueError: Si l'index de ligne est inférieur à 1.
    """
    if at < 1: raise ValueError("L'index de ligne doit être >= 1")
    # `ezodf` utilise un index base zéro, d'où le `at - 1`
    sheet.insert_rows(at - 1, count)


def sanitize_filename(name: str) -> str:
    """
    Nettoie une chaîne de caractères pour la rendre valide en tant que nom de fichier
    en remplaçant les caractères invalides par des tirets.

    Args:
        name (str): Le nom de fichier potentiel.

    Returns:
        str: Le nom de fichier nettoyé.
    """
    return ''.join('-' if ch in INVALID_FS_CHARS else ch for ch in name).strip()


def bump_version_in_stem(stem: str) -> str:
    """
    Recherche un numéro de version (ex: "v1.2.3") dans une chaîne et l'incrémente.
    Si "v1.2.3" est trouvé, il devient "v1.2.4". Si aucun numéro de version n'est
    trouvé, la chaîne originale est retournée.

    Args:
        stem (str): La racine du nom de fichier (sans extension).

    Returns:
        str: La racine du nom de fichier avec le numéro de version incrémenté.
    """
    # Fonction interne pour gérer l'incrémentation
    def _bump(ver):
        parts = ver.split('.')
        try:
            # Incrémente la dernière partie du numéro de version
            parts[-1] = str(int(parts[-1]) + 1)
        except (ValueError, IndexError):
            # En cas d'échec (ex: "1.0a"), retourne la version telle quelle
            return ver
        return '.'.join(parts)

    m = VER_RE.match(stem)
    if not m: return stem  # Pas de version trouvée
    pre, vchar, ver, post = m.groups()
    return f"{pre}{vchar}{_bump(ver)}{post}"


def render_out_name(src_path: Path, name_pattern: str, suffix_tpl: str, parenthesis_replace: str = None,
                    bump_version: bool = False):
    """
    Génère le nom du fichier de sortie en se basant sur un modèle et plusieurs options.

    Args:
        src_path (Path): Le chemin du fichier source.
        name_pattern (str): Le modèle pour le nom de fichier (ex: "${stem}_modifié").
        suffix_tpl (str): Un modèle pour un suffixe à ajouter (ex: "_${date}").
        parenthesis_replace (str, optional): Texte pour remplacer tout contenu entre parenthèses.
        bump_version (bool, optional): Si True, incrémente la version dans le nom.

    Returns:
        str: Le nom de fichier final, nettoyé et prêt à être utilisé.
    """
    # Crée une chaîne de caractères de la date et l'heure actuelle (ex: "20231027-153000")
    dtstr = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    suffix = (suffix_tpl or "").replace("${date}", dtstr)
    
    stem = src_path.stem  # Racine du nom de fichier (ex: "document" pour "document.ods")
    
    # Applique les transformations optionnelles
    if parenthesis_replace: stem = re.sub(r'\([^)]*\)', f'({parenthesis_replace})', stem)
    if bump_version: stem = bump_version_in_stem(stem)
    
    # Remplace les placeholders dans le modèle de nom
    out = name_pattern.replace("${stem}", stem).replace("${ext}", src_path.suffix).replace("${date}", dtstr).replace(
        "${suffix}", suffix)
    
    # S'assure que le fichier a bien une extension .ods si non spécifiée
    if "${ext}" not in name_pattern and not out.lower().endswith(".ods"): out += ".ods"
    
    return sanitize_filename(out)


# --- ODF STYLE ENGINE ---
# Le style dans les fichiers ODS est géré via l'API odfpy, qui permet de créer
# et manipuler les styles directement en Python sans passer par du XML brut.

def _fix_ods_zip(ods_path: str):
    """
    Corrige le fichier ODS après une sauvegarde odfpy en supprimant les entrées
    'mimetype' dupliquées. odfpy écrit mimetype une première fois explicitement
    (non compressé), puis une seconde fois depuis le fichier source (compressé),
    ce que LibreOffice interprète comme une corruption.

    Args:
        ods_path (str): Chemin du fichier ODS à corriger.
    """
    tmp = ods_path + '.tmp'
    try:
        with zipfile.ZipFile(ods_path, 'r') as zin, \
             zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            mimetype_written = False
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'mimetype':
                    if mimetype_written:
                        continue  # Ignore les doublons
                    # Écrit mimetype en premier, non compressé (spec ODF)
                    zi = zipfile.ZipInfo('mimetype')
                    zi.compress_type = zipfile.ZIP_STORED
                    zout.writestr(zi, data)
                    mimetype_written = True
                else:
                    zout.writestr(item, data)
        shutil.move(tmp, ods_path)
    except Exception as e:
        Path(tmp).unlink(missing_ok=True)
        print(f"Err fix zip: {e}")

def _make_odf_style(doc, style_name: str, cfg: dict):
    """
    Fonction interne pour créer un élément de style odfpy à partir d'un dictionnaire
    de propriétés et l'ajouter aux styles automatiques du document.

    Args:
        doc: Le document odfpy ouvert.
        style_name (str): Le nom du style à créer.
        cfg (dict): Dictionnaire de propriétés (ex: {'background': '#FF0000', 'bold': True}).
    """
    style = Style(name=style_name, family="table-cell")

    # Propriétés de la cellule (Fond, Alignement Vertical, Bordure)
    if cfg.get('background') or cfg.get('valign'):
        tcp_attrs = {}
        if cfg.get('background'):
            tcp_attrs['backgroundcolor'] = cfg['background']
            tcp_attrs['border'] = '0.05pt solid #000000'  # Bordure fine noire par défaut
        if cfg.get('valign'):
            vmap = {'top': 'top', 'middle': 'middle', 'bottom': 'bottom'}
            if cfg['valign'] in vmap:
                tcp_attrs['verticalalign'] = vmap[cfg['valign']]
        if cfg.get('wrap'):
            tcp_attrs['wrapoption'] = 'wrap'
        style.addElement(TableCellProperties(**tcp_attrs))

    # Propriétés de paragraphe (Alignement H)
    if cfg.get('halign'):
        style.addElement(ParagraphProperties(textalign=cfg['halign']))

    # Propriétés de texte (Gras, Taille, Couleur)
    if cfg.get('bold') or cfg.get('font_size'):
        tp_attrs = {'color': '#000000'}  # Texte Noir
        if cfg.get('bold'):
            tp_attrs['fontweight'] = 'bold'
            tp_attrs['fontweightasian'] = 'bold'
            tp_attrs['fontweightcomplex'] = 'bold'
        if cfg.get('font_size'):
            fsize = f"{cfg['font_size']}pt"
            tp_attrs['fontsize'] = fsize
            tp_attrs['fontsizeasian'] = fsize
            tp_attrs['fontsizecomplex'] = fsize
        style.addElement(TextProperties(**tp_attrs))

    doc.automaticstyles.addElement(style)


def apply_styles_via_odf(ods_path: str, style_defs: dict):
    """
    Applique des styles de cellule à un fichier ODS via l'API odfpy.

    Args:
        ods_path (str): Le chemin vers le fichier .ods à modifier.
        style_defs (dict): Le dictionnaire de définitions de style à appliquer.
    """
    if not style_defs: return
    try:
        doc = odf_load(ods_path)
        for style_name, cfg in style_defs.items():
            _make_odf_style(doc, style_name, cfg)
        doc.save(ods_path)
        _fix_ods_zip(ods_path)
    except Exception as e:
        print(f"Err ODF Style: {e}")


def restore_colors_preserve_formatting_odf(ods_path, sheet_name, start_row, end_row, column_colors,
                                           exclude_rows=None, protected_prefix=None):
    """
    Fonction clé pour réinitialiser les couleurs de fond d'une plage de cellules
    tout en préservant le formatage existant (gras, bordures, etc.).

    Cette fonction est essentielle car elle permet de "nettoyer" un fichier avant
    d'appliquer de nouvelles modifications, sans pour autant perdre le formatage
    précédent. Elle peut également protéger les styles ajoutés lors de la même
    exécution grâce à un préfixe unique.

    Args:
        ods_path (str): Chemin du fichier ODS.
        sheet_name (str): Nom de la feuille à traiter.
        start_row (int): Ligne de début (base zéro) pour la recoloration.
        end_row (int or None): Ligne de fin (base zéro). Si None, va jusqu'à la fin.
        column_colors (dict): Dictionnaire mappant un index de colonne (int) à une
                              couleur de fond (ex: {0: '#C0C0C0', 1: '#CCFFFF'}).
        exclude_rows (list[int], optional): Liste d'index de lignes (base zéro) à ignorer.
        protected_prefix (str, optional): Un préfixe de style (ex: "Run_123_"). Tout style
                                          dont le nom commence par ce préfixe sera ignoré
                                          par le nettoyage, protégeant ainsi les modifications
                                          faites dans la session courante.
    """
    if exclude_rows is None: exclude_rows = []
    try:
        doc = odf_load(ods_path)

        # Construit un dictionnaire nom_style -> élément Style pour recherche rapide
        existing_styles = {}
        for s in doc.automaticstyles.childNodes:
            sname = s.attributes.get((STYLENS, 'name'))
            if sname:
                existing_styles[sname] = s

        # Recherche de la feuille cible par son nom
        target_sheet = None
        for sheet in doc.spreadsheet.childNodes:
            if sheet.attributes.get((TABLENS, 'name')) == sheet_name:
                target_sheet = sheet
                break
        if not target_sheet: return

        style_mapping = {}  # Cache (style_original, nouvelle_couleur) -> nouveau_nom
        new_style_counter = [1]  # liste pour mutabilité dans la closure

        def get_or_create_restore_style(current_style_name, tcolor):
            key = (current_style_name or "Default", tcolor)
            if key in style_mapping:
                return style_mapping[key]

            n_name = f"RestoreColor{new_style_counter[0]}"
            new_style_counter[0] += 1
            new_style = Style(name=n_name, family="table-cell")

            # Copie les propriétés de l'ancien style en remplaçant la couleur de fond.
            # Les éléments chargés depuis le fichier sont des instances génériques de Element,
            # on crée donc une sous-classe dynamique avec le bon qname pour les recréer proprement.
            tcp_created = False
            if current_style_name and current_style_name in existing_styles:
                ex_style = existing_styles[current_style_name]
                for ch in ex_style.childNodes:
                    TmpCls = type('_OdfElem', (ODFElement,), {'qname': ch.qname})
                    new_elem = TmpCls()
                    new_elem.attributes.update(ch.attributes)
                    if ch.qname == TableCellProperties().qname:
                        new_elem.attributes[(FONS, 'background-color')] = tcolor
                        tcp_created = True
                    new_style.addElement(new_elem)

            # Si aucune propriété cellule n'existait, on en crée une nouvelle
            if not tcp_created:
                new_style.addElement(TableCellProperties(
                    backgroundcolor=tcolor,
                    border='0.05pt solid #000000'
                ))

            doc.automaticstyles.addElement(new_style)
            existing_styles[n_name] = new_style
            style_mapping[key] = n_name
            return n_name

        # Parcours des lignes de la feuille cible
        r_idx = 0
        for row in target_sheet.childNodes:
            if row.qname != (TABLENS, 'table-row'):
                continue
            if r_idx < start_row:
                r_idx += 1
                continue
            if end_row is not None and r_idx > end_row:
                break
            if r_idx in exclude_rows:
                r_idx += 1
                continue

            curr_col = 0
            for cell in row.childNodes:
                rep = int(cell.attributes.get((TABLENS, 'number-columns-repeated')) or 1)

                if cell.qname == (TABLENS, 'covered-table-cell'):
                    curr_col += rep
                    continue

                if cell.qname == (TABLENS, 'table-cell'):
                    current_style_name = cell.attributes.get((TABLENS, 'style-name'))

                    # Protection des styles de la session courante
                    if protected_prefix and current_style_name and current_style_name.startswith(protected_prefix):
                        curr_col += rep
                        continue

                    tcolor = column_colors.get(curr_col)
                    if tcolor:
                        n_name = get_or_create_restore_style(current_style_name, tcolor)
                        cell.attributes[(TABLENS, 'style-name')] = n_name

                curr_col += rep
            r_idx += 1

        doc.save(ods_path)
        _fix_ods_zip(ods_path)
    except Exception as e:
        print(f"Err Restore Color: {e}")


def process_file(src_path: Path, out_dir: Path, suffix_tpl: str, name_pattern: str, ops: list, options: dict):
    """
    Fonction principale du backend qui traite un seul fichier ODS.

    Elle ouvre le document, applique une série d'opérations (définies dans `ops`),
    gère la réinitialisation des couleurs, génère le nom du fichier de sortie,
    sauvegarde le fichier modifié, puis applique les nouveaux styles XML.

    Args:
        src_path (Path): Chemin du fichier ODS source.
        out_dir (Path): Répertoire où sauvegarder le fichier modifié.
        suffix_tpl (str): Modèle pour le suffixe du nom de fichier.
        name_pattern (str): Modèle principal pour le nom de fichier.
        ops (list[dict]): Liste des opérations à effectuer. Chaque opération est un
                          dictionnaire (ex: {'op': 'set_value', 'sheet': 'Feuille1', ...}).
        options (dict): Dictionnaire d'options globales (ex: 'bump_version', 'reset_colors_before').

    Returns:
        str: Un message indiquant le succès ou l'échec de l'opération, incluant
             le nom du fichier généré.
    """
    doc = ezodf.opendoc(str(src_path))
    style_cache, style_defs = {}, {}  # Cache et définitions pour les nouveaux styles
    ct_restore = None  # Paramètres pour la restauration des couleurs

    # 1. Génère un préfixe unique pour cette session. Tous les styles créés
    #    pendant cette exécution commenceront par ce préfixe (ex: "Run_167..._").
    #    Cela permet à `restore_colors_preserve_formatting_odf` de ne pas effacer
    #    les couleurs que l'on vient juste d'ajouter.
    session_prefix = f"Run_{int(dt.datetime.now().timestamp())}_"

    # 2. Prépare la configuration pour la réinitialisation des couleurs si l'option est activée.
    if options.get("reset_colors_before", False):
        col_cols = options.get("column_colors") or {0: "#C0C0C0", 1: "#CCFFFF", 2: "#CCFFFF", 3: "#CCFFCC",
                                                    4: "#CCFFCC", 5: "#FFFFCC", 6: "#FFFFCC", 7: "#FFFFCC",
                                                    9: "#FFFFCC", 10: "#FFFFCC"}
        ct_restore = {
            'sheet_names': options.get("reset_colors_sheets", None),
            'column_colors': col_cols,
            'start_row': options.get("reset_start_row", 0),
            'end_row': options.get("reset_end_row", None),
            'exclude_rows': options.get("exclude_rows", [])
        }

    # 3. Boucle principale : applique chaque opération séquentiellement.
    for op in ops:
        kind = op.get("op")

        if kind == "copy_range":
            src_sheet = get_sheet(doc, op["src_sheet"])
            dst_sheet = get_sheet(doc, op.get("dst_sheet") or op["src_sheet"])
            r1, c1, r2, c2 = parse_range(op["src_range"])
            dr, dc = a1_to_rc(op["dst_tl"])
            h, w = r2 - r1 + 1, c2 - c1 + 1
            for i in range(h):
                for j in range(w):
                    sr, sc = r1 + i, c1 + j
                    val = src_sheet[sr, sc].value if sr < src_sheet.nrows() and sc < src_sheet.ncols() else ""
                    tr, tc = (dr + j, dc + i) if op.get("transpose") else (dr + i, dc + j)
                    ensure_size(dst_sheet, tr, tc)
                    dst_sheet[tr, tc].set_value("" if val is None else val)
            continue

        sheet_name = op.get("sheet")
        if not sheet_name: continue
        try:
            sheet = get_sheet(doc, sheet_name)
        except ValueError:
            continue

        if kind == "insert_rows":
            insert_rows(sheet, int(op["at"]), int(op.get("count", 1)))
            # -- LOGIQUE DE COLORATION RETABLIE ICI --
            bg_color = op.get("background")
            if bg_color:
                key = (False, bg_color, None, None, False, None)
                if key not in style_cache:
                    # On utilise le session_prefix pour que le nettoyeur reconnaisse ce style comme "A GARDER"
                    sn = f"{session_prefix}{len(style_defs) + 1}"
                    style_cache[key] = sn
                    style_defs[sn] = {"bold": False, "background": bg_color, "halign": None, "valign": None,
                                      "wrap": False, "font_size": None}
                else:
                    sn = style_cache[key]

                # On applique la couleur sur toute la ligne insérée (Colonnes A à Z pour être sûr)
                actual_row_idx = int(op["at"]) - 1
                for r in range(actual_row_idx, actual_row_idx + int(op.get("count", 1))):
                    for c in range(11):  # On force sur 11 colonnes (A-Z)
                        ensure_size(sheet, r, c)
                        sheet[r, c].style_name = sn

        elif kind == "merge_cells":
            r1, c1, r2, c2 = parse_range(op["range"])
            ensure_size(sheet, r2, c2)
            sheet[r1, c1].span = (r2 - r1 + 1, c2 - c1 + 1)
        elif kind == "set_value":
            set_cell_value(sheet, op["cell"], op.get("value", ""), op.get("type", "string"))
        elif kind == "fill_range":
            r1, c1, r2, c2 = parse_range(op["range"])
            typ, val = op.get("type", "string"), op.get("value", "")
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    ensure_size(sheet, r, c)
                    sheet[r, c].set_value(int(val) if typ == "int" else float(val) if typ == "float" else str(val))
        elif kind == "paste_grid":
            r0, c0 = a1_to_rc(op["start"])
            infer = bool(op.get("infer", True))
            for i, row in enumerate(op["grid"]):
                for j, raw in enumerate(row):
                    r, c = r0 + i, c0 + j
                    ensure_size(sheet, r, c)
                    val = str(raw).strip()
                    if infer:
                        try:
                            if val.isdigit() or (val.startswith("-") and val[1:].isdigit()):
                                sheet[r, c].set_value(int(val))
                            else:
                                sheet[r, c].set_value(float(val))
                        except:
                            sheet[r, c].set_value(val)
                    else:
                        sheet[r, c].set_value(val)
        elif kind == "clear_range":
            r1, c1, r2, c2 = parse_range(op["range"])
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    ensure_size(sheet, r, c)
                    sheet[r, c].set_value("")

        elif kind == "style_cell":
            cells = op.get("cells", [])
            fsize = op.get("font_size")
            key = (bool(op.get("bold")), op.get("background") or "", op.get("halign") or "", op.get("valign") or "",
                   bool(op.get("wrap")), fsize or "")
            if key not in style_cache:
                sn = f"{session_prefix}{len(style_defs) + 1}"
                style_cache[key] = sn
                style_defs[sn] = {
                    "bold": key[0], "background": key[1] or None,
                    "halign": key[2] or None, "valign": key[3] or None,
                    "wrap": key[4], "font_size": fsize
                }
            else:
                sn = style_cache[key]
            for a1 in cells:
                r, c = a1_to_rc(a1)
                ensure_size(sheet, r, c)
                sheet[r, c].style_name = sn

    # 4. Génère le nom du fichier de sortie et sauvegarde le document initialement.
    #    Les manipulations XML (couleurs, styles) se feront sur cette copie sauvegardée.
    out_name = render_out_name(src_path, name_pattern, suffix_tpl, options.get("parenthesis_replace"),
                               options.get("bump_version", False))
    out_path = out_dir / out_name
    if options.get("dry_run"): return f"[SIMULATION] {out_path.name}"
    doc.saveas(str(out_path))

    # 5. Étape de post-traitement XML.
    # Applique la restauration des couleurs (si activée).
    # Le `protected_prefix` garantit que les styles ajoutés à l'étape 3 ne sont pas effacés.
    if ct_restore:
        sn_list = ct_restore['sheet_names']
        target_sheets = sn_list if sn_list else [s.name for s in doc.sheets]
        for sname in target_sheets:
            restore_colors_preserve_formatting_odf(
                str(out_path),
                sname,
                ct_restore['start_row'],
                ct_restore['end_row'],
                ct_restore['column_colors'],
                ct_restore['exclude_rows'],
                protected_prefix=session_prefix
            )

    # Applique tous les nouveaux styles (couleurs, gras, etc.) qui ont été définis.
    if style_defs: apply_styles_via_odf(str(out_path), style_defs)
    
    return f"Succès: {out_path.name}"


# ==============================================================================
# 2. INTERFACE GRAPHIQUE (GUI)
# ==============================================================================
# Cette section définit l'interface utilisateur de l'application en utilisant PySide6.

class ColorConfigDialog(QDialog):
    """
    Boîte de dialogue modale pour configurer les couleurs de fond à appliquer
    lors de la réinitialisation des couleurs.
    """
    def __init__(self, current_text, parent=None):
        """
        Initialise la boîte de dialogue.

        Args:
            current_text (str): Le texte de configuration actuel à afficher.
            parent (QWidget, optional): Le widget parent.
        """
        super().__init__(parent)
        self.setWindowTitle("Configuration des Couleurs")
        self.resize(400, 500)
        layout = QVBoxLayout(self)
        lbl = QLabel(
            "Format: Colonne=Couleur (une par ligne)\nEx: A=#C0C0C0\nLaissez vide après = pour ignorer la colonne.")
        layout.addWidget(lbl)
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(current_text)
        layout.addWidget(self.text_edit)
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def get_text(self):
        """
        Retourne le texte contenu dans la zone d'édition.

        Returns:
            str: La configuration des couleurs entrée par l'utilisateur.
        """
        return self.text_edit.toPlainText()


class ModernODSApp(QMainWindow):
    """
    Fenêtre principale de l'application.
    Structure et gère tous les widgets de l'interface graphique et connecte
    les actions de l'utilisateur à la logique métier (backend).
    """
    def __init__(self):
        """
        Initialise la fenêtre principale, configure ses propriétés,
        et appelle les méthodes pour construire les différentes sections de l'UI.
        """
        super().__init__()
        self.setWindowTitle("Assistant Matrices ODS - Ultimate V9")
        self.resize(1300, 850)

        # --- Initialisation de l'état de l'application ---
        self.files = []  # Liste des chemins des fichiers à traiter
        self.output_dir = str(Path.cwd() / "Resultats")  # Dossier de sortie par défaut
        self._grids = {} # (Non utilisé actuellement, peut être retiré)
        # Configuration par défaut des couleurs pour la réinitialisation
        self.custom_colors_txt = "\n".join([
            "A=#C0C0C0", "B=#CCFFFF", "C=#CCFFFF", "D=#CCFFCC", "E=#CCFFCC",
            "F=#FFFFCC", "G=#FFFFCC", "H=#FFFFCC", "I=#FFFFFF", "J=#FFFFCC", "K=#FFFFCC"
        ])

        # --- Construction de l'interface ---
        self.setup_menu_bar()
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout(central_widget)

        self.setup_file_section()
        self.setup_options_section()
        self.setup_tabs_section()
        self.setup_run_section()

    def setup_menu_bar(self):
        """Crée la barre de menu supérieure avec les options Fichier > Sauvegarder/Charger."""
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Fichier")
        act_save = QAction("Sauvegarder Configuration...", self)
        act_save.triggered.connect(self.save_profile)
        act_load = QAction("Charger Configuration...", self)
        act_load.triggered.connect(self.load_profile)
        file_menu.addAction(act_save)
        file_menu.addAction(act_load)

    def setup_file_section(self):
        """Crée la section de l'interface pour la sélection des fichiers ODS."""
        group = QGroupBox("1. Sélection des Fichiers")
        layout = QHBoxLayout()
        self.file_list = QListWidget()
        layout.addWidget(self.file_list)
        btn_layout = QVBoxLayout()
        self.btn_add_files = QPushButton("Ajouter Fichiers")
        self.btn_add_folder = QPushButton("Ajouter Dossier")
        self.btn_clear = QPushButton("Tout Vider")
        self.btn_clear.setProperty('class', 'danger')
        self.btn_add_files.clicked.connect(self.add_files)
        self.btn_add_folder.clicked.connect(self.add_folder)
        self.btn_clear.clicked.connect(self.clear_files)
        btn_layout.addWidget(self.btn_add_files);
        btn_layout.addWidget(self.btn_add_folder)
        btn_layout.addStretch();
        btn_layout.addWidget(self.btn_clear)
        layout.addLayout(btn_layout)
        group.setLayout(layout)
        self.main_layout.addWidget(group)

    def setup_options_section(self):
        """Crée la section des options globales (dossier de sortie, versioning, couleurs)."""
        group = QGroupBox("2. Configuration Globale")
        layout = QHBoxLayout()
        
        # --- Colonne de Gauche : Options de sortie ---
        l_left = QVBoxLayout()
        h1 = QHBoxLayout()
        self.chk_src_dir = QCheckBox("Sauvegarder dans le dossier source")
        self.chk_src_dir.toggled.connect(self.toggle_out_ui)
        self.btn_out = QPushButton("Choisir autre dossier...")
        self.btn_out.clicked.connect(self.choose_out_dir)
        self.lbl_out = QLabel(f"Vers: {self.output_dir}")
        self.lbl_out.setStyleSheet("font-size: 10px; color: gray;")
        h1.addWidget(self.chk_src_dir);
        h1.addWidget(self.btn_out)

        h2 = QHBoxLayout()
        self.chk_version = QCheckBox("Incrémenter Version (v1.0 -> v1.1)")
        self.chk_version.setChecked(True)
        self.input_paren = QLineEdit();
        self.input_paren.setPlaceholderText("Remplacer texte entre ( )")
        h2.addWidget(self.chk_version);
        h2.addWidget(self.input_paren)
        l_left.addLayout(h1);
        l_left.addWidget(self.lbl_out);
        l_left.addLayout(h2)

        # --- Colonne de Droite : Options de réinitialisation des couleurs ---
        l_right = QVBoxLayout()
        h3 = QHBoxLayout()
        self.chk_reset = QCheckBox("Réinitialiser Couleurs")
        self.chk_reset.setChecked(True)
        self.btn_colors = QPushButton("Config Couleurs...")
        self.btn_colors.clicked.connect(self.config_colors_dialog)
        h3.addWidget(self.chk_reset);
        h3.addWidget(self.btn_colors)
        
        h4 = QHBoxLayout()
        self.input_reset_sheets = QLineEdit();
        self.input_reset_sheets.setPlaceholderText("Feuilles (vide=toutes)")
        self.input_reset_start = QLineEdit("10");
        self.input_reset_start.setPlaceholderText("Ligne Début")
        self.input_reset_excl = QLineEdit();
        self.input_reset_excl.setPlaceholderText("Exclure (ex: 15,20)")
        h4.addWidget(QLabel("Cible:"));
        h4.addWidget(self.input_reset_sheets)
        h4.addWidget(QLabel("Start:"));
        h4.addWidget(self.input_reset_start)
        h4.addWidget(QLabel("Excl:"));
        h4.addWidget(self.input_reset_excl)
        l_right.addLayout(h3);
        l_right.addLayout(h4)

        layout.addLayout(l_left);
        layout.addStretch();
        layout.addLayout(l_right)
        group.setLayout(layout)
        self.main_layout.addWidget(group)

    def setup_tabs_section(self):
        """Crée le widget à onglets et initialise chaque onglet d'opérations."""
        self.tabs = QTabWidget()
        self.tab_content = QWidget();
        self.tab_struct = QWidget();
        self.tab_style = QWidget();
        self.tab_copy = QWidget()
        self.tabs.addTab(self.tab_content, "Édition Contenu");
        self.tabs.addTab(self.tab_struct, "Structure")
        self.tabs.addTab(self.tab_style, "Mise en Forme");
        self.tabs.addTab(self.tab_copy, "Copie / Transfert")
        
        # Appelle les méthodes de construction pour chaque onglet
        self._setup_content_tab();
        self._setup_struct_tab();
        self._setup_style_tab();
        self._setup_copy_tab()
        self.main_layout.addWidget(self.tabs)

    def delete_row(self, tree_widget):
        """Supprime la ou les lignes sélectionnées d'un QTreeWidget."""
        root = tree_widget.invisibleRootItem()
        for item in tree_widget.selectedItems():
            root.removeChild(item)

    def _setup_content_tab(self):
        """Construit l'interface de l'onglet 'Édition Contenu'."""
        layout = QVBoxLayout(self.tab_content)
        form = QHBoxLayout()
        self.c_sheet = QLineEdit();
        self.c_sheet.setPlaceholderText("Feuille")
        self.c_range = QLineEdit();
        self.c_range.setPlaceholderText("Cellule (A1) ou Plage")
        self.c_type = QComboBox();
        self.c_type.addItems(["string", "int", "float"])
        self.c_val = QLineEdit();
        self.c_val.setPlaceholderText("Valeur")
        btn_add = QPushButton("Ajouter Valeur");
        btn_add.clicked.connect(self.add_content_val)
        form.addWidget(self.c_sheet);
        form.addWidget(self.c_range);
        form.addWidget(self.c_type);
        form.addWidget(self.c_val);
        form.addWidget(btn_add)
        grid_lay = QHBoxLayout()
        self.c_grid_txt = QTextEdit();
        self.c_grid_txt.setPlaceholderText("Collez ici votre grille Excel (TAB ou ;)...")
        self.c_grid_txt.setMaximumHeight(60)
        btn_grid = QPushButton("Ajouter Grille (depuis A1)");
        btn_grid.clicked.connect(self.add_content_grid)
        grid_lay.addWidget(self.c_grid_txt);
        grid_lay.addWidget(btn_grid)
        self.tree_cont = QTreeWidget();
        self.tree_cont.setHeaderLabels(["Type", "Feuille", "Cible", "Détail"])
        layout.addLayout(form);
        layout.addLayout(grid_lay);
        layout.addWidget(self.tree_cont)
        btn_box = QHBoxLayout()
        btn_edit = QPushButton("Modifier la ligne")
        btn_edit.setProperty('class', 'warning')
        btn_edit.clicked.connect(self.edit_content_row)

        btn_del = QPushButton("Supprimer la ligne")
        btn_del.setProperty('class', 'danger')
        btn_del.clicked.connect(lambda: self.delete_row(self.tree_cont))

        btn_box.addWidget(btn_edit)
        btn_box.addWidget(btn_del)
        layout.addLayout(btn_box)

    def _setup_struct_tab(self):
        """Construit l'interface de l'onglet 'Structure' (insertion, fusion)."""
        layout = QVBoxLayout(self.tab_struct)

        # Aide visuelle
        lbl_help = QLabel(
            "Pour insérer : Remplissez 'Qté' et 'Cible'. Pour Fusionner/Effacer : Remplissez seulement 'Cible'.")
        lbl_help.setStyleSheet("color: #AAA; font-style: italic; font-size: 11px;")
        layout.addWidget(lbl_help)

        form = QHBoxLayout()

        # 1. Action
        self.s_action = QComboBox()
        self.s_action.addItems(["Insérer Lignes", "Fusionner", "Effacer"])
        self.s_action.setMinimumWidth(120)

        # 2. Feuille
        self.s_sheet = QLineEdit()
        self.s_sheet.setPlaceholderText("Feuille")

        # 3. Qté et Cible
        self.s_count = QLineEdit()
        self.s_count.setPlaceholderText("Qté")
        self.s_count.setMaximumWidth(50)

        self.s_target = QLineEdit()
        self.s_target.setPlaceholderText("La ligne ou vous allez inserer vos lignes AVANT")

        # 4. COULEUR (Menu Déroulant)
        self.s_color = QComboBox()
        self.s_color.setPlaceholderText("Couleur...")
        # On ajoute : Le Nom affiché, Le Code Hexa caché
        self.s_color.addItem("Aucune", "")
        self.s_color.addItem("Rouge", "#FF0000")
        self.s_color.addItem("Vert", "#00FF00")
        self.s_color.addItem("Bleu", "#0000FF")
        self.s_color.addItem("Jaune", "#FFFF00")
        self.s_color.addItem("Noir", "#000000")
        self.s_color.addItem("Blanc", "#FFFFFF")
        self.s_color.addItem("Gris", "#808080")
        self.s_color.addItem("Orange", "#FFA500")
        self.s_color.setMinimumWidth(100)

        btn = QPushButton("Ajouter")
        btn.clicked.connect(self.add_struct_action)

        form.addWidget(self.s_action)
        form.addWidget(self.s_sheet)
        form.addWidget(self.s_count)
        form.addWidget(self.s_target)
        form.addWidget(self.s_color)  # Ajout du menu
        form.addWidget(btn)

        self.tree_struct = QTreeWidget()
        self.tree_struct.setHeaderLabels(["Action", "Feuille", "Détail (Généré)", "Couleur"])

        layout.addLayout(form)
        layout.addWidget(self.tree_struct)

        btn_box = QHBoxLayout()

        btn_edit = QPushButton("Modifier la ligne")
        btn_edit.setProperty('class', 'warning')  # Orange
        btn_edit.clicked.connect(self.edit_struct_row)  # <-- Connexion nouvelle fonction

        btn_del = QPushButton("Supprimer la ligne")
        btn_del.setProperty('class', 'danger')  # Rouge
        btn_del.clicked.connect(lambda: self.delete_row(self.tree_struct))

        btn_box.addWidget(btn_edit)
        btn_box.addWidget(btn_del)
        layout.addLayout(btn_box)

    def _setup_style_tab(self):
        """Construit l'interface de l'onglet 'Mise en Forme'."""
        layout = QVBoxLayout(self.tab_style)
        r1 = QHBoxLayout()
        self.y_sheet = QLineEdit();
        self.y_sheet.setPlaceholderText("Feuille")
        self.y_cells = QLineEdit();
        self.y_cells.setPlaceholderText("Cellules (A1, B2...)")
        self.y_bold = QCheckBox("Gras");
        self.y_wrap = QCheckBox("Retour a ligne auto")
        self.y_size = QLineEdit();
        self.y_size.setPlaceholderText("Taille (pt)")
        self.y_size.setMaximumWidth(80)
        r1.addWidget(self.y_sheet);
        r1.addWidget(self.y_cells);
        r1.addWidget(self.y_bold);
        r1.addWidget(self.y_wrap)
        r1.addWidget(self.y_size)

        r2 = QHBoxLayout()

        # REMPLACEMENT DU CHAMP FOND PAR UN MENU
        self.y_bg = QComboBox()
        self.y_bg.addItem("Fond: Aucun", "")
        self.y_bg.addItem("Rouge", "#FF0000")
        self.y_bg.addItem("Vert", "#00FF00")
        self.y_bg.addItem("Bleu", "#0000FF")
        self.y_bg.addItem("Jaune", "#FFFF00")
        self.y_bg.addItem("Noir", "#000000")
        self.y_bg.addItem("Blanc", "#FFFFFF")
        self.y_bg.addItem("Gris", "#808080")
        self.y_bg.addItem("Orange", "#FFA500")
        self.y_bg.setMinimumWidth(120)

        self.y_halign = QComboBox();
        self.y_halign.addItems(["", "left", "center", "right"])
        self.y_valign = QComboBox();
        self.y_valign.addItems(["", "top", "middle", "bottom"])

        btn = QPushButton("Ajouter Style");
        btn.clicked.connect(self.add_style_action)

        r2.addWidget(self.y_bg)  # Ajout du menu Fond
        r2.addWidget(QLabel("H-Align:"));
        r2.addWidget(self.y_halign)
        r2.addWidget(QLabel("V-Align:"));
        r2.addWidget(self.y_valign);
        r2.addWidget(btn)

        self.tree_style = QTreeWidget();
        self.tree_style.setHeaderLabels(["Feuille", "Cellules", "Style"])
        layout.addLayout(r1);
        layout.addLayout(r2);
        layout.addWidget(self.tree_style)

        btn_box = QHBoxLayout()
        btn_edit = QPushButton("Modifier la ligne")
        btn_edit.setProperty('class', 'warning')
        btn_edit.clicked.connect(self.edit_style_row)

        btn_del = QPushButton("Supprimer la ligne")
        btn_del.setProperty('class', 'danger')
        btn_del.clicked.connect(lambda: self.delete_row(self.tree_style))

        btn_box.addWidget(btn_edit)
        btn_box.addWidget(btn_del)
        layout.addLayout(btn_box)

    def _setup_copy_tab(self):
        """Construit l'interface de l'onglet 'Copie / Transfert'."""
        layout = QVBoxLayout(self.tab_copy)
        r1 = QHBoxLayout()
        self.cp_src_s = QLineEdit();
        self.cp_src_s.setPlaceholderText("Source Feuille")
        self.cp_src_r = QLineEdit();
        self.cp_src_r.setPlaceholderText("Plage (A1:C5)")
        r1.addWidget(self.cp_src_s);
        r1.addWidget(self.cp_src_r)
        r2 = QHBoxLayout()
        self.cp_dst_s = QLineEdit();
        self.cp_dst_s.setPlaceholderText("Dest Feuille (vide=idem)")
        self.cp_dst_tl = QLineEdit();
        self.cp_dst_tl.setPlaceholderText("Dest A1")
        self.cp_trans = QCheckBox("Transposer")
        btn = QPushButton("Ajouter Copie");
        btn.clicked.connect(self.add_copy_action)
        r2.addWidget(self.cp_dst_s);
        r2.addWidget(self.cp_dst_tl);
        r2.addWidget(self.cp_trans);
        r2.addWidget(btn)
        self.tree_copy = QTreeWidget();
        self.tree_copy.setHeaderLabels(["Source", "Plage", "-> Dest", "A1", "Transp"])
        layout.addLayout(r1);
        layout.addLayout(r2);
        layout.addWidget(self.tree_copy)
        btn_box = QHBoxLayout()
        btn_edit = QPushButton("Modifier la ligne")
        btn_edit.setProperty('class', 'warning')
        btn_edit.clicked.connect(self.edit_copy_row)

        btn_del = QPushButton("Supprimer la ligne")
        btn_del.setProperty('class', 'danger')
        btn_del.clicked.connect(lambda: self.delete_row(self.tree_copy))

        btn_box.addWidget(btn_edit)
        btn_box.addWidget(btn_del)
        layout.addLayout(btn_box)

    def setup_run_section(self):
        """Crée la section finale avec la barre de progression et le bouton d'exécution."""
        layout = QVBoxLayout()
        self.progress = QProgressBar();
        self.progress.setValue(0);
        self.progress.setTextVisible(False)
        self.btn_run = QPushButton("LANCER LE TRAITEMENT")
        self.btn_run.setMinimumHeight(50);
        self.btn_run.setProperty('class', 'success')
        self.btn_run.clicked.connect(self.run_batch)
        layout.addWidget(self.progress);
        layout.addWidget(self.btn_run)
        self.main_layout.addLayout(layout)

    # --- FONCTIONS DE MODIFICATION (EDIT) ---

    def edit_struct_row(self):
        """
        Charge les données de la ligne sélectionnée dans l'onglet 'Structure'
        dans les champs du formulaire pour permettre à l'utilisateur de la modifier.
        La ligne originale est supprimée pour éviter les doublons lors du nouvel ajout.
        """
        item = self.tree_struct.currentItem()
        if not item:
            QMessageBox.warning(self, "Sélection", "Veuillez sélectionner une ligne à modifier.")
            return

        # 1. Récupération des données brutes
        act, sh, detail, col_hex = item.text(0), item.text(1), item.text(2), item.text(3)

        # 2. Remplissage du formulaire
        self.s_action.setCurrentText(act)
        self.s_sheet.setText(sh)

        # Analyse intelligente du détail ("2 avant 15" vs "A1:B2")
        if "avant" in detail and act == "Insérer Lignes":
            try:
                parts = detail.split(" avant ")
                self.s_count.setText(parts[0])
                self.s_target.setText(parts[1])
            except:
                self.s_target.setText(detail)
        else:
            self.s_target.setText(detail)
            self.s_count.clear()

        # Remise de la couleur dans le menu déroulant
        idx = self.s_color.findData(col_hex)
        if idx >= 0:
            self.s_color.setCurrentIndex(idx)
        else:
            self.s_color.setCurrentIndex(0)

        # 3. Suppression de la ligne (pour éviter le doublon quand on cliquera sur Ajouter)
        self.delete_row(self.tree_struct)

    def edit_content_row(self):
        """
        Charge les données de la ligne sélectionnée dans l'onglet 'Contenu'
        dans les champs du formulaire. Gère les cas "Valeur" simple et "Grille".
        """
        item = self.tree_cont.currentItem()
        if not item:
            QMessageBox.warning(self, "Sélection", "Veuillez sélectionner une ligne.")
            return

        typ_lbl, sh, target, detail = item.text(0), item.text(1), item.text(2), item.text(3)

        if typ_lbl == "Grille":
            # Cas spécial Grille : on recharge le texte caché
            raw_grid = item.data(0, Qt.UserRole)
            self.c_grid_txt.setPlainText(raw_grid)
            self.c_sheet.setText(sh)
            self.c_range.setText(target)
            # On vide les champs valeurs simples pour éviter la confusion
            self.c_val.clear()
        else:
            # Cas Valeur simple (ex détail: "string: Bonjour")
            self.c_sheet.setText(sh)
            self.c_range.setText(target)
            if ": " in detail:
                t_str, v_str = detail.split(": ", 1)
                self.c_type.setCurrentText(t_str)
                self.c_val.setText(v_str)

        self.delete_row(self.tree_cont)

    def edit_style_row(self):
        """
        Charge les données de la ligne sélectionnée dans l'onglet 'Mise en Forme'
        dans les champs du formulaire.
        """
        item = self.tree_style.currentItem()
        if not item: return

        sh, cells, desc = item.text(0), item.text(1), item.text(2)

        self.y_sheet.setText(sh)
        self.y_cells.setText(cells)

        # Analyse de la chaîne de description pour re-configurer l'UI
        self.y_bold.setChecked("Gras" in desc)
        self.y_wrap.setChecked("Wrap" in desc)

        m_sz = re.search(r"Sz:(\d+)", desc)
        self.y_size.setText(m_sz.group(1)) if m_sz else self.y_size.clear()

        m_bg = re.search(r"Bg:(#[0-9A-Fa-f]+)", desc)
        if m_bg:
            idx = self.y_bg.findData(m_bg.group(1))
            self.y_bg.setCurrentIndex(idx if idx >= 0 else 0)
        else:
            self.y_bg.setCurrentIndex(0)

        self.y_halign.setCurrentIndex(0)
        self.y_valign.setCurrentIndex(0)
        for val in ["left", "center", "right"]:
            if val in desc: self.y_halign.setCurrentText(val)
        for val in ["top", "middle", "bottom"]:
            if val in desc: self.y_valign.setCurrentText(val)

        self.delete_row(self.tree_style)

    def edit_copy_row(self):
        """
        Charge les données de la ligne sélectionnée dans l'onglet 'Copie / Transfert'
        dans les champs du formulaire.
        """
        item = self.tree_copy.currentItem()
        if not item: return

        ss, sr, ds, dt, trans = item.text(0), item.text(1), item.text(2), item.text(3), item.text(4)

        self.cp_src_s.setText(ss)
        self.cp_src_r.setText(sr)
        self.cp_dst_s.setText("" if ds == "idem" else ds)
        self.cp_dst_tl.setText(dt)
        self.cp_trans.setChecked(trans == "True")

        self.delete_row(self.tree_copy)

    # --- ACTIONS GUI ---
    # Ces méthodes sont directement connectées aux signaux des widgets (ex: clics de bouton).

    def add_files(self):
        """Ouvre une boîte de dialogue pour sélectionner des fichiers ODS et les ajoute à la liste."""
        files, _ = QFileDialog.getOpenFileNames(self, "Choisir fichiers", "", "ODS Files (*.ods)")
        if files:
            for f in files:
                if f not in self.files: self.files.append(f); self.file_list.addItem(Path(f).name)

    def add_folder(self):
        """Ouvre une boîte de dialogue pour sélectionner un dossier et y ajoute tous les fichiers ODS."""
        d = QFileDialog.getExistingDirectory(self, "Choisir dossier")
        if d:
            for f in glob.glob(str(Path(d) / "*.ods")):
                if f not in self.files: self.files.append(f); self.file_list.addItem(Path(f).name)

    def clear_files(self):
        """Vide la liste des fichiers à traiter."""
        self.files = []; self.file_list.clear()

    def toggle_out_ui(self, checked):
        """
        Active ou désactive le bouton de sélection du dossier de sortie en fonction de
        l'état de la case "Sauvegarder dans le dossier source".
        """
        self.btn_out.setEnabled(not checked)
        self.lbl_out.setText("Vers: [Dossier Source de chaque fichier]" if checked else f"Vers: {self.output_dir}")

    def choose_out_dir(self):
        """Ouvre une boîte de dialogue pour choisir le dossier de sortie personnalisé."""
        d = QFileDialog.getExistingDirectory(self, "Dossier de sortie")
        if d: self.output_dir = d; self.lbl_out.setText(f"Vers: {d}")

    def config_colors_dialog(self):
        """Ouvre la boîte de dialogue de configuration des couleurs."""
        dlg = ColorConfigDialog(self.custom_colors_txt, self)
        if dlg.exec(): self.custom_colors_txt = dlg.get_text()

    def add_content_val(self):
        """Ajoute une opération 'set_value' ou 'fill_range' à l'arbre de contenu."""
        s, r, t, v = self.c_sheet.text(), self.c_range.text(), self.c_type.currentText(), self.c_val.text()
        if s and r: QTreeWidgetItem(self.tree_cont, ["Valeur", s, r, f"{t}: {v}"])

    def add_content_grid(self):
        """Ajoute une opération 'paste_grid' à l'arbre de contenu."""
        s, r, txt = self.c_sheet.text(), self.c_range.text(), self.c_grid_txt.toPlainText().strip()
        if s and r and txt:
            item = QTreeWidgetItem(self.tree_cont, ["Grille", s, r, f"{len(txt.splitlines())} lignes"])
            # Stocke le texte brut de la grille dans l'item lui-même pour une utilisation ultérieure
            item.setData(0, Qt.UserRole, txt);
            self.c_grid_txt.clear()

    def add_struct_action(self):
        """Ajoute une opération de structure (insérer, fusionner, effacer) à l'arbre."""
        # 1. Récupération des données
        act = self.s_action.currentText()
        sh = self.s_sheet.text()
        target = self.s_target.text()
        count = self.s_count.text() or "1"
        col = self.s_color.currentData()  # On récupère le code hexa caché

        # 2. Validation des erreurs (On arrête si vide)
        if not sh or not target:
            QMessageBox.warning(self, "Erreur", "La feuille et la Cible/Plage sont obligatoires.")
            return

        if act == "Insérer Lignes" and not target.isdigit():
            QMessageBox.warning(self, "Erreur", "Pour insérer, la Cible doit être un numéro de ligne.")
            return

        # Affiche une boîte de dialogue de confirmation/avertissement personnalisée.
        # Ce bloc est spécifique à un besoin utilisateur et peut être modifié ou retiré.
        mon_titre = "Confirmation d'ajout"
        mon_message = f"ATTENTION AMEL SI TU INSERE UNE LIGNE PENSE A LE FAIRE SUR TOUTES LES MATRICES POUR NE PAS TOUT CASSER"

        # Affiche la boite de dialogue (Information simple)
        QMessageBox.information(self, mon_titre, mon_message)

        # ==============================================================================
        # FIN DE LA POP-UP
        # ==============================================================================

        # 3. Construction du détail
        final_detail = ""
        if act == "Insérer Lignes":
            final_detail = f"{count} avant {target}"
        else:
            final_detail = target

        # 4. Ajout dans l'arbre
        QTreeWidgetItem(self.tree_struct, [act, sh, final_detail, col])

        # 5. Nettoyage des champs
        self.s_color.setCurrentIndex(0)
        self.s_target.clear()
        self.s_count.clear()

    def add_style_action(self):
        """Ajoute une opération de style à l'arbre."""
        s, c = self.y_sheet.text(), self.y_cells.text()
        if s and c:
            desc = []
            if self.y_bold.isChecked(): desc.append("Gras")
            if self.y_size.text(): desc.append(f"Sz:{self.y_size.text()}")

            bg_code = self.y_bg.currentData()  # Récupère le code hex (ex: #FF0000)
            if bg_code: desc.append(f"Bg:{bg_code}")

            if self.y_halign.currentText(): desc.append(self.y_halign.currentText())
            if self.y_valign.currentText(): desc.append(self.y_valign.currentText())
            if self.y_wrap.isChecked(): desc.append("Wrap")

            QTreeWidgetItem(self.tree_style, [s, c, ", ".join(desc)])

    def add_copy_action(self):
        """Ajoute une opération de copie de plage à l'arbre."""
        ss, sr, ds, dt = self.cp_src_s.text(), self.cp_src_r.text(), self.cp_dst_s.text(), self.cp_dst_tl.text()
        if ss and sr and dt: QTreeWidgetItem(self.tree_copy,
                                             [ss, sr, ds if ds else "idem", dt, str(self.cp_trans.isChecked())])

    # --- SAVE / LOAD ---
    
    def save_profile(self):
        """
        Sauvegarde la configuration actuelle (options globales et listes d'actions)
        dans un fichier JSON.
        """
        path, _ = QFileDialog.getSaveFileName(self, "Sauver Config", "", "JSON (*.json)")
        if not path: return
        
        # 1. Collecte toutes les options de l'interface
        data = {
            "options": {
                "out_dir": self.output_dir, "src_dir": self.chk_src_dir.isChecked(),
                "bump": self.chk_version.isChecked(), "paren": self.input_paren.text(),
                "reset": self.chk_reset.isChecked(), "rst_sheets": self.input_reset_sheets.text(),
                "rst_start": self.input_reset_start.text(), "rst_excl": self.input_reset_excl.text(),
                "colors_txt": self.custom_colors_txt
            }, "actions": {"content": [], "struct": [], "style": [], "copy": []}
        }
        root = self.tree_cont.invisibleRootItem()
        for i in range(root.childCount()):
            it = root.child(i)
            data["actions"]["content"].append(
                {"vals": [it.text(0), it.text(1), it.text(2), it.text(3)], "grid": it.data(0, Qt.UserRole)})
        root = self.tree_struct.invisibleRootItem()
        for i in range(root.childCount()):
            it = root.child(i)
            data["actions"]["struct"].append([it.text(0), it.text(1), it.text(2), it.text(3)])
        root = self.tree_style.invisibleRootItem()
        for i in range(root.childCount()):
            it = root.child(i)
            data["actions"]["style"].append([it.text(0), it.text(1), it.text(2)])
        root = self.tree_copy.invisibleRootItem()
        for i in range(root.childCount()):
            it = root.child(i)
            data["actions"]["copy"].append([it.text(0), it.text(1), it.text(2), it.text(3), it.text(4)])
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
            QMessageBox.information(self, "Succès", "Configuration sauvegardée !")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))

    def load_profile(self):
        """
        Charge une configuration depuis un fichier JSON et restaure l'état complet
        de l'interface utilisateur.
        """
        path, _ = QFileDialog.getOpenFileName(self, "Charger Config", "", "JSON (*.json)")
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            # 1. Restaure les options globales
            opts = data.get("options", {})
            if "out_dir" in opts: self.output_dir = opts["out_dir"]; self.lbl_out.setText(f"Vers: {self.output_dir}")
            self.chk_src_dir.setChecked(opts.get("src_dir", False))
            self.chk_version.setChecked(opts.get("bump", True))
            self.input_paren.setText(opts.get("paren", ""))
            self.chk_reset.setChecked(opts.get("reset", False))
            self.input_reset_sheets.setText(opts.get("rst_sheets", ""))
            self.input_reset_start.setText(opts.get("rst_start", "10"))
            self.input_reset_excl.setText(opts.get("rst_excl", ""))
            self.custom_colors_txt = opts.get("colors_txt", self.custom_colors_txt)

            # 2. Vide les arbres actuels et les repeuple avec les données chargées
            self.tree_cont.clear();
            self.tree_struct.clear();
            self.tree_style.clear();
            self.tree_copy.clear()
            for a in data["actions"].get("content", []):
                it = QTreeWidgetItem(self.tree_cont, a["vals"])
                if a.get("grid"): it.setData(0, Qt.UserRole, a["grid"])
            for a in data["actions"].get("struct", []): QTreeWidgetItem(self.tree_struct, a)
            for a in data["actions"].get("style", []): QTreeWidgetItem(self.tree_style, a)
            for a in data["actions"].get("copy", []): QTreeWidgetItem(self.tree_copy, a)
            QMessageBox.information(self, "Succès", "Configuration chargée !")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))

    def run_batch(self):
        """
        La méthode principale qui exécute le traitement par lots.
        
        Elle effectue les étapes suivantes :
        1. Vérifie si des fichiers ont été sélectionnés.
        2. Itère sur chaque arbre d'actions (structure, contenu, etc.).
        3. Traduit les informations de chaque ligne de l'arbre en un dictionnaire d'opération
           compréhensible par la fonction `process_file`.
        4. Rassemble toutes les options globales de l'interface.
        5. Appelle `process_file` pour chaque fichier de la liste.
        6. Met à jour la barre de progression.
        """
        if not self.files:
            QMessageBox.warning(self, "Erreur", "Aucun fichier sélectionné.")
            return
        
        # `ops` contiendra la liste séquentielle de toutes les opérations à effectuer.
        ops = []

        # 1. STRUCT
        root = self.tree_struct.invisibleRootItem()
        for i in range(root.childCount()):
            it = root.child(i)
            act, sh, det, col = it.text(0), it.text(1), it.text(2), it.text(3)
            if act == "Insérer Lignes":
                m = re.search(r"(\d+).* (\d+)", det)
                qt, at = m.groups() if m else (1, 1)
                ops.append({"op": "insert_rows", "sheet": sh, "at": int(at), "count": int(qt),
                            "background": col if col else None})
            elif act == "Fusionner":
                ops.append({"op": "merge_cells", "sheet": sh, "range": det})
            elif act == "Effacer":
                ops.append({"op": "clear_range", "sheet": sh, "range": det})

        # 2. CONTENT
        root = self.tree_cont.invisibleRootItem()
        for i in range(root.childCount()):
            it = root.child(i)
            kind, sh, target, det = it.text(0), it.text(1), it.text(2), it.text(3)
            if kind == "Valeur":
                typ, val = det.split(": ", 1)
                if ":" in target:
                    ops.append({"op": "fill_range", "sheet": sh, "range": target, "type": typ, "value": val})
                else:
                    ops.append({"op": "set_value", "sheet": sh, "cell": target, "type": typ, "value": val})
            elif kind == "Grille":
                raw = it.data(0, Qt.UserRole)
                rows = [l.split('\t' if '\t' in l else ';') for l in raw.splitlines() if l.strip()]
                ops.append({"op": "paste_grid", "sheet": sh, "start": target, "grid": rows, "infer": True})

        # 3. STYLE
        root = self.tree_style.invisibleRootItem()
        for i in range(root.childCount()):
            it = root.child(i)
            sh, cells, desc = it.text(0), it.text(1), it.text(2)
            bg_match = re.search(r"Bg:(#[0-9A-Fa-f]+)", desc)
            sz_match = re.search(r"Sz:(\d+)", desc)
            ops.append({
                "op": "style_cell", "sheet": sh, "cells": [x.strip() for x in cells.split(',')],
                "bold": "Gras" in desc,
                "background": bg_match.group(1) if bg_match else None,
                "font_size": int(sz_match.group(1)) if sz_match else None,
                "halign": None, "valign": None, "wrap": "Wrap" in desc
            })

        # 4. COPY
        root = self.tree_copy.invisibleRootItem()
        for i in range(root.childCount()):
            it = root.child(i)
            ops.append({"op": "copy_range", "src_sheet": it.text(0), "src_range": it.text(1),
                        "dst_sheet": it.text(2) if it.text(2) != "idem" else None, "dst_tl": it.text(3),
                        "transpose": it.text(4) == "True"})

        opts = {
            "dry_run": False, "bump_version": self.chk_version.isChecked(),
            "parenthesis_replace": self.input_paren.text() or None,
            "reset_colors_before": self.chk_reset.isChecked(),
            "reset_start_row": int(self.input_reset_start.text()) - 1, "reset_end_row": None,
            "exclude_rows": [int(x) - 1 for x in self.input_reset_excl.text().split(',') if x.strip().isdigit()]
        }
        rst = self.input_reset_sheets.text().strip()
        if rst: opts["reset_colors_sheets"] = [x.strip() for x in rst.split(',')]

        cc = {}
        for line in self.custom_colors_txt.splitlines():
            if '=' in line:
                k, v = line.split('=', 1)
                k = k.strip().upper();
                v = v.strip()
                if k and len(k) == 1:
                    idx = ord(k) - ord('A')
                    cc[idx] = v if v else None
        if cc: opts["column_colors"] = cc

        # 5. Boucle principale de traitement des fichiers
        self.progress.setMaximum(len(self.files));
        self.progress.setValue(0)
        for i, f in enumerate(self.files):
            try:
                # Détermine le dossier de sortie pour le fichier actuel
                if self.chk_src_dir.isChecked():
                    base = Path(f).parent
                else:
                    base = Path(self.output_dir)
                    if not base.exists(): base.mkdir(parents=True)

                # Appel de la fonction backend pour traiter le fichier
                process_file(Path(f), base, "", "${stem}${suffix}${ext}", ops, opts)
                
                # Mise à jour de l'interface
                self.progress.setValue(i + 1);
                QApplication.processEvents() # Permet à l'UI de rester réactive
            except Exception as e:
                print(e) # Affiche l'erreur dans la console pour le débogage
        
        QMessageBox.information(self, "Terminé", "Traitement terminé !")


if __name__ == "__main__":
    """
    Point d'entrée de l'application.
    Initialise l'application Qt, applique un thème, crée et affiche la fenêtre principale.
    """
    app = QApplication.instance()
    if not app: app = QApplication(sys.argv)

    # 1. Applique le thème "dark_teal" à toute l'application
    apply_stylesheet(app, theme='dark_teal.xml')

    # 2. PATCH CSS : On force la couleur des placeholders en gris clair
    # Cela s'ajoute par dessus le thème dark_teal
    app.setStyleSheet(app.styleSheet() + """
        QLineEdit { placeholder-text-color: #E0E0E0; }
        QTextEdit { placeholder-text-color: #E0E0E0; }
        QLineEdit[text=""] { color: #E0E0E0; }
    """)

    window = ModernODSApp()
    window.show()
    app.exec()
