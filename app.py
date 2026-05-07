"""
============================================================================
SUIVI DE PRODUCTION ORC — Application de bureau
============================================================================

Une application autonome de saisie de déclarations de production et de pannes
pour les lignes ORC. Les données sont stockées dans un fichier Excel local.

Dépendances :
    customtkinter
    openpyxl
    Pillow

Usage :
    python app.py
"""

import os
import sys
import json
import math
from datetime import datetime, date, time
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ===========================================================================
# CONFIGURATION
# ===========================================================================

APP_TITLE = "Suivi de production — Lignes ORC"
APP_VERSION = "v1.0"
CONFIG_FILE = "config.json"

# Palette de couleurs (cohérente avec le mockup HTML)
COLORS = {
    "bleu":         "#2563EB",
    "bleu_fonce":   "#1E3A8A",
    "bleu_bg":      "#EFF6FF",
    "vert":         "#16A34A",
    "vert_bg":      "#DCFCE7",
    "rouge":        "#DC2626",
    "rouge_bg":     "#FEE2E2",
    "orange":       "#EA580C",
    "orange_bg":    "#FFEDD5",
    "violet":       "#7C3AED",
    "violet_bg":    "#EDE9FE",
    "gris_bg":      "#F1F5F9",
    "gris_panel":   "#F8FAFC",
    "gris_bord":    "#E2E8F0",
    "gris_label":   "#64748B",
    "gris_placeh":  "#94A3B8",
    "gris_text":    "#1E293B",
    "blanc":        "#FFFFFF",
}

# Listes pour les menus déroulants
LISTES = {
    "lignes":      ["ORC1", "ORC2"],
    "postes":      ["Matin", "Après-midi", "Nuit"],
    "pilotes":     ["Marie Lambert", "Thomas Bernard", "Julie Petit", "Karim Benali",
                    "Sophie Moreau", "Antoine Dubois", "Nadia Cherif", "Pierre Garnier"],
    "copilotes":   ["Lucas Martin", "Émilie Roux", "Mehdi Saïdi", "Camille Faure",
                    "Alexandre Vidal", "Fatima Khelifi", "Damien Robin", "Léa Fontaine"],
    "tailles":     ["40x60", "40x65", "50x70", "60x60", "65x65", "Traversin 140", "Autre…"],
    "fibres":      ["FIB-001", "FIB-002", "FIB-003", "FIB-MICRO"],
    "rattrapages": ["Pochon / fibre", "Couture", "Emballage", "Presse à souder", "Presse à housse ZIP"],
    "equipements": ["Chargeuse", "Carde", "Étaleur / Tour", "Coupe / Coupe circulaire",
                    "Tapis bascule / Tapis pesé N°1", "Enrouleur pochon", "Pesée / Tapis pesée N°2",
                    "Déviation pochon / Table de distribution", "Enfileuse pochon", "Kinna / Stroebel",
                    "Tapeuse", "Table rotative — Twin pack", "Enfileuse H100", "Enfileuse traversin",
                    "Presse ORS", "Presse à housse ZIP", "Cercleuse", "Enrouleuse traversin",
                    "OF taie", "Traçabilité fibre"],
    "types_panne": ["Panne", "Maintenance hebdomadaire", "Maintenance préventive", "Réglage"],
    "intervenants":["Pierre Garnier (technicien)", "Karim Benali (technicien)",
                    "Damien Robin (technicien)", "Service externe"],
    "nb_personnes": [str(i) for i in range(1, 13)],
}

# En-têtes des colonnes du fichier Excel
EXCEL_HEADERS = [
    "ID", "Type", "Date", "Heure", "Ligne", "Pilote", "Co-pilote", "N° OF",
    "Poste", "Code produit", "Taille", "Code fibre", "Poids garn. (g)", "Durée OF (min)",
    "Réf. taie", "Nb pers.", "Heure début", "Heure fin", "Qté fab.", "Qté emb.",
    "2nd choix", "Manq. taies", "Manq. housse", "Déf. couture",
    "Nb rattrap.", "Durée rattrap. (min)", "Nb pb tech.", "Durée pb tech. (min)",
    "Nb pannes", "Durée pannes (min)",
    "Équipement", "Type panne", "Intervenant", "Détail panne",
    "Détail rattrapages", "Détail pb tech.", "Commentaire",
]


# ===========================================================================
# CONFIG : CHARGEMENT / SAUVEGARDE DU CHEMIN DU FICHIER EXCEL
# ===========================================================================

def get_app_dir():
    """Retourne le dossier de l'application (compatible PyInstaller)."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def load_config():
    """Charge la config (chemin du fichier Excel)."""
    config_path = get_app_dir() / CONFIG_FILE
    if config_path.exists():
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(config):
    """Sauvegarde la config."""
    config_path = get_app_dir() / CONFIG_FILE
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Erreur sauvegarde config : {e}")


# ===========================================================================
# GESTION DU FICHIER EXCEL
# ===========================================================================

class ExcelStore:
    """Gère la lecture et l'écriture du fichier Excel des déclarations."""

    def __init__(self, filepath):
        self.filepath = Path(filepath)
        self._ensure_file()

    def _ensure_file(self):
        """Crée le fichier Excel avec en-têtes s'il n'existe pas."""
        if self.filepath.exists():
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "DATA"
        # Header
        thin = Side(border_style='thin', color='1E293B')
        for i, h in enumerate(EXCEL_HEADERS, start=1):
            c = ws.cell(row=1, column=i, value=h)
            c.font = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='1E3A8A')
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = max(12, len(h) + 2)
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = 'A2'
        wb.save(self.filepath)

    def get_all(self):
        """Retourne toutes les déclarations comme liste de dict."""
        if not self.filepath.exists():
            return []
        try:
            wb = load_workbook(self.filepath, data_only=True)
        except Exception as e:
            messagebox.showerror("Erreur lecture", f"Impossible de lire le fichier Excel :\n{e}")
            return []
        ws = wb["DATA"] if "DATA" in wb.sheetnames else wb.active
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            d = {}
            for i, h in enumerate(EXCEL_HEADERS):
                if i < len(row):
                    d[h] = row[i]
                else:
                    d[h] = None
            rows.append(d)
        wb.close()
        return rows

    def add(self, declaration):
        """Ajoute une déclaration au fichier Excel."""
        try:
            wb = load_workbook(self.filepath)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier Excel :\n{e}")
            return False

        ws = wb["DATA"] if "DATA" in wb.sheetnames else wb.active

        # Calcul du nouvel ID
        existing_ids = []
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0] is not None and isinstance(row[0], (int, float)):
                existing_ids.append(int(row[0]))
        new_id = max(existing_ids) + 1 if existing_ids else 1
        declaration["ID"] = new_id

        # Trouver la prochaine ligne vide
        next_row = ws.max_row + 1
        if ws.cell(row=2, column=1).value is None and ws.max_row == 2:
            next_row = 2

        for i, h in enumerate(EXCEL_HEADERS, start=1):
            value = declaration.get(h, "")
            cell = ws.cell(row=next_row, column=i, value=value)
            cell.font = Font(name='Calibri', size=10)
            cell.alignment = Alignment(horizontal='left', vertical='center')
            if h == "Date" and isinstance(value, (datetime, date)):
                cell.number_format = 'dd/mm/yyyy'
            elif h in ("Heure", "Heure début", "Heure fin") and value:
                cell.number_format = 'hh:mm'

        try:
            wb.save(self.filepath)
            wb.close()
            return True
        except PermissionError:
            messagebox.showerror(
                "Fichier verrouillé",
                "Le fichier Excel est ouvert dans Excel.\nFermez-le et réessayez."
            )
            return False
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'enregistrer :\n{e}")
            return False


# ===========================================================================
# WIDGETS PERSONNALISÉS
# ===========================================================================

class GaugeCanvas(ctk.CTkCanvas):
    """Jauge demi-cercle qui dessine un arc tricolore et une aiguille."""

    def __init__(self, parent, width=320, height=200, **kwargs):
        super().__init__(parent, width=width, height=height,
                         bg=COLORS["blanc"], highlightthickness=0, **kwargs)
        self.w = width
        self.h = height
        self._value = 0
        self.draw(0)

    def draw(self, value):
        self._value = value
        self.delete("all")

        cx, cy = self.w // 2, int(self.h * 0.85)
        radius = int(min(self.w, self.h * 2) * 0.40)
        thickness = 22

        # Arc rouge (0-60%)
        self._draw_arc(cx, cy, radius, 180, 180 - 108, COLORS["rouge"], thickness)
        # Arc orange (60-75%)
        self._draw_arc(cx, cy, radius, 180 - 108, 180 - 135, COLORS["orange"], thickness)
        # Arc vert (75-100%)
        self._draw_arc(cx, cy, radius, 180 - 135, 0, COLORS["vert"], thickness)

        # Graduations
        self.create_text(cx - radius - 10, cy + 12, text="0",
                         font=("Segoe UI", 9), fill=COLORS["gris_label"])
        self.create_text(cx, cy - radius - 18, text="75",
                         font=("Segoe UI", 9), fill=COLORS["gris_label"])
        self.create_text(cx + radius + 12, cy + 12, text="100",
                         font=("Segoe UI", 9), fill=COLORS["gris_label"])

        # Aiguille (rotation de -180° à 0° pour 0% à 100%)
        angle_deg = 180 - (value * 1.8)
        angle_rad = math.radians(angle_deg)
        needle_len = radius - thickness // 2 - 4
        x_end = cx + needle_len * math.cos(angle_rad)
        y_end = cy - needle_len * math.sin(angle_rad)
        self.create_line(cx, cy, x_end, y_end,
                         width=4, fill=COLORS["gris_text"], capstyle="round")
        self.create_oval(cx - 8, cy - 8, cx + 8, cy + 8,
                         fill=COLORS["gris_text"], outline="")

        # Valeur centrale
        self.create_text(cx, cy - radius * 0.45, text=f"{int(value)}%",
                         font=("Segoe UI", 36, "bold"), fill=COLORS["gris_text"])
        self.create_text(cx, cy - radius * 0.45 + 30, text="TRS",
                         font=("Segoe UI", 9), fill=COLORS["gris_label"])

    def _draw_arc(self, cx, cy, radius, start_deg, extent_deg, color, thickness):
        """Dessine un arc épais."""
        x0, y0 = cx - radius, cy - radius
        x1, y1 = cx + radius, cy + radius
        # Tkinter : start est l'angle de départ, extent est l'angle balayé
        # On veut aller de start_deg à start_deg + (extent_deg - start_deg)
        extent = extent_deg - start_deg
        self.create_arc(x0, y0, x1, y1,
                        start=start_deg, extent=extent,
                        style="arc", outline=color, width=thickness)


class StopBar(ctk.CTkFrame):
    """Une ligne d'arrêt avec icône, label, count, valeur, et barre de progression."""

    def __init__(self, parent, icon, label, color, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self.color = color
        self.value = 0
        self.count = 0

        # Ligne du haut
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x")

        self.icon_lbl = ctk.CTkLabel(top, text=icon, width=28, height=28,
                                     font=("Segoe UI", 14),
                                     fg_color=color, text_color=COLORS["blanc"],
                                     corner_radius=6)
        self.icon_lbl.pack(side="left")

        self.name_lbl = ctk.CTkLabel(top, text=label,
                                     font=("Segoe UI", 13, "bold"),
                                     text_color=COLORS["gris_text"], anchor="w")
        self.name_lbl.pack(side="left", padx=(10, 0), fill="x", expand=True)

        self.count_lbl = ctk.CTkLabel(top, text="0",
                                      font=("Segoe UI", 11),
                                      text_color=COLORS["gris_label"],
                                      fg_color=COLORS["gris_bg"], corner_radius=10,
                                      width=40, height=22)
        self.count_lbl.pack(side="left", padx=(0, 10))

        self.value_lbl = ctk.CTkLabel(top, text="0 min",
                                      font=("Segoe UI", 14, "bold"),
                                      text_color=COLORS["gris_text"], width=70)
        self.value_lbl.pack(side="right")

        # Barre de progression
        self.bar = ctk.CTkProgressBar(self, height=8,
                                      fg_color=COLORS["gris_bg"],
                                      progress_color=color,
                                      corner_radius=4)
        self.bar.pack(fill="x", pady=(6, 0), padx=(38, 0))
        self.bar.set(0)

    def update(self, count, value, max_value):
        self.count = count
        self.value = value
        self.count_lbl.configure(text=f"{count}")
        self.value_lbl.configure(text=f"{value} min")
        ratio = value / max_value if max_value > 0 else 0
        self.bar.set(ratio)


# ===========================================================================
# APPLICATION PRINCIPALE
# ===========================================================================

class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry("1280x820")
        self.minsize(1100, 700)
        self.configure(fg_color=COLORS["gris_bg"])

        # Charger la config et sélectionner le fichier Excel
        self.config_data = load_config()
        self.excel_path = self.config_data.get("excel_path")

        if not self.excel_path or not Path(self.excel_path).exists():
            self._select_excel_file()
            if not self.excel_path:
                self.destroy()
                return

        self.store = ExcelStore(self.excel_path)

        # Construire l'UI
        self._build_titlebar()
        self._build_main()
        self._refresh()

    # -----------------------------------------------------------------------
    # SÉLECTION DU FICHIER EXCEL
    # -----------------------------------------------------------------------

    def _select_excel_file(self):
        """Demande à l'utilisateur de choisir/créer le fichier Excel."""
        choice = messagebox.askyesnocancel(
            "Fichier de données",
            "Voulez-vous OUVRIR un fichier Excel existant ?\n\n"
            "• Oui → choisir un fichier existant\n"
            "• Non → créer un nouveau fichier\n"
            "• Annuler → quitter l'application"
        )

        if choice is None:
            return

        if choice:
            path = filedialog.askopenfilename(
                title="Sélectionner le fichier Excel",
                filetypes=[("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*")]
            )
        else:
            path = filedialog.asksaveasfilename(
                title="Créer un nouveau fichier Excel",
                defaultextension=".xlsx",
                initialfile="Donnees_ORC.xlsx",
                filetypes=[("Fichiers Excel", "*.xlsx")]
            )

        if not path:
            return

        self.excel_path = path
        self.config_data["excel_path"] = path
        save_config(self.config_data)

    def _change_excel_file(self):
        """Permet à l'utilisateur de changer le fichier Excel."""
        old_path = self.excel_path
        self._select_excel_file()
        if self.excel_path and self.excel_path != old_path:
            self.store = ExcelStore(self.excel_path)
            self._refresh()
            self.path_label.configure(text=self._short_path(self.excel_path))

    def _short_path(self, p):
        """Raccourcit un chemin pour l'affichage."""
        p = str(p)
        if len(p) <= 60:
            return p
        return "…" + p[-58:]

    # -----------------------------------------------------------------------
    # BARRE DE TITRE
    # -----------------------------------------------------------------------

    def _build_titlebar(self):
        bar = ctk.CTkFrame(self, fg_color=COLORS["bleu_fonce"], corner_radius=0, height=56)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        # Logo + titre
        left = ctk.CTkFrame(bar, fg_color="transparent")
        left.pack(side="left", padx=20, pady=10)

        ctk.CTkLabel(left, text="📋", font=("Segoe UI", 22),
                     fg_color="#FFFFFF20", corner_radius=8,
                     width=36, height=36).pack(side="left")

        ctk.CTkLabel(left, text=APP_TITLE,
                     font=("Segoe UI", 14, "bold"),
                     text_color=COLORS["blanc"]).pack(side="left", padx=(12, 0))

        # À droite
        right = ctk.CTkFrame(bar, fg_color="transparent")
        right.pack(side="right", padx=20)

        self.path_label = ctk.CTkLabel(right, text=self._short_path(self.excel_path),
                                       font=("Segoe UI", 11),
                                       text_color="#BFDBFE")
        self.path_label.pack(side="left", padx=(0, 12))

        ctk.CTkButton(right, text="Changer",
                      width=80, height=28,
                      font=("Segoe UI", 11),
                      fg_color="#FFFFFF20", hover_color="#FFFFFF40",
                      text_color=COLORS["blanc"], corner_radius=6,
                      command=self._change_excel_file).pack(side="left")

    # -----------------------------------------------------------------------
    # ZONE PRINCIPALE
    # -----------------------------------------------------------------------

    def _build_main(self):
        """Conteneur principal avec deux vues : tableau de bord + formulaires."""
        self.container = ctk.CTkFrame(self, fg_color=COLORS["gris_bg"], corner_radius=0)
        self.container.pack(fill="both", expand=True)

        # Construit le tableau de bord (vue par défaut)
        self._build_dashboard()

    def _clear_container(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    def _build_dashboard(self):
        self._clear_container()

        # Scrollable
        scroll = ctk.CTkScrollableFrame(self.container, fg_color=COLORS["gris_bg"],
                                        corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=24, pady=20)

        # ====== HERO ======
        hero = ctk.CTkFrame(scroll, fg_color="transparent")
        hero.pack(fill="x", pady=(0, 16))

        ctk.CTkLabel(hero, text="Tableau de bord",
                     font=("Segoe UI", 26, "bold"),
                     text_color=COLORS["gris_text"], anchor="w").pack(fill="x")
        ctk.CTkLabel(hero, text="État de la production en temps réel.",
                     font=("Segoe UI", 13),
                     text_color=COLORS["gris_label"], anchor="w").pack(fill="x", pady=(2, 0))

        # ====== KPI CARD ======
        kpi_card = ctk.CTkFrame(scroll, fg_color=COLORS["blanc"],
                                corner_radius=12, border_width=1,
                                border_color=COLORS["gris_bord"])
        kpi_card.pack(fill="x", pady=(0, 16))

        # Titre KPI (sera mis à jour avec l'heure de la dernière déclaration)
        title_frame = ctk.CTkFrame(kpi_card, fg_color="transparent")
        title_frame.pack(fill="x", padx=20, pady=(20, 12))

        self.kpi_title_lbl = ctk.CTkLabel(title_frame,
                                          text="TRS de la ligne ORC",
                                          font=("Segoe UI", 16, "bold"),
                                          text_color=COLORS["gris_text"], anchor="w")
        self.kpi_title_lbl.pack(side="left")

        # Body : 2 colonnes (jauge | quantité+arrêts)
        body = ctk.CTkFrame(kpi_card, fg_color="transparent")
        body.pack(fill="x", padx=20, pady=(0, 20))
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        # Colonne gauche : jauge
        gauge_frame = ctk.CTkFrame(body, fg_color="transparent")
        gauge_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        self.gauge = GaugeCanvas(gauge_frame, width=320, height=200)
        self.gauge.pack()

        self.status_lbl = ctk.CTkLabel(gauge_frame, text="—",
                                       font=("Segoe UI", 12, "bold"),
                                       fg_color=COLORS["vert_bg"],
                                       text_color=COLORS["vert"],
                                       corner_radius=999,
                                       width=140, height=28)
        self.status_lbl.pack(pady=(8, 0))

        # Colonne droite : quantité + arrêts
        right_col = ctk.CTkFrame(body, fg_color="transparent")
        right_col.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

        # Card quantité (dégradé bleu)
        qty_card = ctk.CTkFrame(right_col, fg_color=COLORS["bleu"],
                                corner_radius=10)
        qty_card.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(qty_card, text="QUANTITÉ PIÈCES PRODUITES",
                     font=("Segoe UI", 11, "bold"),
                     text_color="#DBEAFE", anchor="w").pack(fill="x", padx=16, pady=(14, 4))

        qty_row = ctk.CTkFrame(qty_card, fg_color="transparent")
        qty_row.pack(fill="x", padx=16)

        ctk.CTkLabel(qty_row, text="📦", font=("Segoe UI", 22),
                     text_color=COLORS["blanc"]).pack(side="left")
        self.qty_value_lbl = ctk.CTkLabel(qty_row, text="0",
                                          font=("Segoe UI", 32, "bold"),
                                          text_color=COLORS["blanc"])
        self.qty_value_lbl.pack(side="left", padx=(10, 6))
        ctk.CTkLabel(qty_row, text="pcs", font=("Segoe UI", 14),
                     text_color="#DBEAFE").pack(side="left", anchor="s", pady=(0, 6))

        self.qty_sub_lbl = ctk.CTkLabel(qty_card, text="—",
                                        font=("Segoe UI", 11),
                                        text_color="#DBEAFE", anchor="w")
        self.qty_sub_lbl.pack(fill="x", padx=16, pady=(4, 14))

        # Card arrêts
        arr_card = ctk.CTkFrame(right_col, fg_color=COLORS["gris_panel"],
                                corner_radius=10, border_width=1,
                                border_color=COLORS["gris_bord"])
        arr_card.pack(fill="both", expand=True)

        arr_head = ctk.CTkFrame(arr_card, fg_color="transparent")
        arr_head.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(arr_head, text="ARRÊTS CUMULÉS",
                     font=("Segoe UI", 11, "bold"),
                     text_color=COLORS["gris_label"], anchor="w").pack(side="left")
        self.arr_total_lbl = ctk.CTkLabel(arr_head, text="0 min",
                                          font=("Segoe UI", 18, "bold"),
                                          text_color=COLORS["rouge"])
        self.arr_total_lbl.pack(side="right")

        ctk.CTkFrame(arr_card, fg_color=COLORS["gris_bord"],
                     height=1).pack(fill="x", padx=16)

        # 3 lignes d'arrêts
        bars_frame = ctk.CTkFrame(arr_card, fg_color="transparent")
        bars_frame.pack(fill="x", padx=16, pady=(10, 14))

        self.bar_pannes = StopBar(bars_frame, "⚡", "Pannes / Maintenance", COLORS["orange"])
        self.bar_pannes.pack(fill="x", pady=4)
        self.bar_ratt = StopBar(bars_frame, "⏱", "Rattrapages", COLORS["violet"])
        self.bar_ratt.pack(fill="x", pady=4)
        self.bar_tech = StopBar(bars_frame, "🔧", "Problèmes techniques", COLORS["rouge"])
        self.bar_tech.pack(fill="x", pady=4)

        # ====== BOUTONS ======
        btns = ctk.CTkFrame(scroll, fg_color="transparent")
        btns.pack(fill="x", pady=(0, 20))
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)

        btn_prod = ctk.CTkButton(btns,
                                 text="📦   Déclarer une production",
                                 font=("Segoe UI", 15, "bold"),
                                 fg_color=COLORS["bleu"], hover_color=COLORS["bleu_fonce"],
                                 text_color=COLORS["blanc"],
                                 corner_radius=12, height=72,
                                 command=self._open_form_production)
        btn_prod.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        btn_panne = ctk.CTkButton(btns,
                                  text="⚡   Déclarer une panne / maintenance",
                                  font=("Segoe UI", 15, "bold"),
                                  fg_color=COLORS["orange"], hover_color="#C2410C",
                                  text_color=COLORS["blanc"],
                                  corner_radius=12, height=72,
                                  command=self._open_form_panne)
        btn_panne.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        # ====== HISTORIQUE ======
        hist_card = ctk.CTkFrame(scroll, fg_color=COLORS["blanc"],
                                 corner_radius=12, border_width=1,
                                 border_color=COLORS["gris_bord"])
        hist_card.pack(fill="x")

        # Header
        hist_head = ctk.CTkFrame(hist_card, fg_color="transparent")
        hist_head.pack(fill="x", padx=20, pady=14)

        ctk.CTkLabel(hist_head, text="Déclarations récentes — 20 dernières",
                     font=("Segoe UI", 15, "bold"),
                     text_color=COLORS["gris_text"], anchor="w").pack(side="left")

        # Filtres
        self.filter_var = ctk.StringVar(value="all")
        filt_frame = ctk.CTkFrame(hist_head, fg_color="transparent")
        filt_frame.pack(side="right")

        for label, value in [("Tout", "all"), ("Production", "production"), ("Pannes", "panne")]:
            btn = ctk.CTkButton(filt_frame, text=label,
                                width=80, height=26, font=("Segoe UI", 11, "bold"),
                                corner_radius=6,
                                fg_color=COLORS["gris_text"] if value == "all" else COLORS["blanc"],
                                text_color=COLORS["blanc"] if value == "all" else COLORS["gris_label"],
                                hover_color=COLORS["gris_text"],
                                border_width=1, border_color=COLORS["gris_bord"],
                                command=lambda v=value, l=label: self._set_filter(v, l))
            btn.pack(side="left", padx=2)
            if value == "all":
                self._active_filter_btn = btn
            btn._filter_value = value

        ctk.CTkFrame(hist_card, fg_color=COLORS["gris_bord"], height=1).pack(fill="x")

        # Tableau (frame scrollable horizontale au cas où)
        self.table_frame = ctk.CTkFrame(hist_card, fg_color=COLORS["blanc"], corner_radius=0)
        self.table_frame.pack(fill="x", padx=20, pady=(0, 20))

    def _set_filter(self, value, label):
        # Reset tous les boutons
        for btn in self._active_filter_btn.master.winfo_children():
            if hasattr(btn, "_filter_value"):
                btn.configure(fg_color=COLORS["blanc"], text_color=COLORS["gris_label"])
        # Activer le sélectionné
        for btn in self._active_filter_btn.master.winfo_children():
            if hasattr(btn, "_filter_value") and btn._filter_value == value:
                btn.configure(fg_color=COLORS["gris_text"], text_color=COLORS["blanc"])
                self._active_filter_btn = btn
                break
        self.filter_var.set(value)
        self._refresh_table()

    # -----------------------------------------------------------------------
    # RAFRAÎCHISSEMENT
    # -----------------------------------------------------------------------

    def _refresh(self):
        """Rafraîchit les KPI et le tableau."""
        self._refresh_kpi()
        self._refresh_table()

    def _refresh_kpi(self):
        decls = self.store.get_all()
        prod_decls = [d for d in decls if d.get("Type") == "production"]

        # Quantité totale
        total_qte = 0
        for d in prod_decls:
            v = d.get("Qté fab.", 0)
            if isinstance(v, (int, float)):
                total_qte += int(v)

        self.qty_value_lbl.configure(text=f"{total_qte:,}".replace(",", " "))
        nb_of = len(prod_decls)
        self.qty_sub_lbl.configure(text=f"{nb_of} OF terminé{'s' if nb_of > 1 else ''}")

        # Arrêts par type
        nb_p, dur_p = 0, 0
        nb_r, dur_r = 0, 0
        nb_t, dur_t = 0, 0
        for d in decls:
            np_ = d.get("Nb pannes") or 0
            dp = d.get("Durée pannes (min)") or 0
            nr = d.get("Nb rattrap.") or 0
            dr = d.get("Durée rattrap. (min)") or 0
            nt = d.get("Nb pb tech.") or 0
            dt = d.get("Durée pb tech. (min)") or 0
            try:
                nb_p += int(np_); dur_p += int(dp)
                nb_r += int(nr); dur_r += int(dr)
                nb_t += int(nt); dur_t += int(dt)
            except (ValueError, TypeError):
                pass

        total = dur_p + dur_r + dur_t
        max_v = max(dur_p, dur_r, dur_t, 1)

        self.bar_pannes.update(nb_p, dur_p, max_v)
        self.bar_ratt.update(nb_r, dur_r, max_v)
        self.bar_tech.update(nb_t, dur_t, max_v)
        self.arr_total_lbl.configure(text=f"{total} min")

        # TRS simulé : 100 - (arrêts en min / temps total disponible)
        # Pour un mockup, on calcule de manière simple
        if total > 0 and nb_of > 0:
            # Calcul simpliste : plus il y a d'arrêts, plus le TRS baisse
            trs = max(20, min(95, 95 - (total / 5)))
        else:
            trs = 85 if nb_of > 0 else 0

        self.gauge.draw(trs)

        # Statut
        if trs >= 75:
            self.status_lbl.configure(text="✓ Objectif atteint",
                                      fg_color=COLORS["vert_bg"], text_color=COLORS["vert"])
        elif trs >= 60:
            self.status_lbl.configure(text="⚠ En vigilance",
                                      fg_color=COLORS["orange_bg"], text_color=COLORS["orange"])
        else:
            self.status_lbl.configure(text="✗ Sous objectif",
                                      fg_color=COLORS["rouge_bg"], text_color=COLORS["rouge"])

        # Titre KPI : ligne + heure de la dernière déclaration
        if decls:
            last = max(decls, key=lambda d: (str(d.get("Date") or ""), str(d.get("Heure") or "")))
            heure = last.get("Heure")
            ligne = last.get("Ligne") or "ORC"
            if isinstance(heure, time):
                heure_str = heure.strftime("%H:%M")
            elif isinstance(heure, datetime):
                heure_str = heure.strftime("%H:%M")
            else:
                heure_str = str(heure or "—")
            self.kpi_title_lbl.configure(text=f"TRS de la ligne {ligne} à {heure_str}")
        else:
            self.kpi_title_lbl.configure(text="TRS de la ligne ORC")

    def _refresh_table(self):
        # Nettoyer
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        decls = self.store.get_all()
        flt = self.filter_var.get()
        if flt != "all":
            decls = [d for d in decls if d.get("Type") == flt]

        # Trier par date+heure desc
        def sort_key(d):
            dt = d.get("Date")
            hr = d.get("Heure")
            d_str = dt.strftime("%Y%m%d") if isinstance(dt, (datetime, date)) else str(dt or "")
            h_str = hr.strftime("%H%M") if isinstance(hr, (datetime, time)) else str(hr or "")
            return d_str + h_str

        decls = sorted(decls, key=sort_key, reverse=True)[:20]

        # Header
        headers = [
            ("DATE / HEURE", 130),
            ("TYPE", 80),
            ("OF", 110),
            ("LIGNE", 60),
            ("PILOTE", 140),
            ("DURÉE", 70),
            ("DÉTAIL", 200),
            ("QUANTITÉ", 90),
        ]
        head_frame = ctk.CTkFrame(self.table_frame, fg_color=COLORS["gris_panel"],
                                  corner_radius=0, height=36)
        head_frame.pack(fill="x")
        head_frame.pack_propagate(False)

        for label, w in headers:
            ctk.CTkLabel(head_frame, text=label, width=w,
                         font=("Segoe UI", 10, "bold"),
                         text_color=COLORS["gris_label"], anchor="w"
                         ).pack(side="left", padx=(8, 0))

        if not decls:
            empty = ctk.CTkLabel(self.table_frame,
                                 text="Aucune déclaration",
                                 font=("Segoe UI", 12, "italic"),
                                 text_color=COLORS["gris_placeh"],
                                 height=60)
            empty.pack(fill="x")
            return

        for i, d in enumerate(decls):
            bg = COLORS["blanc"] if i % 2 == 0 else COLORS["gris_panel"]
            row = ctk.CTkFrame(self.table_frame, fg_color=bg, corner_radius=0, height=36)
            row.pack(fill="x")
            row.pack_propagate(False)

            # Format date/heure
            dt = d.get("Date")
            hr = d.get("Heure")
            d_str = dt.strftime("%d/%m/%Y") if isinstance(dt, (datetime, date)) else (str(dt or ""))
            h_str = hr.strftime("%H:%M") if isinstance(hr, (datetime, time)) else (str(hr or ""))

            type_val = d.get("Type") or ""
            type_label = "Production" if type_val == "production" else "Panne" if type_val == "panne" else "—"
            type_color = COLORS["bleu"] if type_val == "production" else COLORS["orange"]
            type_bg = COLORS["bleu_bg"] if type_val == "production" else COLORS["orange_bg"]

            qte = d.get("Qté fab.")
            qte_str = f"{int(qte):,} pcs".replace(",", " ") if isinstance(qte, (int, float)) and qte else "—"

            duree = d.get("Durée OF (min)") or d.get("Durée pannes (min)") or 0
            try:
                duree_str = f"{int(duree)} min"
            except (ValueError, TypeError):
                duree_str = "—"

            detail_parts = []
            if d.get("Taille"):
                detail_parts.append(str(d.get("Taille")))
            if d.get("Équipement"):
                detail_parts.append(str(d.get("Équipement")))
            detail_str = " · ".join(detail_parts) or "—"

            cols = [
                (f"{d_str} · {h_str}", 130, COLORS["gris_text"], None),
                (type_label, 80, type_color, type_bg),
                (str(d.get("N° OF") or "—"), 110, COLORS["gris_text"], None),
                (str(d.get("Ligne") or "—"), 60, COLORS["gris_label"], None),
                (str(d.get("Pilote") or "—"), 140, COLORS["gris_text"], None),
                (duree_str, 70, COLORS["gris_text"], None),
                (detail_str, 200, COLORS["gris_label"], None),
                (qte_str, 90, COLORS["gris_text"], None),
            ]

            for txt, w, color, bgcol in cols:
                if bgcol:
                    lbl = ctk.CTkLabel(row, text=txt, width=w,
                                       font=("Segoe UI", 11, "bold"),
                                       text_color=color, fg_color=bgcol,
                                       corner_radius=4, anchor="center")
                else:
                    lbl = ctk.CTkLabel(row, text=txt, width=w,
                                       font=("Segoe UI", 11),
                                       text_color=color, anchor="w")
                lbl.pack(side="left", padx=(8, 0))

            # Bordure basse
            ctk.CTkFrame(self.table_frame, fg_color=COLORS["gris_bord"],
                         height=1).pack(fill="x")

    # -----------------------------------------------------------------------
    # FORMULAIRE PRODUCTION
    # -----------------------------------------------------------------------

    def _open_form_production(self):
        FormProduction(self, self.store, on_save=self._refresh)

    def _open_form_panne(self):
        FormPanne(self, self.store, on_save=self._refresh)


# ===========================================================================
# FENÊTRE FORMULAIRE PRODUCTION
# ===========================================================================

class FormProduction(ctk.CTkToplevel):

    def __init__(self, parent, store, on_save=None):
        super().__init__(parent)
        self.store = store
        self.on_save = on_save
        self.title("Déclarer une production")
        self.geometry("900x800")
        self.configure(fg_color=COLORS["gris_bg"])
        self.transient(parent)
        self.grab_set()

        self.rattrap_rows = []
        self.probleme_rows = []

        self._build()

    def _build(self):
        # Header bleu
        head = ctk.CTkFrame(self, fg_color=COLORS["bleu"], corner_radius=0, height=72)
        head.pack(fill="x")
        head.pack_propagate(False)

        ctk.CTkLabel(head, text="📦   Déclarer une production",
                     font=("Segoe UI", 18, "bold"),
                     text_color=COLORS["blanc"], anchor="w"
                     ).pack(side="left", padx=24)

        # Body scrollable
        scroll = ctk.CTkScrollableFrame(self, fg_color=COLORS["gris_bg"], corner_radius=0)
        scroll.pack(fill="both", expand=True)

        body = ctk.CTkFrame(scroll, fg_color=COLORS["blanc"],
                            corner_radius=12, border_width=1,
                            border_color=COLORS["gris_bord"])
        body.pack(fill="x", padx=20, pady=20)

        # ===== Section 1 : Information =====
        self._section(body, "1", "👤", "Information")

        row1 = ctk.CTkFrame(body, fg_color="transparent")
        row1.pack(fill="x", padx=20, pady=(0, 12))

        self.f_date = self._field(row1, "Date *", side="left", width=200)
        self.f_date.insert(0, date.today().strftime("%d/%m/%Y"))

        self.f_ligne = self._dropdown(row1, "Ligne *", LISTES["lignes"], side="left", width=200, default="ORC1")
        self.f_poste = self._dropdown(row1, "Poste *", LISTES["postes"], side="left", width=200, default="Matin")

        row2 = ctk.CTkFrame(body, fg_color="transparent")
        row2.pack(fill="x", padx=20, pady=(0, 12))
        self.f_pilote = self._dropdown(row2, "Pilote *", LISTES["pilotes"], side="left", width=300)
        self.f_copilote = self._dropdown(row2, "Co-pilote", LISTES["copilotes"], side="left", width=300)

        self._divider(body)

        # ===== Section 2 : Production =====
        self._section(body, "2", "📦", "Production")

        rowp1 = ctk.CTkFrame(body, fg_color="transparent")
        rowp1.pack(fill="x", padx=20, pady=(0, 12))
        self.f_numof = self._field(rowp1, "N° OF *", side="left", width=200)
        self.f_codeprod = self._field(rowp1, "Code produit", side="left", width=200)
        self.f_taille = self._dropdown(rowp1, "Taille", LISTES["tailles"], side="left", width=200)

        rowp2 = ctk.CTkFrame(body, fg_color="transparent")
        rowp2.pack(fill="x", padx=20, pady=(0, 12))
        self.f_fibre = self._dropdown(rowp2, "Code fibre", LISTES["fibres"], side="left", width=200)
        self.f_poids = self._field(rowp2, "Poids garnissage (g)", side="left", width=200)
        self.f_reftaie = self._field(rowp2, "Réf. taie", side="left", width=200)

        rowp3 = ctk.CTkFrame(body, fg_color="transparent")
        rowp3.pack(fill="x", padx=20, pady=(0, 12))
        self.f_hdebut = self._field(rowp3, "Heure début (hh:mm)", side="left", width=200)
        self.f_hfin = self._field(rowp3, "Heure fin (hh:mm)", side="left", width=200)
        self.f_duree_lbl = self._auto_field(rowp3, "Durée OF", "—", side="left", width=200)
        self.f_nbpers = self._dropdown(rowp3, "Nb personnes", LISTES["nb_personnes"], side="left", width=140)

        # Recalcul auto
        self.f_hdebut.bind("<KeyRelease>", self._recompute)
        self.f_hfin.bind("<KeyRelease>", self._recompute)

        rowp4 = ctk.CTkFrame(body, fg_color="transparent")
        rowp4.pack(fill="x", padx=20, pady=(0, 12))
        self.f_qtefab = self._field(rowp4, "Qté fabriquée", side="left", width=200)
        self.f_qteemb = self._field(rowp4, "Qté emballée", side="left", width=200)
        self.f_equiv_lbl = self._auto_field(rowp4, "Équivalence", "—", side="left", width=200)
        self.f_2nd = self._field(rowp4, "Taies 2nd choix", side="left", width=140)

        self.f_qtefab.bind("<KeyRelease>", self._recompute)
        self.f_qteemb.bind("<KeyRelease>", self._recompute)
        self.f_nbpers.configure(command=lambda v: self._recompute())

        rowp5 = ctk.CTkFrame(body, fg_color="transparent")
        rowp5.pack(fill="x", padx=20, pady=(0, 12))
        self.f_cad_min_lbl = self._auto_field(rowp5, "Cadence (or/min)", "—", side="left", width=200)
        self.f_cad_pers_lbl = self._auto_field(rowp5, "Cadence (or/pers/h)", "—", side="left", width=200)

        self._divider(body)

        # ===== Section 3 : Matière première =====
        self._section(body, "3", "🧵", "Matière première", optional=True)

        rowm = ctk.CTkFrame(body, fg_color="transparent")
        rowm.pack(fill="x", padx=20, pady=(0, 12))
        self.f_manq_taies = self._field(rowm, "Manquants taies", side="left", width=200)
        self.f_manq_housse = self._field(rowm, "Manquants housse / encart", side="left", width=300)

        self._divider(body)

        # ===== Section 4 : Rattrapages =====
        self._section(body, "4", "⏱", "Rattrapages", optional=True, hint="(jusqu'à 5)")

        self.rattrap_container = ctk.CTkFrame(body, fg_color="transparent")
        self.rattrap_container.pack(fill="x", padx=20, pady=(0, 8))

        ctk.CTkButton(body, text="+ Ajouter un rattrapage",
                      font=("Segoe UI", 12, "bold"),
                      fg_color=COLORS["bleu_bg"], hover_color=COLORS["blanc"],
                      text_color=COLORS["bleu"],
                      border_width=1, border_color="#BFDBFE",
                      corner_radius=6, height=36,
                      command=self._add_rattrap).pack(fill="x", padx=20, pady=(0, 12))

        self._add_rattrap()  # Une ligne par défaut

        self._divider(body)

        # ===== Section 5 : Problèmes techniques =====
        self._section(body, "5", "🔧", "Problèmes techniques", optional=True, hint="(jusqu'à 10)")

        rowdef = ctk.CTkFrame(body, fg_color="transparent")
        rowdef.pack(fill="x", padx=20, pady=(0, 12))
        self.f_def_couture = self._field(rowdef, "Nombre de défauts couture", side="left", width=300)

        self.probleme_container = ctk.CTkFrame(body, fg_color="transparent")
        self.probleme_container.pack(fill="x", padx=20, pady=(0, 8))

        ctk.CTkButton(body, text="+ Ajouter un problème technique",
                      font=("Segoe UI", 12, "bold"),
                      fg_color=COLORS["bleu_bg"], hover_color=COLORS["blanc"],
                      text_color=COLORS["bleu"],
                      border_width=1, border_color="#BFDBFE",
                      corner_radius=6, height=36,
                      command=self._add_probleme).pack(fill="x", padx=20, pady=(0, 12))

        self._add_probleme()

        self._divider(body)

        # ===== Section 6 : Commentaire =====
        self._section(body, "6", "💬", "Commentaire général", optional=True)

        self.f_comment = ctk.CTkTextbox(body, height=100,
                                        font=("Segoe UI", 12),
                                        fg_color=COLORS["blanc"],
                                        border_color=COLORS["gris_bord"],
                                        border_width=1, corner_radius=6)
        self.f_comment.pack(fill="x", padx=20, pady=(0, 20))

        # ===== Footer =====
        footer = ctk.CTkFrame(self, fg_color=COLORS["gris_panel"], corner_radius=0, height=72)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        ctk.CTkLabel(footer, text="Champs * obligatoires",
                     font=("Segoe UI", 11),
                     text_color=COLORS["gris_label"]
                     ).pack(side="left", padx=20)

        btn_save = ctk.CTkButton(footer, text="✓   Enregistrer la déclaration",
                                 font=("Segoe UI", 13, "bold"),
                                 fg_color=COLORS["bleu"], hover_color=COLORS["bleu_fonce"],
                                 text_color=COLORS["blanc"],
                                 corner_radius=8, height=44, width=240,
                                 command=self._save)
        btn_save.pack(side="right", padx=20, pady=14)

        btn_cancel = ctk.CTkButton(footer, text="Annuler",
                                   font=("Segoe UI", 13),
                                   fg_color=COLORS["blanc"], hover_color=COLORS["gris_bg"],
                                   text_color=COLORS["gris_label"],
                                   border_width=1, border_color=COLORS["gris_bord"],
                                   corner_radius=8, height=44, width=120,
                                   command=self.destroy)
        btn_cancel.pack(side="right", padx=(0, 0), pady=14)

    # -- Helpers UI --

    def _section(self, parent, num, icon, title, optional=False, hint=""):
        sec = ctk.CTkFrame(parent, fg_color="transparent")
        sec.pack(fill="x", padx=20, pady=(20, 12))

        pill_color = COLORS["bleu"] if not optional else COLORS["gris_bord"]
        pill_text_color = COLORS["blanc"] if not optional else COLORS["gris_label"]

        ctk.CTkLabel(sec, text=num, width=30, height=30,
                     font=("Segoe UI", 14, "bold"),
                     fg_color=pill_color, text_color=pill_text_color,
                     corner_radius=15).pack(side="left")

        ctk.CTkLabel(sec, text=icon, font=("Segoe UI", 18),
                     text_color=COLORS["gris_text"]).pack(side="left", padx=(10, 6))

        ctk.CTkLabel(sec, text=title, font=("Segoe UI", 15, "bold"),
                     text_color=COLORS["gris_text"]).pack(side="left")

        if hint:
            ctk.CTkLabel(sec, text=hint, font=("Segoe UI", 11),
                         text_color=COLORS["gris_placeh"]).pack(side="left", padx=(8, 0))

    def _divider(self, parent):
        ctk.CTkFrame(parent, fg_color=COLORS["gris_bord"], height=1
                     ).pack(fill="x", padx=20, pady=(8, 0))

    def _field(self, parent, label_text, side="left", width=200, expand=False, default=""):
        wrap = ctk.CTkFrame(parent, fg_color="transparent")
        wrap.pack(side=side, padx=4, fill="x", expand=expand)

        ctk.CTkLabel(wrap, text=label_text.upper(),
                     font=("Segoe UI", 10, "bold"),
                     text_color=COLORS["rouge"] if "*" in label_text else COLORS["gris_label"],
                     anchor="w").pack(fill="x", pady=(0, 4))

        entry = ctk.CTkEntry(wrap, width=width, height=36,
                             font=("Segoe UI", 12),
                             fg_color=COLORS["blanc"],
                             border_color=COLORS["gris_bord"],
                             border_width=1, corner_radius=6,
                             text_color=COLORS["gris_text"])
        entry.pack(fill="x")
        if default:
            entry.insert(0, default)
        return entry

    def _dropdown(self, parent, label_text, values, side="left", width=200, expand=False, default=""):
        wrap = ctk.CTkFrame(parent, fg_color="transparent")
        wrap.pack(side=side, padx=4, fill="x", expand=expand)

        ctk.CTkLabel(wrap, text=label_text.upper(),
                     font=("Segoe UI", 10, "bold"),
                     text_color=COLORS["rouge"] if "*" in label_text else COLORS["gris_label"],
                     anchor="w").pack(fill="x", pady=(0, 4))

        var = ctk.StringVar(value=default if default else "Sélectionner…")
        combo = ctk.CTkComboBox(wrap, width=width, height=36,
                                values=values, variable=var,
                                font=("Segoe UI", 12),
                                fg_color=COLORS["blanc"],
                                border_color=COLORS["gris_bord"],
                                button_color=COLORS["bleu"],
                                button_hover_color=COLORS["bleu_fonce"],
                                dropdown_fg_color=COLORS["blanc"],
                                dropdown_text_color=COLORS["gris_text"],
                                text_color=COLORS["gris_text"],
                                state="readonly")
        combo.pack(fill="x")
        return combo

    def _auto_field(self, parent, label_text, value, side="left", width=200):
        wrap = ctk.CTkFrame(parent, fg_color="transparent")
        wrap.pack(side=side, padx=4)

        ctk.CTkLabel(wrap, text=label_text.upper(),
                     font=("Segoe UI", 10, "bold"),
                     text_color=COLORS["gris_label"],
                     anchor="w").pack(fill="x", pady=(0, 4))

        lbl = ctk.CTkLabel(wrap, text=f"{value}  · auto",
                           width=width, height=36,
                           font=("Segoe UI", 12, "bold"),
                           text_color=COLORS["bleu_fonce"],
                           fg_color=COLORS["bleu_bg"], corner_radius=6,
                           anchor="w")
        lbl.pack(fill="x")
        lbl._auto_value = value
        return lbl

    def _add_rattrap(self):
        if len(self.rattrap_rows) >= 5:
            messagebox.showinfo("Limite", "Maximum 5 rattrapages.")
            return
        self._add_repeater_row(self.rattrap_container, self.rattrap_rows,
                                LISTES["rattrapages"], "Type de rattrapage")

    def _add_probleme(self):
        if len(self.probleme_rows) >= 10:
            messagebox.showinfo("Limite", "Maximum 10 problèmes techniques.")
            return
        self._add_repeater_row(self.probleme_container, self.probleme_rows,
                                LISTES["equipements"], "Équipement concerné")

    def _add_repeater_row(self, container, store_list, options, placeholder):
        idx = len(store_list) + 1
        row = ctk.CTkFrame(container, fg_color=COLORS["gris_panel"],
                           corner_radius=8, border_width=1, border_color=COLORS["gris_bord"])
        row.pack(fill="x", pady=4)

        # Numéro
        num = ctk.CTkLabel(row, text=str(idx), width=24,
                           font=("Segoe UI", 11, "bold"),
                           text_color=COLORS["gris_label"])
        num.pack(side="left", padx=(10, 8), pady=8)

        # Type / Équipement
        var_type = ctk.StringVar(value="Sélectionner…")
        combo = ctk.CTkComboBox(row, width=240, height=32, values=options,
                                variable=var_type, font=("Segoe UI", 11),
                                fg_color=COLORS["blanc"],
                                border_color=COLORS["gris_bord"],
                                button_color=COLORS["bleu"],
                                state="readonly")
        combo.pack(side="left", padx=4, pady=8)

        # Commentaire
        comment = ctk.CTkEntry(row, height=32, font=("Segoe UI", 11),
                               placeholder_text="Commentaire (optionnel)",
                               fg_color=COLORS["blanc"],
                               border_color=COLORS["gris_bord"])
        comment.pack(side="left", padx=4, pady=8, fill="x", expand=True)

        # Durée
        duree = ctk.CTkEntry(row, width=80, height=32,
                             font=("Segoe UI", 11),
                             placeholder_text="min",
                             fg_color=COLORS["blanc"],
                             border_color=COLORS["gris_bord"])
        duree.pack(side="left", padx=4, pady=8)

        # Supprimer
        def remove():
            row.destroy()
            store_list[:] = [r for r in store_list if r["frame"] != row]
            for i, r in enumerate(store_list, start=1):
                r["num_lbl"].configure(text=str(i))

        del_btn = ctk.CTkButton(row, text="✕", width=32, height=32,
                                font=("Segoe UI", 14),
                                fg_color="transparent",
                                hover_color=COLORS["rouge_bg"],
                                text_color=COLORS["gris_placeh"],
                                command=remove)
        del_btn.pack(side="left", padx=(4, 10), pady=8)

        store_list.append({
            "frame": row, "num_lbl": num,
            "type": combo, "comment": comment, "duree": duree,
        })

    def _recompute(self, *_):
        # Durée OF
        try:
            hd = self._parse_time(self.f_hdebut.get())
            hf = self._parse_time(self.f_hfin.get())
            if hd and hf:
                duree = (hf[0] * 60 + hf[1]) - (hd[0] * 60 + hd[1])
                if duree < 0:
                    duree += 24 * 60
                self.f_duree_lbl.configure(text=f"{duree} min  · auto")
                self.f_duree_lbl._auto_value = duree
            else:
                self.f_duree_lbl.configure(text="—  · auto")
                self.f_duree_lbl._auto_value = 0
        except Exception:
            pass

        # Cadences
        try:
            qte = float(self.f_qtefab.get() or 0)
            duree = self.f_duree_lbl._auto_value or 0
            nb_pers = float(self.f_nbpers.get() or 0)

            if qte and duree:
                self.f_cad_min_lbl.configure(text=f"{qte/duree:.2f}  · auto")
            else:
                self.f_cad_min_lbl.configure(text="—  · auto")

            if qte and duree and nb_pers:
                cad_pers_h = qte / nb_pers / (duree / 60)
                self.f_cad_pers_lbl.configure(text=f"{int(cad_pers_h)}  · auto")
            else:
                self.f_cad_pers_lbl.configure(text="—  · auto")
        except Exception:
            pass

        # Équivalence
        try:
            qte = float(self.f_qtefab.get() or 0)
            qte_emb = float(self.f_qteemb.get() or 0)
            if qte and qte_emb:
                self.f_equiv_lbl.configure(text=f"{(qte_emb/qte)*100:.1f}%  · auto")
            else:
                self.f_equiv_lbl.configure(text="—  · auto")
        except Exception:
            pass

    def _parse_time(self, s):
        s = (s or "").strip().replace("h", ":").replace(".", ":")
        if ":" not in s:
            return None
        try:
            parts = s.split(":")
            return (int(parts[0]), int(parts[1]))
        except Exception:
            return None

    def _save(self):
        # Validation minimale
        numof = self.f_numof.get().strip()
        pilote = self.f_pilote.get()

        if not numof:
            messagebox.showerror("Champ obligatoire", "Le N° d'OF est obligatoire.")
            return
        if pilote == "Sélectionner…":
            messagebox.showerror("Champ obligatoire", "Le pilote est obligatoire.")
            return

        # Date
        date_str = self.f_date.get().strip()
        try:
            d_obj = datetime.strptime(date_str, "%d/%m/%Y").date()
        except ValueError:
            messagebox.showerror("Date invalide", "Format attendu : jj/mm/aaaa")
            return

        # Construire la déclaration
        decl = {h: "" for h in EXCEL_HEADERS}
        decl["Type"] = "production"
        decl["Date"] = d_obj
        decl["Heure"] = datetime.now().time().replace(microsecond=0)
        decl["Ligne"] = self.f_ligne.get()
        decl["Pilote"] = pilote
        decl["Co-pilote"] = self.f_copilote.get() if self.f_copilote.get() != "Sélectionner…" else ""
        decl["N° OF"] = numof
        decl["Poste"] = self.f_poste.get()
        decl["Code produit"] = self.f_codeprod.get()
        decl["Taille"] = self.f_taille.get() if self.f_taille.get() != "Sélectionner…" else ""
        decl["Code fibre"] = self.f_fibre.get() if self.f_fibre.get() != "Sélectionner…" else ""
        decl["Poids garn. (g)"] = self._to_int(self.f_poids.get())
        decl["Réf. taie"] = self.f_reftaie.get()
        decl["Nb pers."] = self._to_int(self.f_nbpers.get())

        hd = self._parse_time(self.f_hdebut.get())
        hf = self._parse_time(self.f_hfin.get())
        if hd: decl["Heure début"] = time(hd[0], hd[1])
        if hf: decl["Heure fin"] = time(hf[0], hf[1])

        decl["Durée OF (min)"] = self.f_duree_lbl._auto_value or 0
        decl["Qté fab."] = self._to_int(self.f_qtefab.get())
        decl["Qté emb."] = self._to_int(self.f_qteemb.get())
        decl["2nd choix"] = self._to_int(self.f_2nd.get())
        decl["Manq. taies"] = self._to_int(self.f_manq_taies.get())
        decl["Manq. housse"] = self._to_int(self.f_manq_housse.get())
        decl["Déf. couture"] = self._to_int(self.f_def_couture.get())

        # Rattrapages
        nb_r, dur_r, det_r = 0, 0, []
        for r in self.rattrap_rows:
            t = r["type"].get()
            if t and t != "Sélectionner…":
                nb_r += 1
                d = self._to_int(r["duree"].get())
                dur_r += d
                det_r.append(f"{t} ({d}min){' - ' + r['comment'].get() if r['comment'].get() else ''}")
        decl["Nb rattrap."] = nb_r
        decl["Durée rattrap. (min)"] = dur_r
        decl["Détail rattrapages"] = "; ".join(det_r)

        # Problèmes
        nb_p, dur_p, det_p = 0, 0, []
        for r in self.probleme_rows:
            t = r["type"].get()
            if t and t != "Sélectionner…":
                nb_p += 1
                d = self._to_int(r["duree"].get())
                dur_p += d
                det_p.append(f"{t} ({d}min){' - ' + r['comment'].get() if r['comment'].get() else ''}")
        decl["Nb pb tech."] = nb_p
        decl["Durée pb tech. (min)"] = dur_p
        decl["Détail pb tech."] = "; ".join(det_p)

        decl["Commentaire"] = self.f_comment.get("1.0", "end").strip()

        # Sauvegarder
        if self.store.add(decl):
            messagebox.showinfo("Enregistré", "La déclaration de production a été enregistrée.")
            if self.on_save:
                self.on_save()
            self.destroy()

    def _to_int(self, s):
        try:
            return int(float(s.strip())) if s and s.strip() else 0
        except (ValueError, AttributeError):
            return 0


# ===========================================================================
# FENÊTRE FORMULAIRE PANNE
# ===========================================================================

class FormPanne(ctk.CTkToplevel):

    def __init__(self, parent, store, on_save=None):
        super().__init__(parent)
        self.store = store
        self.on_save = on_save
        self.title("Déclarer une panne")
        self.geometry("800x680")
        self.configure(fg_color=COLORS["gris_bg"])
        self.transient(parent)
        self.grab_set()

        self._build()

    def _build(self):
        head = ctk.CTkFrame(self, fg_color=COLORS["orange"], corner_radius=0, height=72)
        head.pack(fill="x")
        head.pack_propagate(False)

        ctk.CTkLabel(head, text="⚡   Déclarer une panne / maintenance",
                     font=("Segoe UI", 18, "bold"),
                     text_color=COLORS["blanc"], anchor="w"
                     ).pack(side="left", padx=24)

        scroll = ctk.CTkScrollableFrame(self, fg_color=COLORS["gris_bg"], corner_radius=0)
        scroll.pack(fill="both", expand=True)

        body = ctk.CTkFrame(scroll, fg_color=COLORS["blanc"],
                            corner_radius=12, border_width=1,
                            border_color=COLORS["gris_bord"])
        body.pack(fill="x", padx=20, pady=20)

        # Section 1 : Contexte
        FormProduction._section(self, body, "1", "📅", "Contexte")

        row1 = ctk.CTkFrame(body, fg_color="transparent")
        row1.pack(fill="x", padx=20, pady=(0, 12))
        self.f_date = FormProduction._field(self, row1, "Date *", width=200,
                                             default=date.today().strftime("%d/%m/%Y"))
        self.f_ligne = FormProduction._dropdown(self, row1, "Ligne *", LISTES["lignes"],
                                                 width=200, default="ORC1")
        self.f_poste = FormProduction._dropdown(self, row1, "Poste", LISTES["postes"],
                                                 width=200, default="Matin")

        FormProduction._divider(self, body)

        # Section 2 : Détails
        FormProduction._section(self, body, "2", "⚡", "Détails de l'intervention")

        row2 = ctk.CTkFrame(body, fg_color="transparent")
        row2.pack(fill="x", padx=20, pady=(0, 12))
        self.f_hdebut = FormProduction._field(self, row2, "Heure début (hh:mm)", width=180)
        self.f_hfin = FormProduction._field(self, row2, "Heure fin (hh:mm)", width=180)
        self.f_type = FormProduction._dropdown(self, row2, "Type",
                                                LISTES["types_panne"], width=240,
                                                default="Panne")

        row3 = ctk.CTkFrame(body, fg_color="transparent")
        row3.pack(fill="x", padx=20, pady=(0, 12))
        self.f_equip = FormProduction._dropdown(self, row3, "Équipement concerné *",
                                                 LISTES["equipements"], width=300)
        self.f_intervenant = FormProduction._dropdown(self, row3, "Intervenant",
                                                       LISTES["intervenants"], width=300)

        row4 = ctk.CTkFrame(body, fg_color="transparent")
        row4.pack(fill="x", padx=20, pady=(0, 12))
        self.f_detail = FormProduction._field(self, row4, "Nature / détail de l'intervention",
                                               width=600, expand=True)

        FormProduction._divider(self, body)

        # Section 3 : Commentaire
        FormProduction._section(self, body, "3", "💬", "Commentaire général", optional=True)

        self.f_comment = ctk.CTkTextbox(body, height=100,
                                        font=("Segoe UI", 12),
                                        fg_color=COLORS["blanc"],
                                        border_color=COLORS["gris_bord"],
                                        border_width=1, corner_radius=6)
        self.f_comment.pack(fill="x", padx=20, pady=(0, 20))

        # Footer
        footer = ctk.CTkFrame(self, fg_color=COLORS["gris_panel"], corner_radius=0, height=72)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        ctk.CTkLabel(footer, text="Champs * obligatoires",
                     font=("Segoe UI", 11),
                     text_color=COLORS["gris_label"]
                     ).pack(side="left", padx=20)

        ctk.CTkButton(footer, text="✓   Enregistrer la panne",
                      font=("Segoe UI", 13, "bold"),
                      fg_color=COLORS["orange"], hover_color="#C2410C",
                      text_color=COLORS["blanc"],
                      corner_radius=8, height=44, width=220,
                      command=self._save).pack(side="right", padx=20, pady=14)

        ctk.CTkButton(footer, text="Annuler",
                      font=("Segoe UI", 13),
                      fg_color=COLORS["blanc"], hover_color=COLORS["gris_bg"],
                      text_color=COLORS["gris_label"],
                      border_width=1, border_color=COLORS["gris_bord"],
                      corner_radius=8, height=44, width=120,
                      command=self.destroy).pack(side="right", pady=14)

    def _save(self):
        equip = self.f_equip.get()
        if equip == "Sélectionner…" or not equip:
            messagebox.showerror("Champ obligatoire", "L'équipement concerné est obligatoire.")
            return

        date_str = self.f_date.get().strip()
        try:
            d_obj = datetime.strptime(date_str, "%d/%m/%Y").date()
        except ValueError:
            messagebox.showerror("Date invalide", "Format attendu : jj/mm/aaaa")
            return

        # Calcul durée
        duree = 0
        hd_str = self.f_hdebut.get()
        hf_str = self.f_hfin.get()
        try:
            parts_d = hd_str.replace("h", ":").split(":")
            parts_f = hf_str.replace("h", ":").split(":")
            if len(parts_d) >= 2 and len(parts_f) >= 2:
                duree = (int(parts_f[0])*60 + int(parts_f[1])) - (int(parts_d[0])*60 + int(parts_d[1]))
                if duree < 0:
                    duree += 24*60
        except (ValueError, IndexError):
            pass

        decl = {h: "" for h in EXCEL_HEADERS}
        decl["Type"] = "panne"
        decl["Date"] = d_obj
        decl["Heure"] = datetime.now().time().replace(microsecond=0)
        decl["Ligne"] = self.f_ligne.get()
        decl["Poste"] = self.f_poste.get()
        decl["Pilote"] = self.f_intervenant.get() if self.f_intervenant.get() != "Sélectionner…" else ""
        decl["N° OF"] = "—"
        decl["Durée pannes (min)"] = duree

        if hd_str:
            try:
                p = hd_str.replace("h", ":").split(":")
                decl["Heure début"] = time(int(p[0]), int(p[1]))
            except (ValueError, IndexError):
                pass
        if hf_str:
            try:
                p = hf_str.replace("h", ":").split(":")
                decl["Heure fin"] = time(int(p[0]), int(p[1]))
            except (ValueError, IndexError):
                pass

        decl["Nb pannes"] = 1
        decl["Équipement"] = equip
        decl["Type panne"] = self.f_type.get()
        decl["Intervenant"] = self.f_intervenant.get() if self.f_intervenant.get() != "Sélectionner…" else ""
        decl["Détail panne"] = self.f_detail.get()
        decl["Commentaire"] = self.f_comment.get("1.0", "end").strip()

        if self.store.add(decl):
            messagebox.showinfo("Enregistré", "La déclaration de panne a été enregistrée.")
            if self.on_save:
                self.on_save()
            self.destroy()


# ===========================================================================
# POINT D'ENTRÉE
# ===========================================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
