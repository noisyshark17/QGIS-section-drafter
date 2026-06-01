"""
custom_additions.py
====================
Custom extensions for the QGIS section drafter plugin.

This file is NOT shipped with the plugin and will NOT be overwritten
by a QGIS auto-update.  main_dialog.py calls `apply_custom_patches()`
at the bottom of the file — that single call is the only thing that
will be lost if QGIS overwrites main_dialog.py.

Everything in this file is permanent.

Features added
──────────────
  1.  "Matplotlib Window" output option — each plot opens in a floating
      matplotlib window via plt.show(block=False).

  2.  Section azimuth spin box — added to the Plot Settings group so
      the user can enter the azimuth of the cross-section bearing.

  3.  _build_structural_projections() — calculate apparent dip.
      Formula: tan(app_dip) = tan(true_dip) × cos(dip_dir − section_az)

  4.  _plot_structure_on_section() — detect Struct_Dist / DIP / DIP_DIR
      columns in the Excel and draw apparent-dip tick marks on the section.

  NOTE: Drill hole trace plotting (feature 5 in earlier versions) was
  removed — it is now a native plugin feature in the Plotting tab under
  the "Drill Hole Overlay" group.  Set up a collar CSV there instead.

How to re-apply after a plugin update
──────────────────────────────────────
If QGIS overwrites main_dialog.py and the features disappear again,
just add these three lines back at the very end of main_dialog.py:

    try:
        from .custom_additions import apply_custom_patches
        apply_custom_patches(CombinedGeospatialToolDialog)
    except Exception as e:
        print(f"custom_additions: {e}")
"""

import math
import re

import pandas as pd
from qgis.core import QgsMessageLog, Qgis

try:
    import matplotlib.pyplot as plt
    _MPL_AVAILABLE = True
except ImportError:
    _MPL_AVAILABLE = False

from qgis.PyQt.QtWidgets import QHBoxLayout, QLabel, QDoubleSpinBox


# ── UI injection ──────────────────────────────────────────────────────────────

def _inject_ui(self):
    """
    Called once after setup_plotting_tab() to add our custom UI elements:
      • 'Matplotlib Window' item in the output combo
      • Section azimuth spin box below the display label
    Safe to call multiple times — checks for duplicates first.
    """
    # 1. Add Matplotlib Window combo option
    if hasattr(self, 'plot_output_combo'):
        items = [self.plot_output_combo.itemText(i)
                 for i in range(self.plot_output_combo.count())]
        if 'Matplotlib Window' not in items:
            self.plot_output_combo.addItem('Matplotlib Window')

    # 2. Add section azimuth spin box (only once)
    if hasattr(self, 'section_azimuth_spin'):
        return   # already injected

    if not hasattr(self, 'plot_output_combo'):
        return   # UI not built yet

    # Find the settings_group layout by walking the plot tab widget tree
    from qgis.PyQt.QtWidgets import QGroupBox, QVBoxLayout
    spin_parent_layout = None
    try:
        for group_box in self.findChildren(QGroupBox):
            if group_box.title() == "Plot Settings":
                spin_parent_layout = group_box.layout()
                break
    except Exception:
        pass

    if spin_parent_layout is None:
        return

    az_layout = QHBoxLayout()
    az_layout.addWidget(QLabel("Section azimuth (°):"))
    self.section_azimuth_spin = QDoubleSpinBox()
    self.section_azimuth_spin.setRange(0, 360)
    self.section_azimuth_spin.setValue(0.0)
    self.section_azimuth_spin.setDecimals(1)
    self.section_azimuth_spin.setSuffix("  °")
    self.section_azimuth_spin.setToolTip(
        "Azimuth of the cross-section (degrees from north, clockwise).\n"
        "Used to project true structural dip onto the section plane.\n"
        "Columns needed in Excel:  Struct_Dist, Struct_Dip, Struct_DipDir"
    )
    az_layout.addWidget(self.section_azimuth_spin)
    az_layout.addStretch()
    spin_parent_layout.addLayout(az_layout)


# ── Matplotlib Window output handler ─────────────────────────────────────────

def _handle_matplotlib_window_output(self, figures):
    """
    Open each figure in a floating matplotlib window.
    Replaces the normal in-dialog display when 'Matplotlib Window' is selected.
    """
    if not _MPL_AVAILABLE:
        return
    for fig in figures:
        plt.figure(fig.number)
        plt.show(block=False)


# ── Structural dip projection ─────────────────────────────────────────────────

def _build_structural_projections(self, true_dip_deg, dip_dir_deg, section_az_deg):
    """
    Return the apparent dip angle (degrees) of a structural measurement
    projected onto a cross-section with the given azimuth.

    Formula:  tan(apparent_dip) = tan(true_dip) × cos(dip_direction − section_azimuth)

    Parameters
    ----------
    true_dip_deg   : true dip (degrees, positive downward)
    dip_dir_deg    : dip direction / azimuth (degrees from north, clockwise)
    section_az_deg : azimuth of the cross-section (degrees from north, clockwise)

    Returns
    -------
    float : apparent dip in degrees (positive = dipping in the section direction)
    """
    td   = math.radians(abs(true_dip_deg))
    beta = math.radians(dip_dir_deg - section_az_deg)
    app  = math.atan(math.tan(td) * math.cos(beta))
    return math.degrees(app)


def _plot_structure_on_section(self, ax, df):
    """
    Detect structural measurement columns in the dataframe and draw
    apparent-dip tick marks on the cross-section axes.

    Expected column name patterns (case-insensitive):
      Distance  : Struct_Dist, Structure_Distance, Struct_Distance_N …
      Dip       : Struct_Dip,  Dip,   DIP,  DIP_N …
      Dip dir   : Struct_DipDir, DipDir, DIP_DIR, Azimuth, Azimuth_N …
    """
    cols = list(df.columns)

    # Skip entirely if no structural-looking columns present
    has_struct = any(
        re.search(r'(?i)(struct|\bdip\b|dip_dir|dipdir|azimuth)', c)
        for c in cols
    )
    if not has_struct:
        return

    section_az_deg = 0.0
    spin = getattr(self, 'section_azimuth_spin', None)
    if spin is not None:
        section_az_deg = spin.value()

    # Detect column groups (distance / dip / dip-direction)
    dist_pat   = re.compile(r'(?i)^(struct(?:ure)?[-_]?dist(?:ance)?[-_]?\d*|struct[-_]?d[-_]?\d*)$')
    dip_pat    = re.compile(r'(?i)^(struct(?:ure)?[-_]?dip[-_]?\d*|\bdip[-_]?\d*)$')
    dipdir_pat = re.compile(r'(?i)^(struct(?:ure)?[-_]?dip[-_]?dir[-_]?\d*|dipdir[-_]?\d*|azimuth[-_]?\d*)$')

    dist_cols   = [c for c in cols if dist_pat.match(c)]
    dip_cols    = [c for c in cols if dip_pat.match(c)]
    dipdir_cols = [c for c in cols if dipdir_pat.match(c)]

    if not (dist_cols and dip_cols and dipdir_cols):
        return

    tick_color = '#7B1FA2'
    tick_len_m = 15          # half-length of the tick mark in metres

    for dist_col, dip_col, dir_col in zip(dist_cols, dip_cols, dipdir_cols):
        mask = df[dist_col].notna() & df[dip_col].notna() & df[dir_col].notna()
        if not mask.any():
            continue

        added_label = False
        for _, row in df[mask].iterrows():
            try:
                app_dip = _build_structural_projections(
                    self,
                    float(row[dip_col]),
                    float(row[dir_col]),
                    section_az_deg,
                )
            except (ValueError, TypeError):
                continue

            x0 = float(row[dist_col])

            # Estimate elevation at x0 from the nearest elevation column
            y0 = None
            for col in cols:
                if re.search(r'(?i)(elev|height|rl)', col) and df[col].notna().any():
                    idx_nearest = (df[dist_col] - x0).abs().argsort().iloc[0]
                    y0 = float(df[col].iloc[idx_nearest])
                    break
            if y0 is None:
                continue

            dx = tick_len_m * math.cos(math.radians(app_dip))
            dy = tick_len_m * math.sin(math.radians(app_dip))
            label = 'Structural dip' if not added_label else '_nolegend_'
            ax.annotate(
                '', xy=(x0 + dx, y0 + dy), xytext=(x0 - dx, y0 - dy),
                arrowprops=dict(arrowstyle='-', color=tick_color, lw=1.5),
                zorder=6,
            )
            ax.plot(x0, y0, 's', color=tick_color, markersize=5,
                    label=label, zorder=6,
                    markeredgecolor='black', markeredgewidth=0.5)
            added_label = True

        if added_label:
            QgsMessageLog.logMessage(
                f"Structural dip: {dist_col}/{dip_col}/{dir_col} "
                f"at section az={section_az_deg:.0f}°",
                "IntegratedProfileAnalyzer", Qgis.Info
            )


# NOTE: _plot_drill_traces_on_section was removed — drill hole plotting is now
# a native feature of the plugin (Plotting tab → Drill Hole Overlay group).


# ── Patch wrapper for Matplotlib Window output ────────────────────────────────

def _patched_setup_plotting_tab(original_fn):
    """Wrap setup_plotting_tab to call _inject_ui() after the tab is built."""
    def wrapper(self):
        original_fn(self)
        try:
            _inject_ui(self)
        except Exception as e:
            QgsMessageLog.logMessage(
                f"custom_additions: UI injection failed: {e}",
                "IntegratedProfileAnalyzer", Qgis.Warning
            )
    return wrapper


# ── Main entry point ──────────────────────────────────────────────────────────

def apply_custom_patches(cls):
    """
    Monkey-patch all custom additions onto cls (CombinedGeospatialToolDialog).
    Call this once at the bottom of main_dialog.py.
    """
    # Structural methods
    cls._build_structural_projections = _build_structural_projections
    cls._plot_structure_on_section     = _plot_structure_on_section

    # Wrap setup_plotting_tab to inject UI after it runs
    if hasattr(cls, 'setup_plotting_tab'):
        cls.setup_plotting_tab = _patched_setup_plotting_tab(cls.setup_plotting_tab)

    # Wrap generate_plots to handle Matplotlib Window output mode
    if hasattr(cls, 'generate_plots'):
        original_generate = cls.generate_plots

        def patched_generate_plots(self):
            original_generate(self)
            # After the original runs, if Matplotlib Window was selected,
            # the figures are already in the display widget — call plt.show
            output_type = getattr(self, 'plot_output_combo', None)
            if output_type and output_type.currentText() == 'Matplotlib Window':
                if _MPL_AVAILABLE:
                    plt.show(block=False)

        cls.generate_plots = patched_generate_plots

    # Ensure _plot_structure_on_section and _plot_drill_traces_on_section
    # are called in the plot loop by patching _plot_marker_columns.
    if hasattr(cls, '_plot_marker_columns'):
        original_markers = cls._plot_marker_columns

        def patched_plot_marker_columns(self, ax, df):
            original_markers(self, ax, df)
            try:
                _plot_structure_on_section(self, ax, df)
            except Exception as e:
                QgsMessageLog.logMessage(
                    f"custom_additions: structure plot error: {e}",
                    "IntegratedProfileAnalyzer", Qgis.Warning
                )
        cls._plot_marker_columns = patched_plot_marker_columns

    QgsMessageLog.logMessage(
        "custom_additions: patches applied successfully.",
        "IntegratedProfileAnalyzer", Qgis.Info
    )
