#!/usr/bin/python3
# Excel for Raspberry Pi: a Windows 11 style spreadsheet that opens and saves .xlsx files.
import os
import sys
import re
import json
import math
import time
import shutil
import signal
import threading
import subprocess

import gi
gi.require_version("Gtk", "3.0")
gi.require_version("Gdk", "3.0")
gi.require_version("GdkPixbuf", "2.0")
from gi.repository import Gtk, Gdk, GLib, GdkPixbuf, Pango, Gio, GObject  # noqa: E402,F401
import cairo  # noqa: E402

HOME = os.path.expanduser("~")
C_ACCENT = "#4CC2FF"
def hexc(h, a=1.0):
    h = h.lstrip("#")
    return (int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255, a)


def _rr(cr, x, y, w, h, r):
    r = min(r, w / 2, h / 2)
    cr.new_sub_path()
    cr.arc(x + w - r, y + r, r, -math.pi / 2, 0)
    cr.arc(x + w - r, y + h - r, r, 0, math.pi / 2)
    cr.arc(x + r, y + h - r, r, math.pi / 2, math.pi)
    cr.arc(x + r, y + r, r, math.pi, 1.5 * math.pi)
    cr.close_path()


def _line(cr, *pts):
    cr.move_to(*pts[0])
    for p in pts[1:]:
        cr.line_to(*p)
    cr.stroke()


def draw_glyph(cr, name, s, col):
    """Monochrome line icons for buttons."""
    cr.save()
    cr.scale(s / 16.0, s / 16.0)
    cr.set_source_rgba(*col)
    cr.set_line_width(1.15)
    cr.set_line_cap(cairo.LINE_CAP_ROUND)
    cr.set_line_join(cairo.LINE_JOIN_ROUND)
    pi = math.pi
    L = lambda *p: _line(cr, *p)  # noqa: E731
    if name == "back":
        L((13.5, 8), (2.5, 8))
        L((7, 3.5), (2.5, 8), (7, 12.5))
    elif name == "forward":
        L((2.5, 8), (13.5, 8))
        L((9, 3.5), (13.5, 8), (9, 12.5))
    elif name == "up":
        L((8, 13.5), (8, 2.5))
        L((3.5, 7), (8, 2.5), (12.5, 7))
    elif name == "refresh":
        cr.arc(8, 8, 5.5, -pi * 0.35, pi * 1.55)
        cr.stroke()
        L((12.9, 1.8), (12.6, 4.9), (9.6, 4.6))
    elif name == "plus":
        L((8, 2.5), (8, 13.5))
        L((2.5, 8), (13.5, 8))
    elif name == "new":
        cr.set_source_rgba(*hexc(C_ACCENT))
        cr.arc(8, 8, 6.5, 0, 2 * pi)
        cr.stroke()
        L((8, 4.8), (8, 11.2))
        L((4.8, 8), (11.2, 8))
    elif name == "cut":
        cr.arc(4.5, 12, 2, 0, 2 * pi)
        cr.stroke()
        cr.arc(11.5, 12, 2, 0, 2 * pi)
        cr.stroke()
        L((5.8, 10.5), (11, 1.8))
        L((10.2, 10.5), (5, 1.8))
    elif name == "copy":
        _rr(cr, 5, 1.5, 9, 10.5, 1.5)
        cr.stroke()
        cr.move_to(3, 4.5)
        cr.line_to(3, 13)
        cr.arc_negative(4.5, 13, 1.5, pi, pi / 2)
        cr.line_to(11, 14.5)
        cr.stroke()
    elif name == "paste":
        _rr(cr, 2.5, 3, 11, 11.5, 1.5)
        cr.stroke()
        _rr(cr, 5.5, 1.5, 5, 3, 1)
        cr.stroke()
        L((5.5, 8), (10.5, 8))
        L((5.5, 11), (9, 11))
    elif name == "rename":
        _rr(cr, 1.5, 4.5, 13, 7, 1.3)
        cr.stroke()
        L((4, 8), (8, 8))
        L((10.5, 2.5), (10.5, 13.5))
        L((9, 2.5), (12, 2.5))
        L((9, 13.5), (12, 13.5))
    elif name == "delete" or name == "trash":
        L((2.5, 4), (13.5, 4))
        L((6, 4), (6.5, 2), (9.5, 2), (10, 4))
        cr.move_to(4, 4)
        cr.line_to(4.8, 14)
        cr.line_to(11.2, 14)
        cr.line_to(12, 4)
        cr.stroke()
        L((6.8, 6.8), (7, 11.5))
        L((9.2, 6.8), (9, 11.5))
    elif name == "sort":
        L((4.5, 2.5), (4.5, 13.5))
        L((2, 5), (4.5, 2.5), (7, 5))
        L((11.5, 2.5), (11.5, 13.5))
        L((9, 11), (11.5, 13.5), (14, 11))
    elif name == "view":
        for y in (3, 8, 13):
            L((2, y), (14, y))
    elif name == "more":
        for x in (3, 8, 13):
            cr.arc(x, 8, 1.1, 0, 2 * pi)
            cr.fill()
    elif name == "chevdown":
        L((4, 6), (8, 10), (12, 6))
    elif name == "chevright":
        L((6, 4), (10, 8), (6, 12))
    elif name == "close":
        L((3.5, 3.5), (12.5, 12.5))
        L((12.5, 3.5), (3.5, 12.5))
    elif name == "min":
        L((3, 8), (13, 8))
    elif name == "max":
        _rr(cr, 3, 3, 10, 10, 1.2)
        cr.stroke()
    elif name == "restore":
        _rr(cr, 3, 5, 8, 8, 1.2)
        cr.stroke()
        L((5, 3), (12, 3), (13, 4), (13, 11))
    elif name == "search":
        cr.arc(6.6, 6.6, 4.6, 0, 2 * pi)
        cr.stroke()
        L((10, 10), (14, 14))
    elif name == "pane":
        _rr(cr, 1.5, 2.5, 13, 11, 1.5)
        cr.stroke()
        L((10, 2.5), (10, 13.5))
    elif name == "terminal":
        _rr(cr, 1.5, 2.5, 13, 11, 1.5)
        cr.stroke()
        L((4, 6), (6.5, 8), (4, 10))
        L((8, 10.5), (11.5, 10.5))
    elif name == "properties":
        _rr(cr, 2.5, 1.5, 11, 13, 1.5)
        cr.stroke()
        for y in (5, 8, 11):
            cr.arc(5.2, y, 0.8, 0, 2 * pi)
            cr.fill()
            L((7.2, y), (11, y))
    elif name == "pin":
        cr.move_to(6, 2)
        cr.line_to(10, 2)
        cr.line_to(9.5, 6.5)
        cr.line_to(12, 9)
        cr.line_to(4, 9)
        cr.line_to(6.5, 6.5)
        cr.close_path()
        cr.stroke()
        L((8, 9), (8, 14))
    elif name == "pin-fill":
        cr.set_line_width(1.0)
        cr.translate(8, 8)
        cr.rotate(pi / 4)
        cr.translate(-8, -8)
        cr.move_to(6, 2)
        cr.line_to(10, 2)
        cr.line_to(9.5, 6.5)
        cr.line_to(12, 9)
        cr.line_to(4, 9)
        cr.line_to(6.5, 6.5)
        cr.close_path()
        cr.stroke()
        L((8, 9), (8, 14))
    elif name == "folder":
        cr.move_to(1.5, 4)
        cr.line_to(6, 4)
        cr.line_to(7.5, 5.5)
        cr.line_to(14.5, 5.5)
        cr.line_to(14.5, 13)
        cr.line_to(1.5, 13)
        cr.close_path()
        cr.stroke()
    elif name == "file":
        cr.move_to(3.5, 1.5)
        cr.line_to(9.5, 1.5)
        cr.line_to(12.5, 4.5)
        cr.line_to(12.5, 14.5)
        cr.line_to(3.5, 14.5)
        cr.close_path()
        cr.stroke()
        L((9.5, 1.5), (9.5, 4.5), (12.5, 4.5))
    elif name == "zip":
        _rr(cr, 2.5, 1.5, 11, 13, 1.5)
        cr.stroke()
        for y in (3, 5, 7):
            L((7, y), (9, y))
        _rr(cr, 6.5, 9, 3, 3.5, 0.6)
        cr.stroke()
    elif name == "extract":
        _rr(cr, 2.5, 6.5, 11, 8, 1.5)
        cr.stroke()
        L((8, 1.5), (8, 10))
        L((5.5, 7.5), (8, 10), (10.5, 7.5))
    elif name == "restoreitem":
        cr.arc(8.5, 8.5, 5, -pi * 0.9, pi * 0.6)
        cr.stroke()
        L((2, 3.5), (3.6, 6.6), (6.6, 5.4))
    elif name == "open":
        _rr(cr, 1.5, 1.5, 13, 13, 2)
        cr.stroke()
        L((7, 9), (12.5, 3.5))
        L((8.5, 3.5), (12.5, 3.5), (12.5, 7.5))
    elif name == "newtab":
        _rr(cr, 1.5, 3.5, 13, 10, 1.5)
        cr.stroke()
        L((1.5, 6), (8, 6))
        L((8, 3.5), (8, 6))
    elif name == "check":
        L((3, 8.5), (6.5, 12), (13, 4.5))
    elif name == "eject":
        cr.move_to(8, 3)
        cr.line_to(13, 9)
        cr.line_to(3, 9)
        cr.close_path()
        cr.stroke()
        L((3, 12.5), (13, 12.5))
    elif name == "details-view":
        for y in (3.5, 8, 12.5):
            _rr(cr, 1.5, y - 1.2, 2.4, 2.4, 0.5)
            cr.stroke()
            L((6, y), (14.5, y))
    elif name == "icons-view":
        for x in (2, 9):
            for y in (2, 9):
                _rr(cr, x, y, 5, 5, 1.2)
                cr.stroke()
    elif name == "selectall":
        _rr(cr, 2, 2, 12, 12, 1.5)
        cr.stroke()
        L((5, 8), (7.2, 10.2), (11, 6))
    elif name == "info":
        cr.arc(8, 8, 6, 0, 2 * pi)
        cr.stroke()
        L((8, 7), (8, 11.5))
        cr.arc(8, 4.8, 0.9, 0, 2 * pi)
        cr.fill()
    elif name == "share":
        cr.arc(4, 8, 1.8, 0, 2 * pi)
        cr.stroke()
        cr.arc(12, 3.8, 1.8, 0, 2 * pi)
        cr.stroke()
        cr.arc(12, 12.2, 1.8, 0, 2 * pi)
        cr.stroke()
        L((5.6, 7.1), (10.4, 4.7))
        L((5.6, 8.9), (10.4, 11.3))
    elif name == "network":
        _rr(cr, 5, 1.5, 6, 4.5, 1)
        cr.stroke()
        L((8, 6), (8, 9))
        L((3, 9), (13, 9))
        L((3, 9), (3, 11))
        L((13, 9), (13, 11))
        _rr(cr, 1, 11, 4, 3.5, 0.8)
        cr.stroke()
        _rr(cr, 11, 11, 4, 3.5, 0.8)
        cr.stroke()
    elif name == "link":
        cr.arc(5.5, 10.5, 3, pi * 0.25, pi * 1.75)
        cr.stroke()
        cr.arc(10.5, 5.5, 3, pi * 1.25, pi * 0.75)
        cr.stroke()
        L((6.5, 9.5), (9.5, 6.5))
    else:
        draw_glyph_extra(cr, name, L, pi)
    cr.restore()




# ----------------------------------------------------------------------------
# Shared Windows 11 dark style
# ----------------------------------------------------------------------------
BASE_CSS = """
* { font-family: "Segoe UI Variable Text", "Segoe UI", "Selawik", "Open Sans", "Noto Sans", sans-serif; }
window.fx, window.fx-dialog { background-color: #202020; color: #ffffff; font-size: 12px; }
.fx label { color: #ffffff; }
headerbar.fx-title { background: #202020; border: none; box-shadow: none; min-height: 40px; padding: 0; margin: 0; }
headerbar.fx-title:backdrop { background: #202020; }
.app-title { font-size: 12px; color: #ffffff; margin-left: 4px; }
.app-title:backdrop { color: rgba(255,255,255,0.55); }
button.wc { background: transparent; border: none; border-radius: 0; min-width: 46px; min-height: 32px; padding: 0; box-shadow: none; }
button.wc:hover { background-color: rgba(255,255,255,0.06); }
button.wc.close:hover { background-color: #C42B1C; }
.tabstrip { margin-top: 6px; margin-left: 6px; }
button.tab { outline-color: transparent; background: transparent; border: none; border-radius: 8px 8px 0 0; min-height: 34px;
   padding: 0 6px 0 12px; box-shadow: none; margin-right: 1px; }
button.tab:hover { background-color: rgba(255,255,255,0.04); }
button.tab.active { background-color: #2C2C2C; }
button.tab label { font-size: 12px; }
button.tab-x { background: transparent; border: none; border-radius: 4px; min-width: 24px; min-height: 24px; padding: 0; box-shadow: none; }
button.tab-x:hover { background-color: rgba(255,255,255,0.08); }
button.tab-new { background: transparent; border: none; border-radius: 4px; min-width: 32px; min-height: 28px; padding: 0;
   box-shadow: none; margin-top: 4px; margin-left: 2px; }
button.tab-new:hover { background-color: rgba(255,255,255,0.06); }
.bar { background-color: #2C2C2C; padding: 4px 8px; }
.bar-bottom { border-bottom: 1px solid rgba(255,255,255,0.06); }
button.tb, button.nb { background: transparent; border: none; border-radius: 4px; min-height: 32px; padding: 0 9px; box-shadow: none; color: #ffffff; }
button.nb { min-width: 34px; padding: 0; }
button.tb:hover, button.nb:hover { background-color: rgba(255,255,255,0.06); }
button.tb:active, button.nb:active { background-color: rgba(255,255,255,0.04); }
button.tb:disabled label { color: rgba(255,255,255,0.36); }
button.tb label { font-size: 12px; }
button.tb.on, button.nb.on { background-color: rgba(255,255,255,0.09); }
.vsep { background-color: rgba(255,255,255,0.08); min-width: 1px; margin: 6px 6px; }
.status { background-color: #202020; padding: 4px 12px; border-top: 1px solid rgba(255,255,255,0.06); }
.status label { font-size: 12px; color: rgba(255,255,255,0.786); }
.sec { color: rgba(255,255,255,0.786); }
.ter { color: rgba(255,255,255,0.545); }
.big-title { font-size: 20px; font-weight: 600; }
.sect-title { font-size: 14px; font-weight: 600; }
entry.field, entry.fx-search { background-color: rgba(255,255,255,0.06); color: #ffffff; border: 1px solid rgba(255,255,255,0.07);
   border-bottom: 1px solid rgba(255,255,255,0.54); border-radius: 4px; min-height: 30px; padding: 0 10px; box-shadow: none;
   font-size: 12px; caret-color: #ffffff; }
entry.field:focus, entry.fx-search:focus { background-color: #1c1c1c; border-bottom: 2px solid #4CC2FF; }
entry selection, textview text selection, entry selection:focus, textview text selection:focus { background-color: #0078D4; color: #ffffff; }
scrollbar { background: transparent; border: none; }
scrollbar slider { background-color: rgba(255,255,255,0.45); border-radius: 6px; min-width: 3px; min-height: 3px; border: none; }
scrollbar:hover slider { min-width: 6px; min-height: 6px; background-color: rgba(255,255,255,0.55); }
scrollbar trough { background: transparent; border: none; }
scale trough { background-color: rgba(255,255,255,0.3); border: none; border-radius: 2px; min-height: 4px; min-width: 4px; }
scale highlight { background-color: #4CC2FF; border: none; border-radius: 2px; min-height: 4px; min-width: 4px; }
scale slider { background-color: #4CC2FF; border: 5px solid #454545; border-radius: 50%; min-width: 12px; min-height: 12px; box-shadow: none; }
scale slider:hover { border-width: 4px; }
menu, .menu, .context-menu { background-color: #2C2C2C; color: #ffffff; border: 1px solid rgba(0,0,0,0.45); border-radius: 8px; padding: 4px; }
menu menuitem { padding: 6px 12px; border-radius: 4px; color: #ffffff; min-height: 18px; }
menu menuitem:hover { background-color: rgba(255,255,255,0.06); }
menu menuitem:disabled label { color: rgba(255,255,255,0.36); }
menu menuitem label { color: #ffffff; font-size: 12px; }
menu menuitem arrow { color: rgba(255,255,255,0.786); }
menu separator { background-color: rgba(255,255,255,0.08); margin: 4px 0; min-height: 1px; }
window.fx button, window.fx-dialog button { background-image: none; text-shadow: none; -gtk-icon-shadow: none; }
button.primary { background-color: #4CC2FF; color: #000000; border: none; border-radius: 4px; min-height: 30px; padding: 0 16px; box-shadow: none; }
button.primary label { color: #000000; }
button.primary:hover { background-color: #48B1E8; }
button.secondary { background-color: rgba(255,255,255,0.06); color: #ffffff; border: 1px solid rgba(255,255,255,0.09); border-radius: 4px; min-height: 30px; padding: 0 16px; box-shadow: none; }
button.secondary:hover { background-color: rgba(255,255,255,0.09); }
.dlg-body { background-color: #2B2B2B; padding: 24px; }
.dlg-foot { background-color: #202020; padding: 22px 24px; border-top: 1px solid rgba(0,0,0,0.3); }
.dlg-title { font-size: 20px; font-weight: 600; }
.fx-dialog label { color: #ffffff; }
.fx-dialog label.sec { color: rgba(255,255,255,0.786); }
.fx-dialog label.ter { color: rgba(255,255,255,0.545); }
checkbutton check { background-color: rgba(0,0,0,0.1); border: 1px solid rgba(255,255,255,0.6); border-radius: 4px; min-width: 18px; min-height: 18px; -gtk-icon-source: none; color: #000000; }
checkbutton check:checked { background-color: #4CC2FF; border-color: #4CC2FF; -gtk-icon-source: -gtk-icontheme("object-select-symbolic"); }
checkbutton label { color: #ffffff; font-size: 12px; margin-left: 6px; }
radiobutton radio { background-color: rgba(0,0,0,0.1); border: 1px solid rgba(255,255,255,0.6); border-radius: 50%; min-width: 18px; min-height: 18px; -gtk-icon-source: none; }
radiobutton radio:checked { background-color: #000000; border: 5px solid #4CC2FF; }
radiobutton label { color: #ffffff; font-size: 12px; margin-left: 6px; }
switch { background-color: transparent; border: 1px solid rgba(255,255,255,0.6); border-radius: 12px; min-width: 40px; min-height: 18px; }
switch:checked { background-color: #4CC2FF; border-color: #4CC2FF; }
switch slider { background-color: rgba(255,255,255,0.8); border-radius: 50%; min-width: 12px; min-height: 12px; margin: 3px; border: none; box-shadow: none; }
switch:checked slider { background-color: #000000; }
switch image { color: transparent; }
combobox button, combobox box { background-color: rgba(255,255,255,0.06); color: #ffffff; border: 1px solid rgba(255,255,255,0.07); border-radius: 4px; box-shadow: none; min-height: 30px; }
combobox button cellview, combobox button label, combobox arrow { color: #ffffff; }
spinbutton { background-color: rgba(255,255,255,0.06); color: #ffffff; border: 1px solid rgba(255,255,255,0.07); border-radius: 4px; box-shadow: none; }
spinbutton entry { background: transparent; border: none; box-shadow: none; color: #ffffff; }
spinbutton button { background: transparent; border: none; box-shadow: none; color: #ffffff; }
tooltip, tooltip.background { background-color: #2C2C2C; border: 1px solid rgba(0,0,0,0.4); border-radius: 4px; }
tooltip label { color: #ffffff; font-size: 12px; }
.set-group { background-color: rgba(255,255,255,0.05); border-radius: 4px; border: 1px solid rgba(0,0,0,0.2); padding: 12px 16px; }
.set-head { font-size: 13px; }
.set-desc { font-size: 12px; color: rgba(255,255,255,0.6); }
"""


def apply_css(extra=""):
    prov = Gtk.CssProvider()
    prov.load_from_data((BASE_CSS + extra).encode())
    Gtk.StyleContext.add_provider_for_screen(Gdk.Screen.get_default(), prov,
                                             Gtk.STYLE_PROVIDER_PRIORITY_APPLICATION + 10)
    Gtk.Settings.get_default().set_property("gtk-application-prefer-dark-theme", True)


def draw_glyph_extra(cr, name, L, pi):
    if name == "save":
        _rr(cr, 2, 2, 12, 12, 1.5)
        cr.stroke()
        _rr(cr, 4.5, 2, 6, 4, 0.5)
        cr.stroke()
        _rr(cr, 4.5, 9, 7, 5, 0.5)
        cr.stroke()
    elif name == "undo":
        cr.move_to(3, 7)
        cr.curve_to(6, 3.5, 13, 4.5, 13, 10)
        cr.stroke()
        L((3, 3), (3, 7), (7, 7))
    elif name == "redo":
        cr.move_to(13, 7)
        cr.curve_to(10, 3.5, 3, 4.5, 3, 10)
        cr.stroke()
        L((13, 3), (13, 7), (9, 7))
    elif name == "replace":
        L((2, 5), (10, 5))
        L((7.5, 2.5), (10, 5), (7.5, 7.5))
        L((14, 11), (6, 11))
        L((8.5, 8.5), (6, 11), (8.5, 13.5))
    elif name == "zoomin":
        cr.arc(6.6, 6.6, 4.6, 0, 2 * pi)
        cr.stroke()
        L((10, 10), (14, 14))
        L((4.6, 6.6), (8.6, 6.6))
        L((6.6, 4.6), (6.6, 8.6))
    elif name == "zoomout":
        cr.arc(6.6, 6.6, 4.6, 0, 2 * pi)
        cr.stroke()
        L((10, 10), (14, 14))
        L((4.6, 6.6), (8.6, 6.6))
    elif name == "fit":
        L((2, 5.5), (2, 2), (5.5, 2))
        L((10.5, 2), (14, 2), (14, 5.5))
        L((14, 10.5), (14, 14), (10.5, 14))
        L((5.5, 14), (2, 14), (2, 10.5))
    elif name == "wrap":
        L((2, 3), (14, 3))
        cr.move_to(2, 8)
        cr.line_to(11.5, 8)
        cr.curve_to(15, 8, 15, 13, 11.5, 13)
        cr.line_to(8, 13)
        cr.stroke()
        L((9.5, 11.3), (8, 13), (9.5, 14.7))
        L((2, 13), (5, 13))
    elif name == "play":
        cr.move_to(4.5, 2.5)
        cr.line_to(13, 8)
        cr.line_to(4.5, 13.5)
        cr.close_path()
        cr.fill()
    elif name == "pause":
        _rr(cr, 3.5, 2.5, 3, 11, 0.8)
        cr.fill()
        _rr(cr, 9.5, 2.5, 3, 11, 0.8)
        cr.fill()
    elif name in ("next", "prev"):
        if name == "prev":
            cr.translate(16, 0)
            cr.scale(-1, 1)
        cr.move_to(3, 3)
        cr.line_to(10.5, 8)
        cr.line_to(3, 13)
        cr.close_path()
        cr.fill()
        _rr(cr, 11, 3, 2, 10, 0.6)
        cr.fill()
    elif name == "shuffle":
        L((2, 4), (5, 4), (11, 12), (14, 12))
        L((2, 12), (5, 12), (11, 4), (14, 4))
        L((12, 2), (14, 4), (12, 6))
        L((12, 10), (14, 12), (12, 14))
    elif name in ("repeat", "repeat1"):
        cr.move_to(3, 9)
        cr.line_to(3, 6)
        cr.arc(5, 6, 2, pi, 1.5 * pi)
        cr.line_to(13, 4)
        cr.stroke()
        L((11, 2), (13, 4), (11, 6))
        cr.move_to(13, 7)
        cr.line_to(13, 10)
        cr.arc(11, 10, 2, 0, 0.5 * pi)
        cr.line_to(3, 12)
        cr.stroke()
        L((5, 10), (3, 12), (5, 14))
        if name == "repeat1":
            cr.set_font_size(5)
            cr.move_to(6.8, 9.8)
            cr.show_text("1")
    elif name in ("volume", "mute"):
        cr.move_to(1.5, 6)
        cr.line_to(4.5, 6)
        cr.line_to(8, 2.5)
        cr.line_to(8, 13.5)
        cr.line_to(4.5, 10)
        cr.line_to(1.5, 10)
        cr.close_path()
        cr.stroke()
        if name == "mute":
            L((10.5, 6), (14.5, 10))
            L((14.5, 6), (10.5, 10))
        else:
            cr.arc(8, 8, 3.5, -pi / 3.5, pi / 3.5)
            cr.stroke()
            cr.arc(8, 8, 6.2, -pi / 3.5, pi / 3.5)
            cr.stroke()
    elif name == "fullscreen":
        L((2, 6), (2, 2), (6, 2))
        L((10, 2), (14, 2), (14, 6))
        L((14, 10), (14, 14), (10, 14))
        L((6, 14), (2, 14), (2, 10))
        L((2, 2), (6, 6))
        L((14, 14), (10, 10))
    elif name == "unfullscreen":
        L((6, 2), (6, 6), (2, 6))
        L((10, 2), (10, 6), (14, 6))
        L((10, 14), (10, 10), (14, 10))
        L((6, 14), (6, 10), (2, 10))
    elif name == "rotate":
        cr.arc(8, 9, 5, pi * 1.15, pi * 2.85)
        cr.stroke()
        L((2.5, 2.5), (3.4, 6.4), (7.2, 5.4))
    elif name == "heart":
        cr.move_to(8, 13.5)
        cr.curve_to(1, 8, 1.5, 2.5, 5, 2.5)
        cr.curve_to(6.5, 2.5, 7.5, 3.5, 8, 4.5)
        cr.curve_to(8.5, 3.5, 9.5, 2.5, 11, 2.5)
        cr.curve_to(14.5, 2.5, 15, 8, 8, 13.5)
        cr.stroke()
    elif name == "pen":
        L((2.5, 13.5), (3.2, 10.5), (11, 2.7), (13.3, 5), (5.5, 12.8), (2.5, 13.5))
        L((9.5, 4.2), (11.8, 6.5))
    elif name == "highlighter":
        L((3, 13.5), (13, 13.5))
        cr.move_to(5, 11)
        cr.line_to(11, 3)
        cr.line_to(13.5, 5)
        cr.line_to(7.5, 12.5)
        cr.line_to(5, 12.5)
        cr.close_path()
        cr.stroke()
    elif name == "eraser":
        cr.move_to(2.5, 10)
        cr.line_to(9, 3.5)
        cr.line_to(13.5, 8)
        cr.line_to(8, 13.5)
        cr.line_to(5.5, 13.5)
        cr.close_path()
        cr.stroke()
        L((6, 6.5), (10.5, 11))
        L((8, 13.5), (14, 13.5))
    elif name == "crop":
        L((4.5, 1.5), (4.5, 11.5), (14.5, 11.5))
        L((1.5, 4.5), (11.5, 4.5), (11.5, 14.5))
    elif name == "camera":
        _rr(cr, 1.5, 4, 13, 9.5, 1.6)
        cr.stroke()
        L((5, 4), (6, 2.5), (10, 2.5), (11, 4))
        cr.arc(8, 8.8, 2.6, 0, 2 * pi)
        cr.stroke()
    elif name == "timer":
        cr.arc(8, 9, 5.5, 0, 2 * pi)
        cr.stroke()
        L((8, 9), (8, 5.8))
        L((6.3, 1.8), (9.7, 1.8))
    elif name == "settings":
        n = 8
        cr.new_path()
        for i in range(n * 2):
            a = i * pi / n
            r = 6.6 if i % 2 == 0 else 5.0
            a1, a2 = a - pi / (n * 2.4), a + pi / (n * 2.4)
            cr.line_to(8 + r * math.cos(a1), 8 + r * math.sin(a1))
            cr.line_to(8 + r * math.cos(a2), 8 + r * math.sin(a2))
        cr.close_path()
        cr.stroke()
        cr.arc(8, 8, 2.2, 0, 2 * pi)
        cr.stroke()
    elif name == "history":
        cr.arc(8, 8, 6, pi * 1.05, pi * 2.85)
        cr.stroke()
        L((1.2, 6.2), (2.3, 8.6), (4.6, 7.3))
        L((8, 4.6), (8, 8.2), (10.5, 9.7))
    elif name == "image":
        _rr(cr, 1.5, 2.5, 13, 11, 1.5)
        cr.stroke()
        L((1.5, 11), (5.5, 7), (9, 10.5), (11, 8.5), (14.5, 12))
        cr.arc(10.5, 5.5, 1.2, 0, 2 * pi)
        cr.stroke()
    elif name == "music":
        cr.arc(4.5, 12, 2, 0, 2 * pi)
        cr.stroke()
        cr.arc(12, 10.5, 2, 0, 2 * pi)
        cr.stroke()
        L((6.5, 12), (6.5, 3.5), (14, 2), (14, 10.5))
    elif name == "video":
        _rr(cr, 1.5, 3.5, 9.5, 9, 1.5)
        cr.stroke()
        L((11, 6.5), (14.5, 4.5), (14.5, 11.5), (11, 9.5))
    elif name == "queue":
        for y in (3, 7):
            L((2, y), (14, y))
        L((2, 11), (8, 11))
        cr.move_to(10.5, 9)
        cr.line_to(14, 11.5)
        cr.line_to(10.5, 14)
        cr.close_path()
        cr.fill()
    elif name == "edit":
        L((2.5, 13.5), (3.2, 10.5), (11, 2.7), (13.3, 5), (5.5, 12.8), (2.5, 13.5))
    elif name == "slideshow":
        _rr(cr, 1.5, 2.5, 13, 9, 1.5)
        cr.stroke()
        cr.move_to(6.5, 5)
        cr.line_to(10, 7)
        cr.line_to(6.5, 9)
        cr.close_path()
        cr.fill()
        L((5, 14), (11, 14))
    elif name == "print":
        _rr(cr, 1.5, 6, 13, 6, 1.2)
        cr.stroke()
        L((4.5, 6), (4.5, 1.5), (11.5, 1.5), (11.5, 6))
        _rr(cr, 4.5, 9.5, 7, 5, 0.5)
        cr.stroke()
    elif name == "wallpaper":
        _rr(cr, 1.5, 2.5, 13, 9, 1.2)
        cr.stroke()
        L((1.5, 9.5), (5, 6), (8, 9), (10, 7), (14.5, 10.5))
        L((5.5, 14), (10.5, 14))
        L((8, 11.5), (8, 14))
    elif name == "list":
        for y in (3.5, 8, 12.5):
            cr.arc(2.5, y, 0.9, 0, 2 * pi)
            cr.fill()
            L((5.5, y), (14, y))
    elif name == "add":
        L((8, 2.5), (8, 13.5))
        L((2.5, 8), (13.5, 8))
    elif name == "split":
        _rr(cr, 1.5, 2.5, 13, 11, 1.5)
        cr.stroke()
        L((8, 2.5), (8, 13.5))
    elif name == "keyboard":
        _rr(cr, 1, 3.5, 14, 9, 1.5)
        cr.stroke()
        for x in (4, 6.7, 9.3, 12):
            cr.rectangle(x - 0.5, 6, 1, 1)
        cr.fill()
        L((4.5, 10), (11.5, 10))


class Icon(Gtk.DrawingArea):
    """Monochrome glyph icon."""

    def __init__(self, name, size=16, color=None):
        super().__init__()
        self.name = name
        self.size = size
        self.color = color
        self.set_size_request(size, size)
        self.set_valign(Gtk.Align.CENTER)
        self.connect("draw", self._draw)

    def set_icon(self, name):
        self.name = name
        self.queue_draw()

    def _draw(self, w, cr):
        col = self.color or (1, 1, 1, 1)
        p = self.get_parent()
        if not self.is_sensitive() or (p and not p.is_sensitive()):
            col = (1, 1, 1, 0.36)
        draw_glyph(cr, self.name, self.size, col)


class WinTitlebar(Gtk.HeaderBar):
    """Title bar with Windows 11 caption buttons. Pack your own widgets with pack_start()."""

    def __init__(self, window):
        super().__init__()
        self.win = window
        self.get_style_context().add_class("fx-title")
        self.set_show_close_button(False)
        self.set_custom_title(Gtk.Box())
        wc = Gtk.Box(spacing=0)
        wc.set_valign(Gtk.Align.START)
        for name, cb, extra in (("min", lambda *_: window.iconify(), ""),
                                ("max", lambda *_: self.toggle_max(), ""),
                                ("close", lambda *_: window.close(), "close")):
            b = Gtk.Button()
            b.get_style_context().add_class("wc")
            if extra:
                b.get_style_context().add_class(extra)
            ic = Icon(name, 10)
            b.add(ic)
            b.connect("clicked", cb)
            if name == "max":
                self.max_icon = ic
            wc.pack_start(b, False, False, 0)
        self.pack_end(wc)
        window.connect("window-state-event", self._wstate)
        window.set_titlebar(self)

    def toggle_max(self):
        if self.win.is_maximized():
            self.win.unmaximize()
        else:
            self.win.maximize()

    def _wstate(self, w, ev):
        maxed = bool(ev.new_window_state & Gdk.WindowState.MAXIMIZED)
        self.max_icon.set_icon("restore" if maxed else "max")
        return False


def app_title(text, logo_draw):
    """Little logo + app name for the left of the title bar."""
    box = Gtk.Box(spacing=8)
    box.set_margin_start(14)
    da = Gtk.DrawingArea()
    da.set_size_request(16, 16)
    da.set_valign(Gtk.Align.CENTER)
    da.connect("draw", lambda w, cr: logo_draw(cr, 16))
    box.pack_start(da, False, False, 0)
    lb = Gtk.Label(label=text)
    lb.get_style_context().add_class("app-title")
    box.pack_start(lb, False, False, 0)
    return box, lb


class JsonSettings(dict):
    def __init__(self, app_dir, defaults):
        super().__init__(defaults)
        self.dir = os.path.join(HOME, ".config", app_dir)
        self.file = os.path.join(self.dir, "settings.json")
        try:
            with open(self.file) as f:
                self.update(json.load(f))
        except (OSError, ValueError):
            pass

    def save(self):
        try:
            os.makedirs(self.dir, exist_ok=True)
            tmp = self.file + ".tmp"
            with open(tmp, "w") as f:
                json.dump(self, f, indent=1)
            os.replace(tmp, self.file)
        except OSError:
            pass


def vsep():
    s = Gtk.Box()
    s.get_style_context().add_class("vsep")
    return s


def logo_pixbuf(draw, size):
    surf = cairo.ImageSurface(cairo.FORMAT_ARGB32, size, size)
    draw(cairo.Context(surf), size)
    return Gdk.pixbuf_get_from_surface(surf, 0, 0, size, size)
def label(text="", cls=None, xalign=0.0, ellipsize=False, wrap=False, markup=False):
    lb = Gtk.Label()
    if markup:
        lb.set_markup(text)
    else:
        lb.set_text(text)
    lb.set_xalign(xalign)
    if cls:
        for c in cls.split():
            lb.get_style_context().add_class(c)
    if ellipsize:
        lb.set_ellipsize(Pango.EllipsizeMode.END)
    if wrap:
        lb.set_line_wrap(True)
        lb.set_line_wrap_mode(Pango.WrapMode.WORD_CHAR)
    return lb


def add_cls(w, *cls):
    ctx = w.get_style_context()
    for c in cls:
        ctx.add_class(c)
    return w


def tb_button(icon, text, cb=None, chev=False, tip=None):
    b = Gtk.Button()
    add_cls(b, "tb")
    box = Gtk.Box(spacing=7)
    if icon:
        box.pack_start(Icon(icon), False, False, 0)
    if text:
        box.pack_start(Gtk.Label(label=text), False, False, 0)
    if chev:
        box.pack_start(Icon("chevdown", 10), False, False, 0)
    b.add(box)
    b.set_relief(Gtk.ReliefStyle.NONE)
    b.set_tooltip_text(tip or text or None)
    if cb:
        b.connect("clicked", lambda w: cb(w))
    return b


def icon_button(icon, cb, cls="nb", tip=None, size=16):
    b = Gtk.Button()
    add_cls(b, cls)
    b.add(Icon(icon, size))
    b.set_relief(Gtk.ReliefStyle.NONE)
    if tip:
        b.set_tooltip_text(tip)
    b.connect("clicked", lambda w: cb(w))
    return b


def build_menu(spec):
    """spec: list of None (separator) or dicts: label, cb, icon, place, accel, sensitive, checked, sub."""
    menu = Gtk.Menu()
    add_cls(menu, "context-menu")
    for it in spec:
        if it is None:
            menu.append(Gtk.SeparatorMenuItem())
            continue
        mi = Gtk.MenuItem()
        box = Gtk.Box(spacing=12)
        ic = None
        if it.get("checked"):
            ic = Icon("check", 16)
        elif it.get("pixbuf") is not None:
            ic = Gtk.Image.new_from_pixbuf(it["pixbuf"])
        elif it.get("icon"):
            ic = Icon(it["icon"], 16)
        if ic is None:
            ic = Gtk.Box()
            ic.set_size_request(16, 16)
        box.pack_start(ic, False, False, 0)
        lb = Gtk.Label(label=it["label"])
        lb.set_xalign(0)
        box.pack_start(lb, True, True, 0)
        if it.get("accel"):
            al = Gtk.Label(label=it["accel"])
            al.get_style_context().add_class("dim-label")
            al.set_opacity(0.55)
            al.set_margin_start(24)
            box.pack_end(al, False, False, 0)
        mi.add(box)
        mi.set_sensitive(it.get("sensitive", True))
        if it.get("sub") is not None:
            mi.set_submenu(build_menu(it["sub"]))
        elif it.get("cb"):
            cb = it["cb"]
            mi.connect("activate", lambda w, cb=cb: GLib.idle_add(cb))
        menu.append(mi)
    menu.show_all()
    return menu


def popup_menu(menu, widget, event=None):
    menu.attach_to_widget(widget, None)
    if event is not None:
        menu.popup_at_pointer(event)
    else:
        menu.popup_at_widget(widget, Gdk.Gravity.SOUTH_WEST, Gdk.Gravity.NORTH_WEST, None)


# ----------------------------------------------------------------------------
# Windows 11 style dialogs
# ----------------------------------------------------------------------------
class WinDialog(Gtk.Window):
    def __init__(self, parent, title, body=None, primary="OK", secondary="Cancel", width=440, extra=None,
                 markup=False):
        super().__init__(transient_for=parent, modal=True)
        add_cls(self, "fx-dialog")
        self.set_decorated(False)
        self.set_resizable(False)
        self.set_default_size(width, -1)
        self.set_position(Gtk.WindowPosition.CENTER_ON_PARENT)
        self.set_type_hint(Gdk.WindowTypeHint.DIALOG)
        self.set_title(title)
        self.cb = None
        outer = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.body = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        add_cls(self.body, "dlg-body")
        self.body.pack_start(label(title, "dlg-title"), False, False, 0)
        if body:
            lb = label(body, wrap=True, markup=markup)
            lb.set_max_width_chars(60)
            self.body.pack_start(lb, False, False, 0)
        if extra:
            self.body.pack_start(extra, False, False, 0)
        outer.pack_start(self.body, True, True, 0)
        self.foot = Gtk.Box(spacing=8, homogeneous=True)
        add_cls(self.foot, "dlg-foot")
        self.primary_btn = None
        if primary:
            pb = Gtk.Button(label=primary)
            add_cls(pb, "primary")
            pb.connect("clicked", lambda *_: self.done(True))
            self.foot.pack_start(pb, True, True, 0)
            pb.set_can_default(True)
            self.set_default(pb)
            self.primary_btn = pb
        if secondary:
            sb = Gtk.Button(label=secondary)
            add_cls(sb, "secondary")
            sb.connect("clicked", lambda *_: self.done(False))
            self.foot.pack_start(sb, True, True, 0)
        if primary or secondary:
            outer.pack_start(self.foot, False, False, 0)
        self.add(outer)
        self.connect("key-press-event", self._key)

    def _key(self, w, ev):
        if ev.keyval == Gdk.KEY_Escape:
            self.done(False)
            return True
        return False

    def done(self, ok):
        cb = self.cb
        self.cb = None
        self.hide()
        if cb:
            cb(ok)
        self.destroy()

    def ask(self, cb=None):
        self.cb = cb
        self.show_all()
        if self.primary_btn:
            self.primary_btn.grab_focus()


def info_dialog(parent, title, body):
    WinDialog(parent, title, body, primary="OK", secondary=None).ask()


def confirm(parent, title, body, yes, cb, no="Cancel"):
    WinDialog(parent, title, body, primary=yes, secondary=no).ask(cb)


def error_list_dialog(parent, title, errors):
    lines = []
    for p, msg in errors[:12]:
        lines.append(f"{os.path.basename(p) or p}: {msg}" if p else msg)
    if len(errors) > 12:
        lines.append(f"...and {len(errors) - 12} more")
    info_dialog(parent, title, "\n".join(lines))




# ----------------------------------------------------------------------------
# Excel: formula engine
# ----------------------------------------------------------------------------
import datetime  # noqa: E402
import fnmatch  # noqa: E402
import random  # noqa: E402
import csv  # noqa: E402
import io  # noqa: E402
import statistics  # noqa: E402

sys.setrecursionlimit(20000)
EPOCH = datetime.datetime(1899, 12, 30)


def col_name(c):
    s = ""
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_index(name):
    n = 0
    for ch in name.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def cell_name(r, c):
    return f"{col_name(c)}{r + 1}"


def parse_cell_name(s):
    m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", s.strip())
    if not m or int(m.group(2)) < 1:
        return None
    return int(m.group(2)) - 1, col_index(m.group(1))


def to_dt(serial):
    return EPOCH + datetime.timedelta(days=float(serial))


def to_serial(dt):
    if isinstance(dt, datetime.datetime):
        d = dt - EPOCH
    elif isinstance(dt, datetime.date):
        d = datetime.datetime(dt.year, dt.month, dt.day) - EPOCH
    elif isinstance(dt, datetime.time):
        return (dt.hour * 3600 + dt.minute * 60 + dt.second) / 86400
    else:
        return dt
    return d.days + d.seconds / 86400


class XlErr:
    __slots__ = ("code",)

    def __init__(self, code):
        self.code = code

    def __str__(self):
        return self.code

    def __eq__(self, o):
        return isinstance(o, XlErr) and o.code == self.code

    def __hash__(self):
        return hash(self.code)


class XlRaise(Exception):
    def __init__(self, err):
        super().__init__(err.code)
        self.err = err


def err(code):
    raise XlRaise(XlErr(code))


class Rng:
    """A rectangular block of cells, used as a function argument."""

    def __init__(self, book, sheet, r1, c1, r2, c2):
        self.book, self.sheet = book, sheet
        self.r1, self.c1 = min(r1, r2), min(c1, c2)
        self.r2, self.c2 = max(r1, r2), max(c1, c2)
        mr, mc = sheet.extent()
        self.r2 = min(self.r2, mr)
        self.c2 = min(self.c2, mc)

    @property
    def rows(self):
        return max(0, self.r2 - self.r1 + 1)

    @property
    def cols(self):
        return max(0, self.c2 - self.c1 + 1)

    def get(self, i, j):
        return self.book.value(self.sheet, self.r1 + i, self.c1 + j)

    def matrix(self):
        return [[self.get(i, j) for j in range(self.cols)] for i in range(self.rows)]

    def values(self):
        """Values of the non-empty cells, row by row."""
        keys = sorted(k for k in self.sheet.cells if self.r1 <= k[0] <= self.r2 and self.c1 <= k[1] <= self.c2)
        for r, c in keys:
            v = self.book.value(self.sheet, r, c)
            if v is not None and v != "":
                yield v

    def all_values(self):
        for i in range(self.rows):
            for j in range(self.cols):
                yield self.get(i, j)


# ---- tokenizer / parser
TOK_RE = re.compile(r"""
 (?P<ws>\s+)
|(?P<str>"(?:[^"]|"")*")
|(?P<err>\#(?:DIV/0!|VALUE!|REF!|NAME\?|N/A|NUM!|NULL!|CIRC!))
|(?P<func>[A-Za-z_][A-Za-z0-9_.]*(?=\s*\())
|(?P<ref>(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(?:\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[0-9]+:\$?[0-9]+))
|(?P<num>(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)
|(?P<bool>(?i:TRUE|FALSE)(?![A-Za-z0-9_(]))
|(?P<name>[A-Za-z_][A-Za-z0-9_.]*)
|(?P<op><=|>=|<>|[-+*/^&=<>%(),;:{}])
""", re.X)


def tokenize(s):
    out = []
    pos = 0
    while pos < len(s):
        m = TOK_RE.match(s, pos)
        if not m:
            raise ValueError(f"bad character at {pos}")
        pos = m.end()
        k = m.lastgroup
        if k == "ws":
            continue
        v = m.group(k)
        if k == "ref" and re.fullmatch(r"[A-Za-z]{1,3}\d+", v) and pos < len(s) and (s[pos].isalnum() or s[pos] == "_"):
            k = "name"
        out.append((k, v))
    return out


def split_sheet(ref):
    if "!" in ref:
        sh, ref = ref.rsplit("!", 1)
        if sh.startswith("'"):
            sh = sh[1:-1].replace("''", "'")
        return sh, ref
    return None, ref


def parse_ref_part(p):
    """'$A$1' -> (row, col); 'A' -> (None, col); '5' -> (row, None)."""
    m = re.fullmatch(r"\$?([A-Za-z]{1,3})?\$?([0-9]+)?", p)
    col = col_index(m.group(1)) if m.group(1) else None
    row = int(m.group(2)) - 1 if m.group(2) else None
    return row, col


class Parser:
    def __init__(self, toks):
        self.t = toks
        self.i = 0

    def peek(self, k=None, v=None):
        if self.i >= len(self.t):
            return None
        tk = self.t[self.i]
        if k and tk[0] != k:
            return None
        if v and tk[1] != v:
            return None
        return tk

    def take(self, k=None, v=None):
        tk = self.peek(k, v)
        if tk is None:
            raise ValueError(f"expected {v or k}")
        self.i += 1
        return tk

    def parse(self):
        n = self.cmp()
        if self.i != len(self.t):
            raise ValueError("unexpected " + self.t[self.i][1])
        return n

    def cmp(self):
        n = self.concat()
        while self.peek("op") and self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            n = ("bin", op, n, self.concat())
        return n

    def concat(self):
        n = self.add()
        while self.peek("op", "&"):
            self.take()
            n = ("bin", "&", n, self.add())
        return n

    def add(self):
        n = self.mul()
        while self.peek("op") and self.peek()[1] in "+-":
            op = self.take()[1]
            n = ("bin", op, n, self.mul())
        return n

    def mul(self):
        n = self.pow()
        while self.peek("op") and self.peek()[1] in "*/":
            op = self.take()[1]
            n = ("bin", op, n, self.pow())
        return n

    def pow(self):
        n = self.unary()
        while self.peek("op", "^"):
            self.take()
            n = ("bin", "^", n, self.unary())
        return n

    def unary(self):
        if self.peek("op", "-"):
            self.take()
            return ("neg", self.unary())
        if self.peek("op", "+"):
            self.take()
            return self.unary()
        return self.postfix()

    def postfix(self):
        n = self.atom()
        while self.peek("op", "%"):
            self.take()
            n = ("pct", n)
        return n

    def atom(self):
        tk = self.take()
        k, v = tk
        if k == "num":
            return ("num", float(v))
        if k == "str":
            return ("str", v[1:-1].replace('""', '"'))
        if k == "bool":
            return ("bool", v.upper() == "TRUE")
        if k == "err":
            return ("err", v)
        if k == "ref":
            sh, rest = split_sheet(v)
            if ":" in rest:
                a, b = rest.split(":")
                r1, c1 = parse_ref_part(a)
                r2, c2 = parse_ref_part(b)
                return ("range", sh, 0 if r1 is None else r1, 0 if c1 is None else c1,
                        10 ** 7 if r2 is None else r2, 10 ** 5 if c2 is None else c2)
            r, c = parse_ref_part(rest)
            return ("ref", sh, r, c)
        if k == "func":
            name = v.upper()
            if name.startswith("_XLFN."):
                name = name[6:]
            self.take("op", "(")
            args = []
            if not self.peek("op", ")"):
                while True:
                    if self.peek("op") and self.peek()[1] in (",", ";", ")"):
                        args.append(("blank",))
                    else:
                        args.append(self.cmp())
                    if self.peek("op") and self.peek()[1] in (",", ";"):
                        self.take()
                        continue
                    break
            self.take("op", ")")
            return ("fn", name, args)
        if k == "op" and v == "(":
            n = self.cmp()
            self.take("op", ")")
            return n
        if k == "op" and v == "{":
            rows = [[]]
            while not self.peek("op", "}"):
                if self.peek("op", ","):
                    self.take()
                    continue
                if self.peek("op", ";"):
                    self.take()
                    rows.append([])
                    continue
                rows[-1].append(self.unary())
            self.take("op", "}")
            return ("array", rows)
        if k == "name":
            return ("name", v)
        raise ValueError("unexpected " + v)


def parse_formula(text):
    return Parser(tokenize(text)).parse()


# ---- coercion
def to_num(v):
    if isinstance(v, XlErr):
        raise XlRaise(v)
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if v is None or v == "":
        return 0.0
    if isinstance(v, str):
        p = parse_number(v)
        if p is not None:
            return p[0]
        err("#VALUE!")
    if isinstance(v, Rng):
        return to_num(scalar(v))
    err("#VALUE!")


def to_str(v):
    if isinstance(v, XlErr):
        raise XlRaise(v)
    if isinstance(v, Rng):
        return to_str(scalar(v))
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return general(v)
    return str(v)


def to_bool(v):
    if isinstance(v, XlErr):
        raise XlRaise(v)
    if isinstance(v, Rng):
        return to_bool(scalar(v))
    if isinstance(v, str):
        if v.upper() in ("TRUE", "FALSE"):
            return v.upper() == "TRUE"
        err("#VALUE!")
    return bool(to_num(v))


def scalar(v):
    if isinstance(v, Rng):
        if v.rows == 0 or v.cols == 0:
            return None
        return v.get(0, 0)
    return v


def general(v):
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        if v != v:
            return "#NUM!"
        if float(v).is_integer() and abs(v) < 1e15:
            return str(int(v))
        a = abs(v)
        if a >= 1e11 or a < 1e-9:
            return f"{v:.5E}".replace("E+0", "E+").replace("E-0", "E-")
        s = f"{v:.10g}"
        if "e" in s:
            s = f"{v:.10f}".rstrip("0").rstrip(".")
        return s
    return str(v)


NUM_RE = re.compile(r"^\s*([-+])?\s*(\$)?\s*([0-9]{1,3}(?:,[0-9]{3})+|[0-9]*\.?[0-9]+(?:[eE][-+]?[0-9]+)?|[0-9]+\.)(?:(?<=[0-9])\.([0-9]+))?\s*(%)?\s*$")


def parse_number(s):
    """Text typed into a cell -> (number, suggested format) or None."""
    m = NUM_RE.match(s)
    if not m:
        return None
    sign, dollar, body, frac, pct = m.groups()
    try:
        v = float(body.replace(",", "") + ("." + frac if frac else ""))
    except ValueError:
        return None
    if sign == "-":
        v = -v
    fmt = None
    if pct:
        v /= 100
        fmt = "0%" if v * 100 == int(v * 100) else "0.00%"
    elif dollar:
        fmt = '"$"#,##0.00'
    elif "," in body:
        fmt = "#,##0" if not frac else "#,##0.00"
    return v, fmt


DATE_RES = [(re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$"), "mdy"), (re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})$"), "ymd"),
            (re.compile(r"^(\d{1,2})/(\d{1,2})$"), "md")]


def parse_date(s):
    s = s.strip()
    for rx, kind in DATE_RES:
        m = rx.match(s)
        if not m:
            continue
        g = [int(x) for x in m.groups()]
        try:
            if kind == "mdy":
                y = g[2] + (2000 if g[2] < 30 else 1900 if g[2] < 100 else 0)
                d = datetime.datetime(y, g[0], g[1])
                return to_serial(d), "m/d/yyyy"
            if kind == "ymd":
                return to_serial(datetime.datetime(g[0], g[1], g[2])), "yyyy-mm-dd"
            if kind == "md":
                return to_serial(datetime.datetime(datetime.date.today().year, g[0], g[1])), "d-mmm"
        except ValueError:
            return None
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?$", s)
    if m:
        h, mi, sec = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        if m.group(4):
            h = h % 12 + (12 if m.group(4).lower() == "pm" else 0)
        if h < 24 and mi < 60 and sec < 60:
            return (h * 3600 + mi * 60 + sec) / 86400, "h:mm AM/PM" if m.group(4) else "h:mm"
    return None


def parse_input(raw):
    """Cell text -> ('f', ast) | ('v', value)."""
    if raw is None or raw == "":
        return ("v", None)
    if raw.startswith("=") and len(raw) > 1:
        try:
            return ("f", parse_formula(raw[1:]))
        except (ValueError, IndexError):
            return ("v", XlErr("#NAME?"))
    if raw.startswith("'"):
        return ("v", raw[1:])
    if raw.upper() in ("TRUE", "FALSE"):
        return ("v", raw.upper() == "TRUE")
    if re.fullmatch(r"#(?:DIV/0!|VALUE!|REF!|NAME\?|N/A|NUM!|NULL!)", raw):
        return ("v", XlErr(raw))
    p = parse_number(raw)
    if p is not None:
        return ("v", p[0])
    d = parse_date(raw)
    if d is not None:
        return ("v", d[0])
    return ("v", raw)


# ---- comparison
def type_rank(v):
    if isinstance(v, XlErr):
        return 4
    if isinstance(v, bool):
        return 3
    if isinstance(v, str):
        return 2
    return 1


def compare(a, b):
    if a is None:
        a = "" if isinstance(b, str) else 0.0 if not isinstance(b, bool) else False
    if b is None:
        b = "" if isinstance(a, str) else 0.0 if not isinstance(a, bool) else False
    ra, rb = type_rank(a), type_rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if isinstance(a, str):
        a, b = a.lower(), b.lower()
    return (a > b) - (a < b)


def criteria(c):
    """SUMIF / COUNTIF criteria -> predicate."""
    if isinstance(c, (int, float)) and not isinstance(c, bool):
        return lambda v: isinstance(v, (int, float)) and not isinstance(v, bool) and float(v) == float(c)
    if isinstance(c, bool):
        return lambda v: v is c
    s = to_str(c)
    m = re.match(r"^(<=|>=|<>|<|>|=)?(.*)$", s, re.S)
    op, rest = m.group(1) or "=", m.group(2)
    num = parse_number(rest)
    if num is not None:
        n = num[0]

        def pred(v):
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                return op == "<>"
            return {"=": v == n, "<>": v != n, "<": v < n, ">": v > n, "<=": v <= n, ">=": v >= n}[op]
        return pred
    low = rest.lower()
    wild = any(ch in rest for ch in "*?")

    def pred_s(v):
        if v is None:
            v = ""
        if not isinstance(v, str):
            sv = to_str(v).lower() if not isinstance(v, XlErr) else v.code.lower()
            if op in ("=", "<>"):
                eq = sv == low
                return eq if op == "=" else not eq
            return False
        sv = v.lower()
        if op in ("=", "<>"):
            eq = fnmatch.fnmatchcase(sv, low) if wild else sv == low
            return eq if op == "=" else not eq
        c2 = (sv > low) - (sv < low)
        return {"<": c2 < 0, ">": c2 > 0, "<=": c2 <= 0, ">=": c2 >= 0}[op]
    return pred_s


# ---- functions
FUNCS = {}
LAZY = set()


def fn(*names, lazy=False):
    def deco(f):
        for n in names:
            FUNCS[n] = f
            if lazy:
                LAZY.add(n)
        return f
    return deco


def flat(args, include_text=False):
    """Numbers from the arguments: range cells that hold numbers, direct arguments coerced."""
    for a in args:
        if isinstance(a, Rng):
            for v in a.values():
                if isinstance(v, XlErr):
                    raise XlRaise(v)
                if isinstance(v, bool):
                    continue
                if isinstance(v, (int, float)):
                    yield float(v)
                elif include_text:
                    yield 0.0
        elif isinstance(a, list):
            for row in a:
                yield from flat(row, include_text)
        elif a is None:
            continue
        else:
            yield to_num(a)


def nargs(args, lo, hi=None):
    if len(args) < lo or (hi is not None and len(args) > hi):
        err("#VALUE!")


def opt(args, i, default):
    return args[i] if len(args) > i and args[i] is not None else default


@fn("SUM")
def f_sum(b, a):
    return sum(flat(a))


@fn("PRODUCT")
def f_product(b, a):
    p = 1.0
    for v in flat(a):
        p *= v
    return p


@fn("AVERAGE")
def f_avg(b, a):
    v = list(flat(a))
    if not v:
        err("#DIV/0!")
    return sum(v) / len(v)


@fn("MEDIAN")
def f_median(b, a):
    v = list(flat(a))
    if not v:
        err("#NUM!")
    return float(statistics.median(v))


@fn("MIN")
def f_min(b, a):
    v = list(flat(a))
    return min(v) if v else 0.0


@fn("MAX")
def f_max(b, a):
    v = list(flat(a))
    return max(v) if v else 0.0


@fn("STDEV", "STDEV.S")
def f_stdev(b, a):
    v = list(flat(a))
    if len(v) < 2:
        err("#DIV/0!")
    return statistics.stdev(v)


@fn("COUNT")
def f_count(b, a):
    n = 0
    for x in a:
        if isinstance(x, Rng):
            n += sum(1 for v in x.values() if isinstance(v, (int, float)) and not isinstance(v, bool))
        elif isinstance(x, (int, float)) and not isinstance(x, bool):
            n += 1
    return float(n)


@fn("COUNTA")
def f_counta(b, a):
    n = 0
    for x in a:
        if isinstance(x, Rng):
            n += sum(1 for _ in x.values())
        elif x is not None:
            n += 1
    return float(n)


@fn("COUNTBLANK")
def f_countblank(b, a):
    nargs(a, 1, 1)
    r = a[0]
    return float(sum(1 for v in r.all_values() if v is None or v == "")) if isinstance(r, Rng) else 0.0


def _if_pairs(a, sum_idx):
    rng = a[0]
    pred = criteria(scalar(a[1]))
    tgt = a[sum_idx] if len(a) > sum_idx and isinstance(a[sum_idx], Rng) else rng
    if not isinstance(rng, Rng):
        err("#VALUE!")
    out = []
    for i in range(rng.rows):
        for j in range(rng.cols):
            if pred(rng.get(i, j)):
                out.append(tgt.get(i, j) if i < tgt.rows and j < tgt.cols else None)
    return out


@fn("SUMIF")
def f_sumif(b, a):
    nargs(a, 2, 3)
    return sum(v for v in _if_pairs(a, 2) if isinstance(v, (int, float)) and not isinstance(v, bool))


@fn("AVERAGEIF")
def f_averageif(b, a):
    nargs(a, 2, 3)
    v = [v for v in _if_pairs(a, 2) if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if not v:
        err("#DIV/0!")
    return sum(v) / len(v)


@fn("COUNTIF")
def f_countif(b, a):
    nargs(a, 2, 2)
    return float(len(_if_pairs(a, 9)))


def _ifs(a, start):
    """SUMIFS-style pairs starting at index start -> list of (i, j) that match all."""
    pairs = list(zip(a[start::2], a[start + 1::2]))
    if not pairs or not all(isinstance(r, Rng) for r, _ in pairs):
        err("#VALUE!")
    r0 = pairs[0][0]
    hits = []
    preds = [(r, criteria(scalar(c))) for r, c in pairs]
    for i in range(r0.rows):
        for j in range(r0.cols):
            if all(p(r.get(i, j)) for r, p in preds):
                hits.append((i, j))
    return hits


@fn("SUMIFS")
def f_sumifs(b, a):
    tgt = a[0]
    tot = 0.0
    for i, j in _ifs(a, 1):
        v = tgt.get(i, j)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            tot += v
    return tot


@fn("COUNTIFS")
def f_countifs(b, a):
    return float(len(_ifs(a, 0)))


@fn("AVERAGEIFS")
def f_averageifs(b, a):
    tgt = a[0]
    v = [tgt.get(i, j) for i, j in _ifs(a, 1)]
    v = [x for x in v if isinstance(x, (int, float)) and not isinstance(x, bool)]
    if not v:
        err("#DIV/0!")
    return sum(v) / len(v)


@fn("SUMPRODUCT")
def f_sumproduct(b, a):
    mats = [x.matrix() if isinstance(x, Rng) else [[x]] for x in a]
    tot = 0.0
    for i in range(len(mats[0])):
        for j in range(len(mats[0][0])):
            p = 1.0
            for m in mats:
                try:
                    v = m[i][j]
                except IndexError:
                    err("#VALUE!")
                p *= v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
            tot += p
    return tot


@fn("IF", lazy=True)
def f_if(ev, nodes):
    if len(nodes) < 2 or len(nodes) > 3:
        err("#VALUE!")
    c = to_bool(scalar(ev(nodes[0])))
    if c:
        return ev(nodes[1])
    return ev(nodes[2]) if len(nodes) > 2 else False


@fn("IFS", lazy=True)
def f_ifs(ev, nodes):
    for c, v in zip(nodes[0::2], nodes[1::2]):
        if to_bool(scalar(ev(c))):
            return ev(v)
    err("#N/A")


@fn("IFERROR", lazy=True)
def f_iferror(ev, nodes):
    try:
        v = scalar(ev(nodes[0]))
        if isinstance(v, XlErr):
            return ev(nodes[1])
        return v
    except XlRaise:
        return ev(nodes[1])


@fn("IFNA", lazy=True)
def f_ifna(ev, nodes):
    try:
        v = scalar(ev(nodes[0]))
        if isinstance(v, XlErr) and v.code == "#N/A":
            return ev(nodes[1])
        return v
    except XlRaise as e:
        if e.err.code == "#N/A":
            return ev(nodes[1])
        raise


@fn("ISERROR", lazy=True)
def f_iserror(ev, nodes):
    try:
        return isinstance(scalar(ev(nodes[0])), XlErr)
    except XlRaise:
        return True


@fn("ISNA", lazy=True)
def f_isna(ev, nodes):
    try:
        v = scalar(ev(nodes[0]))
        return isinstance(v, XlErr) and v.code == "#N/A"
    except XlRaise as e:
        return e.err.code == "#N/A"


@fn("ISBLANK")
def f_isblank(b, a):
    return scalar(a[0]) is None


@fn("ISNUMBER")
def f_isnumber(b, a):
    v = scalar(a[0])
    return isinstance(v, (int, float)) and not isinstance(v, bool)


@fn("ISTEXT")
def f_istext(b, a):
    return isinstance(scalar(a[0]), str)


@fn("AND")
def f_and(b, a):
    vals = []
    for x in a:
        if isinstance(x, Rng):
            vals += [bool(v) for v in x.values() if isinstance(v, (bool, int, float))]
        else:
            vals.append(to_bool(x))
    if not vals:
        err("#VALUE!")
    return all(vals)


@fn("OR")
def f_or(b, a):
    vals = []
    for x in a:
        if isinstance(x, Rng):
            vals += [bool(v) for v in x.values() if isinstance(v, (bool, int, float))]
        else:
            vals.append(to_bool(x))
    if not vals:
        err("#VALUE!")
    return any(vals)


@fn("XOR")
def f_xor(b, a):
    return sum(1 for x in a if to_bool(x)) % 2 == 1


@fn("NOT")
def f_not(b, a):
    nargs(a, 1, 1)
    return not to_bool(a[0])


@fn("TRUE")
def f_true(b, a):
    return True


@fn("FALSE")
def f_false(b, a):
    return False


def _round(x, d, mode):
    m = 10 ** d
    v = x * m
    if mode == 0:
        r = math.floor(abs(v) + 0.5 + 1e-9)
    elif mode == 1:
        r = math.ceil(abs(v) - 1e-9)
    else:
        r = math.floor(abs(v) + 1e-9)
    return math.copysign(r, x) / m


@fn("ROUND")
def f_round(b, a):
    nargs(a, 1, 2)
    return _round(to_num(a[0]), int(to_num(opt(a, 1, 0))), 0)


@fn("ROUNDUP")
def f_roundup(b, a):
    return _round(to_num(a[0]), int(to_num(opt(a, 1, 0))), 1)


@fn("ROUNDDOWN", "TRUNC")
def f_rounddown(b, a):
    return _round(to_num(a[0]), int(to_num(opt(a, 1, 0))), 2)


@fn("INT")
def f_int(b, a):
    return float(math.floor(to_num(a[0])))


@fn("ABS")
def f_abs(b, a):
    return abs(to_num(a[0]))


@fn("SQRT")
def f_sqrt(b, a):
    v = to_num(a[0])
    if v < 0:
        err("#NUM!")
    return math.sqrt(v)


@fn("POWER")
def f_power(b, a):
    return _pow(to_num(a[0]), to_num(a[1]))


def _pow(x, y):
    try:
        r = x ** y
    except ZeroDivisionError:
        err("#DIV/0!")
    except OverflowError:
        err("#NUM!")
    if isinstance(r, complex):
        err("#NUM!")
    return float(r)


@fn("MOD")
def f_mod(b, a):
    d = to_num(a[1])
    if d == 0:
        err("#DIV/0!")
    return to_num(a[0]) - d * math.floor(to_num(a[0]) / d)


@fn("EXP")
def f_exp(b, a):
    return math.exp(to_num(a[0]))


@fn("LN")
def f_ln(b, a):
    v = to_num(a[0])
    if v <= 0:
        err("#NUM!")
    return math.log(v)


@fn("LOG")
def f_log(b, a):
    v = to_num(a[0])
    base = to_num(opt(a, 1, 10.0))
    if v <= 0 or base <= 0 or base == 1:
        err("#NUM!")
    return math.log(v, base)


@fn("LOG10")
def f_log10(b, a):
    v = to_num(a[0])
    if v <= 0:
        err("#NUM!")
    return math.log10(v)


@fn("PI")
def f_pi(b, a):
    return math.pi


def _mk1(name, f):
    FUNCS[name] = lambda b, a: float(f(to_num(a[0])))


for _n, _f in (("SIN", math.sin), ("COS", math.cos), ("TAN", math.tan), ("ASIN", math.asin), ("ACOS", math.acos),
               ("ATAN", math.atan), ("RADIANS", math.radians), ("DEGREES", math.degrees),
               ("SIGN", lambda x: (x > 0) - (x < 0)), ("FACT", lambda x: math.factorial(int(x)))):
    _mk1(_n, _f)


@fn("RAND")
def f_rand(b, a):
    return random.random()


@fn("RANDBETWEEN")
def f_randbetween(b, a):
    return float(random.randint(math.ceil(to_num(a[0])), math.floor(to_num(a[1]))))


@fn("CEILING", "CEILING.MATH")
def f_ceiling(b, a):
    x, s = to_num(a[0]), to_num(opt(a, 1, 1.0))
    return 0.0 if s == 0 else math.ceil(x / s - 1e-12) * s


@fn("FLOOR", "FLOOR.MATH")
def f_floor(b, a):
    x, s = to_num(a[0]), to_num(opt(a, 1, 1.0))
    if s == 0:
        err("#DIV/0!")
    return math.floor(x / s + 1e-12) * s


@fn("LEN")
def f_len(b, a):
    return float(len(to_str(a[0])))


@fn("LEFT")
def f_left(b, a):
    n = int(to_num(opt(a, 1, 1.0)))
    if n < 0:
        err("#VALUE!")
    return to_str(a[0])[:n]


@fn("RIGHT")
def f_right(b, a):
    n = int(to_num(opt(a, 1, 1.0)))
    if n < 0:
        err("#VALUE!")
    s = to_str(a[0])
    return s[len(s) - n:] if n else ""


@fn("MID")
def f_mid(b, a):
    s, st, n = to_str(a[0]), int(to_num(a[1])), int(to_num(a[2]))
    if st < 1 or n < 0:
        err("#VALUE!")
    return s[st - 1:st - 1 + n]


@fn("UPPER")
def f_upper(b, a):
    return to_str(a[0]).upper()


@fn("LOWER")
def f_lower(b, a):
    return to_str(a[0]).lower()


@fn("PROPER")
def f_proper(b, a):
    return to_str(a[0]).title()


@fn("TRIM")
def f_trim(b, a):
    return " ".join(to_str(a[0]).split())


@fn("CONCATENATE", "CONCAT")
def f_concat(b, a):
    out = []
    for x in a:
        if isinstance(x, Rng):
            out += [to_str(v) for v in x.all_values()]
        else:
            out.append(to_str(x))
    return "".join(out)


@fn("TEXTJOIN")
def f_textjoin(b, a):
    sep, skip = to_str(a[0]), to_bool(a[1])
    out = []
    for x in a[2:]:
        vals = list(x.all_values()) if isinstance(x, Rng) else [x]
        for v in vals:
            s = to_str(v)
            if s or not skip:
                out.append(s)
    return sep.join(out)


@fn("REPT")
def f_rept(b, a):
    return to_str(a[0]) * max(0, int(to_num(a[1])))


@fn("SUBSTITUTE")
def f_substitute(b, a):
    s, old, new = to_str(a[0]), to_str(a[1]), to_str(a[2])
    if len(a) > 3:
        n = int(to_num(a[3]))
        idx = -1
        for _ in range(n):
            idx = s.find(old, idx + 1)
            if idx < 0:
                return s
        return s[:idx] + new + s[idx + len(old):]
    return s.replace(old, new) if old else s


@fn("FIND")
def f_find(b, a):
    sub, s = to_str(a[0]), to_str(a[1])
    i = s.find(sub, int(to_num(opt(a, 2, 1.0))) - 1)
    if i < 0:
        err("#VALUE!")
    return float(i + 1)


@fn("SEARCH")
def f_search(b, a):
    sub, s = to_str(a[0]), to_str(a[1])
    st = int(to_num(opt(a, 2, 1.0))) - 1
    if any(ch in sub for ch in "*?"):
        for i in range(st, len(s) + 1):
            for j in range(len(s), i - 1, -1):
                if fnmatch.fnmatchcase(s[i:j].lower(), sub.lower()):
                    return float(i + 1)
        err("#VALUE!")
    i = s.lower().find(sub.lower(), st)
    if i < 0:
        err("#VALUE!")
    return float(i + 1)


@fn("EXACT")
def f_exact(b, a):
    return to_str(a[0]) == to_str(a[1])


@fn("VALUE")
def f_value(b, a):
    v = scalar(a[0])
    if isinstance(v, (int, float)):
        return float(v)
    p = parse_number(to_str(v)) or parse_date(to_str(v))
    if p is None:
        err("#VALUE!")
    return p[0]


@fn("TEXT")
def f_text(b, a):
    v = scalar(a[0])
    if isinstance(v, str):
        p = parse_number(v)
        if p is None:
            return v
        v = p[0]
    return format_value(to_num(v), to_str(a[1]))


@fn("CHAR")
def f_char(b, a):
    return chr(int(to_num(a[0])))


@fn("CODE")
def f_code(b, a):
    s = to_str(a[0])
    if not s:
        err("#VALUE!")
    return float(ord(s[0]))


@fn("TODAY")
def f_today(b, a):
    return float(to_serial(datetime.date.today()))


@fn("NOW")
def f_now(b, a):
    return to_serial(datetime.datetime.now())


@fn("DATE")
def f_date(b, a):
    y, m, d = int(to_num(a[0])), int(to_num(a[1])), int(to_num(a[2]))
    if y < 1900:
        y += 1900
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    try:
        return float(to_serial(datetime.date(y, m, 1)) + d - 1)
    except ValueError:
        err("#NUM!")


@fn("TIME")
def f_time(b, a):
    return (to_num(a[0]) * 3600 + to_num(a[1]) * 60 + to_num(a[2])) / 86400 % 1


def _dt(a):
    v = scalar(a[0])
    if isinstance(v, str):
        p = parse_date(v)
        if p is None:
            err("#VALUE!")
        v = p[0]
    return to_dt(to_num(v))


@fn("YEAR")
def f_year(b, a):
    return float(_dt(a).year)


@fn("MONTH")
def f_month(b, a):
    return float(_dt(a).month)


@fn("DAY")
def f_day(b, a):
    return float(_dt(a).day)


@fn("HOUR")
def f_hour(b, a):
    return float(_dt(a).hour)


@fn("MINUTE")
def f_minute(b, a):
    return float(_dt(a).minute)


@fn("WEEKDAY")
def f_weekday(b, a):
    d = _dt(a)
    t = int(to_num(opt(a, 1, 1.0)))
    wd = d.isoweekday() % 7 + 1  # Sunday = 1
    if t == 2:
        return float(d.isoweekday())
    if t == 3:
        return float(d.weekday())
    return float(wd)


@fn("DAYS")
def f_days(b, a):
    return float(math.floor(to_num(a[0])) - math.floor(to_num(a[1])))


@fn("EDATE")
def f_edate(b, a):
    d = to_dt(to_num(a[0]))
    m = d.month - 1 + int(to_num(a[1]))
    y, m = d.year + m // 12, m % 12 + 1
    import calendar
    day = min(d.day, calendar.monthrange(y, m)[1])
    return float(to_serial(datetime.date(y, m, day)))


@fn("CHOOSE")
def f_choose(b, a):
    i = int(to_num(a[0]))
    if i < 1 or i >= len(a):
        err("#VALUE!")
    return a[i]


def _lookup_pos(val, seq, mode):
    """Position of val in seq. mode 0 exact, 1 largest <= val, -1 smallest >= val."""
    if mode == 0:
        pred = criteria(val) if isinstance(val, str) and any(c in val for c in "*?") else None
        for i, v in enumerate(seq):
            if pred(v) if pred else (v is not None and type_rank(v) == type_rank(val) and compare(v, val) == 0):
                return i
        err("#N/A")
    best = None
    for i, v in enumerate(seq):
        if v is None or type_rank(v) != type_rank(val):
            continue
        c = compare(v, val)
        if mode == 1:
            if c <= 0:
                best = i
            else:
                break
        else:
            if c >= 0:
                best = i
            else:
                break
    if best is None:
        err("#N/A")
    return best


@fn("VLOOKUP")
def f_vlookup(b, a):
    nargs(a, 3, 4)
    val, tbl, ci = scalar(a[0]), a[1], int(to_num(a[2]))
    approx = to_bool(opt(a, 3, True))
    if not isinstance(tbl, Rng):
        err("#VALUE!")
    if ci < 1 or ci > tbl.c2 - tbl.c1 + 1:
        err("#REF!")
    i = _lookup_pos(val, [tbl.get(r, 0) for r in range(tbl.rows)], 1 if approx else 0)
    return tbl.get(i, ci - 1)


@fn("HLOOKUP")
def f_hlookup(b, a):
    val, tbl, ri = scalar(a[0]), a[1], int(to_num(a[2]))
    approx = to_bool(opt(a, 3, True))
    if not isinstance(tbl, Rng):
        err("#VALUE!")
    j = _lookup_pos(val, [tbl.get(0, c) for c in range(tbl.cols)], 1 if approx else 0)
    return tbl.get(ri - 1, j)


@fn("XLOOKUP")
def f_xlookup(b, a):
    val, look, ret = scalar(a[0]), a[1], a[2]
    seq = list(look.all_values())
    try:
        i = _lookup_pos(val, seq, 0)
    except XlRaise:
        if len(a) > 3 and a[3] is not None:
            return a[3]
        raise
    if look.cols == 1:
        return ret.get(i, 0)
    return ret.get(0, i)


@fn("MATCH")
def f_match(b, a):
    val, rng = scalar(a[0]), a[1]
    mode = int(to_num(opt(a, 2, 1.0)))
    seq = list(rng.all_values()) if isinstance(rng, Rng) else [rng]
    return float(_lookup_pos(val, seq, mode) + 1)


@fn("INDEX")
def f_index(b, a):
    rng = a[0]
    r = int(to_num(opt(a, 1, 0.0)))
    c = int(to_num(opt(a, 2, 0.0)))
    if not isinstance(rng, Rng):
        return rng
    if rng.rows == 1 and len(a) == 2:
        r, c = 1, r
    if c == 0:
        c = 1
    if r < 1 or c < 1 or r > rng.r2 - rng.r1 + 1 or c > rng.c2 - rng.c1 + 1:
        err("#REF!")
    return rng.get(r - 1, c - 1)


@fn("ROWS")
def f_rows(b, a):
    return float(a[0].rows) if isinstance(a[0], Rng) else 1.0


@fn("COLUMNS")
def f_columns(b, a):
    return float(a[0].cols) if isinstance(a[0], Rng) else 1.0


@fn("NA")
def f_na(b, a):
    err("#N/A")


@fn("PMT")
def f_pmt(b, a):
    rate, n, pv = to_num(a[0]), to_num(a[1]), to_num(a[2])
    fv = to_num(opt(a, 3, 0.0))
    typ = to_num(opt(a, 4, 0.0))
    if n == 0:
        err("#NUM!")
    if rate == 0:
        return -(pv + fv) / n
    f = (1 + rate) ** n
    return -(rate * (pv * f + fv)) / ((1 + rate * typ) * (f - 1))


@fn("FV")
def f_fv(b, a):
    rate, n, pmt = to_num(a[0]), to_num(a[1]), to_num(a[2])
    pv = to_num(opt(a, 3, 0.0))
    typ = to_num(opt(a, 4, 0.0))
    if rate == 0:
        return -(pv + pmt * n)
    f = (1 + rate) ** n
    return -(pv * f + pmt * (1 + rate * typ) * (f - 1) / rate)


FUNC_GROUPS = [
    ("Math", ["SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "PRODUCT", "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "ABS",
              "SQRT", "POWER", "MOD", "MEDIAN", "STDEV", "SUMPRODUCT", "RAND", "RANDBETWEEN", "PI"]),
    ("Conditional", ["IF", "IFS", "IFERROR", "AND", "OR", "NOT", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF"]),
    ("Lookup", ["VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "CHOOSE"]),
    ("Text", ["CONCAT", "TEXTJOIN", "LEFT", "RIGHT", "MID", "LEN", "UPPER", "LOWER", "PROPER", "TRIM", "SUBSTITUTE",
              "FIND", "TEXT", "VALUE"]),
    ("Date & time", ["TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY", "WEEKDAY", "DAYS", "EDATE", "TIME", "HOUR", "MINUTE"]),
    ("Financial", ["PMT", "FV"]),
]


# ---- number formats
def _date_fmt(v, fmt):
    try:
        d = to_dt(v)
    except (OverflowError, ValueError):
        return "#" * 6
    ampm = "AM/PM" in fmt.upper()
    out = []
    i = 0
    f = fmt
    last_h = False
    while i < len(f):
        rest = f[i:]
        low = rest.lower()
        m = re.match(r"(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am/pm|a/p)", low)
        if m:
            t = m.group(1)
            if t in ("yyyy",):
                out.append(f"{d.year:04d}")
            elif t == "yy":
                out.append(f"{d.year % 100:02d}")
            elif t in ("mm", "m") and (last_h or re.match(r"[^a-z]*s", low[len(t):])):
                out.append(f"{d.minute:02d}" if t == "mm" else str(d.minute))
            elif t == "mmmmm":
                out.append(d.strftime("%B")[0])
            elif t == "mmmm":
                out.append(d.strftime("%B"))
            elif t == "mmm":
                out.append(d.strftime("%b"))
            elif t == "mm":
                out.append(f"{d.month:02d}")
            elif t == "m":
                out.append(str(d.month))
            elif t == "dddd":
                out.append(d.strftime("%A"))
            elif t == "ddd":
                out.append(d.strftime("%a"))
            elif t == "dd":
                out.append(f"{d.day:02d}")
            elif t == "d":
                out.append(str(d.day))
            elif t in ("hh", "h"):
                h = d.hour % 12 or 12 if ampm else d.hour
                out.append(f"{h:02d}" if t == "hh" else str(h))
            elif t in ("ss", "s"):
                out.append(f"{d.second:02d}")
            elif t == "am/pm":
                out.append("AM" if d.hour < 12 else "PM")
            elif t == "a/p":
                out.append("A" if d.hour < 12 else "P")
            last_h = t in ("hh", "h")
            i += len(t)
            continue
        ch = f[i]
        if ch == '"':
            j = f.find('"', i + 1)
            j = len(f) if j < 0 else j
            out.append(f[i + 1:j])
            i = j + 1
            continue
        if ch == "\\":
            out.append(f[i + 1:i + 2])
            i += 2
            continue
        if ch not in "[]":
            out.append(ch)
        i += 1
    return "".join(out)


def is_date_fmt(fmt):
    f = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", fmt).lower()
    return bool(re.search(r"[ydhs]", f)) or ("m" in f and not re.search(r"[0#?]", f))


def format_value(v, fmt):
    """Value + Excel number format -> display text."""
    if v is None:
        return ""
    if isinstance(v, XlErr):
        return v.code
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, str):
        return v
    if not fmt or fmt == "General":
        return general(v)
    if fmt == "@":
        return general(v)
    secs = fmt.split(";")
    if len(secs) > 1 and v < 0:
        f = secs[1]
        v = -v
    elif len(secs) > 2 and v == 0:
        f = secs[2]
    else:
        f = secs[0]
    f = re.sub(r"\[(?:Red|Black|Blue|Green|Color\d+|\$[^\]]*)\]", "", f, flags=re.I)
    f = f.replace("_)", " ").replace("_(", "").replace("* ", "")
    f = re.sub(r"_.", "", f)
    if is_date_fmt(f):
        return _date_fmt(v, f)
    if f.upper() == "GENERAL":
        return general(v)
    pct = f.count("%")
    if pct:
        v *= 100 ** pct
    sci = re.search(r"E[+-]", f, re.I)
    numpart = re.search(r"[0#?,.]*[0#?][0#?,.]*", f)
    if not numpart:
        return f.replace('"', "")
    np_ = numpart.group(0)
    prefix = f[:numpart.start()]
    suffix = f[numpart.end():]
    dec = 0
    if "." in np_:
        dec = sum(1 for ch in np_.split(".", 1)[1] if ch in "0#?")
    comma = "," in np_.split(".")[0]
    neg = v < 0 and len(secs) == 1
    a = abs(v)
    if sci:
        s = f"{a:.{dec}E}"
        mant, ex = s.split("E")
        exp_digits = len(re.search(r"E[+-]([0#]+)", f, re.I).group(1)) if re.search(r"E[+-]([0#]+)", f, re.I) else 2
        s = f"{mant}E{ex[0]}{int(ex[1:]):0{exp_digits}d}"
        suffix = re.sub(r"E[+-][0#]+", "", suffix, flags=re.I)
    else:
        s = f"{a:,.{dec}f}" if comma else f"{a:.{dec}f}"
        req_frac = np_.split(".", 1)[1].count("0") if "." in np_ else 0
        if dec > req_frac and "." in s:
            whole, frac = s.split(".")
            frac = frac.rstrip("0")
            frac = frac + "0" * max(0, req_frac - len(frac))
            s = whole + ("." + frac if frac else ("." if np_.endswith(".") else ""))
        if np_.lstrip(",").startswith("#") and s.startswith("0") and not s.startswith("0.") and a < 1 and dec == 0:
            s = s[1:] or "0"
    clean = lambda t: re.sub(r'"([^"]*)"', r"\1", t).replace("\\", "")  # noqa: E731
    out = clean(prefix) + s + clean(suffix)
    return ("-" + out) if neg and s.strip("0.,") else out


# ---- workbook
class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}        # (row, col) -> {"raw": str, "f": {format}}
        self.colw = {}         # col -> width px at 100%
        self.rowh = {}         # row -> height px at 100%
        self.charts = []       # dicts: type, range (r1, c1, r2, c2), title, row, col, dx, dy, w, h
        self.freeze = (0, 0)   # frozen rows, cols
        self._ext = None

    def extent(self):
        """(last used row, last used col)."""
        if self._ext is None:
            if self.cells:
                self._ext = (max(k[0] for k in self.cells), max(k[1] for k in self.cells))
            else:
                self._ext = (-1, -1)
        return self._ext

    def touch(self):
        self._ext = None

    def get(self, r, c):
        return self.cells.get((r, c))

    def raw(self, r, c):
        cl = self.cells.get((r, c))
        return cl["raw"] if cl else ""

    def fmt(self, r, c):
        cl = self.cells.get((r, c))
        return cl["f"] if cl else {}


class Book:
    def __init__(self):
        self.sheets = [Sheet("Sheet1")]
        self.cache = {}
        self.busy = set()
        self.cur = None

    def sheet_named(self, name):
        n = name.lower()
        for s in self.sheets:
            if s.name.lower() == n:
                return s
        return None

    def parsed(self, cell):
        p = cell.get("p")
        if p is None:
            p = cell["p"] = parse_input(cell["raw"])
        return p

    def value(self, sh, r, c):
        key = (id(sh), r, c)
        if key in self.cache:
            return self.cache[key]
        cell = sh.cells.get((r, c))
        if cell is None:
            return None
        p = self.parsed(cell)
        if p[0] == "v":
            v = p[1]
        else:
            if key in self.busy:
                return XlErr("#CIRC!")
            self.busy.add(key)
            try:
                v = self.ev(p[1], sh)
                if isinstance(v, Rng):
                    v = scalar(v)
                elif isinstance(v, list):
                    v = v[0][0] if v and v[0] else None
                if v is None:
                    v = 0.0
            except XlRaise as e:
                v = e.err
            except RecursionError:
                v = XlErr("#CIRC!")
            except (ZeroDivisionError,):
                v = XlErr("#DIV/0!")
            except (OverflowError, ValueError, TypeError, IndexError):
                v = XlErr("#NUM!")
            finally:
                self.busy.discard(key)
        self.cache[key] = v
        return v

    def recalc(self):
        self.cache.clear()
        for sh in self.sheets:
            for (r, c) in sorted(sh.cells):
                self.value(sh, r, c)

    def resolve_sheet(self, name, cur):
        if name is None:
            return cur
        s = self.sheet_named(name)
        if s is None:
            err("#REF!")
        return s

    def ev(self, n, sh):
        k = n[0]
        if k == "num":
            return n[1]
        if k == "str":
            return n[1]
        if k == "bool":
            return n[1]
        if k == "err":
            raise XlRaise(XlErr(n[1]))
        if k == "blank":
            return None
        if k == "ref":
            s = self.resolve_sheet(n[1], sh)
            if n[2] is None or n[3] is None:
                err("#REF!")
            v = self.value(s, n[2], n[3])
            return v
        if k == "range":
            s = self.resolve_sheet(n[1], sh)
            return Rng(self, s, n[2], n[3], n[4], n[5])
        if k == "array":
            return [[self.ev(x, sh) for x in row] for row in n[1]]
        if k == "neg":
            return -to_num(scalar(self.ev(n[1], sh)))
        if k == "pct":
            return to_num(scalar(self.ev(n[1], sh))) / 100
        if k == "name":
            err("#NAME?")
        if k == "bin":
            op = n[1]
            a = scalar(self.ev(n[2], sh))
            b = scalar(self.ev(n[3], sh))
            if isinstance(a, XlErr):
                raise XlRaise(a)
            if isinstance(b, XlErr):
                raise XlRaise(b)
            if op == "&":
                return to_str(a) + to_str(b)
            if op in ("=", "<>", "<", ">", "<=", ">="):
                c = compare(a, b)
                return {"=": c == 0, "<>": c != 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]
            x, y = to_num(a), to_num(b)
            if op == "+":
                return x + y
            if op == "-":
                return x - y
            if op == "*":
                return x * y
            if op == "/":
                if y == 0:
                    err("#DIV/0!")
                return x / y
            if op == "^":
                return _pow(x, y)
        if k == "fn":
            name = n[1]
            f = FUNCS.get(name)
            if f is None:
                err("#NAME?")
            if name in LAZY:
                return f(lambda node: self.ev(node, sh), n[2])
            args = []
            for a in n[2]:
                v = self.ev(a, sh)
                args.append(v)
            return f(self, args)
        err("#VALUE!")


# ---- reference rewriting (copy/paste, fill, insert/delete rows and columns)
REF_RE = re.compile(r"""(?<![A-Za-z0-9_.$'"])((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)(?::(\$?)([A-Za-z]{1,3})(\$?)([0-9]+))?(?![A-Za-z0-9_(!])""")
STR_SPLIT = re.compile(r'("(?:[^"]|"")*")')


def rewrite_refs(formula, fn_):
    """fn_(sheet_prefix_or_None, row, col, rabs, cabs, part) -> (row, col) or None (#REF!). part: 0 single, 1 start, 2 end."""
    if not formula.startswith("="):
        return formula
    pieces = STR_SPLIT.split(formula)
    for i in range(0, len(pieces), 2):
        def sub(m):
            sh = m.group(1)
            shn = None
            if sh:
                shn = sh[:-1]
                if shn.startswith("'"):
                    shn = shn[1:-1].replace("''", "'")
            out = []
            parts = [(m.group(2), m.group(3), m.group(4), m.group(5))]
            if m.group(7):
                parts.append((m.group(6), m.group(7), m.group(8), m.group(9)))
            for pi_, (ca, cl, ra, rw) in enumerate(parts):
                r, c = int(rw) - 1, col_index(cl)
                res = fn_(shn, r, c, bool(ra), bool(ca), 0 if len(parts) == 1 else pi_ + 1)
                if res is None:
                    return (sh or "") + "#REF!"
                r2, c2 = res
                if r2 < 0 or c2 < 0 or c2 > 16383 or r2 > 1048575:
                    return (sh or "") + "#REF!"
                out.append(f"{ca}{col_name(c2)}{ra}{r2 + 1}")
            return (sh or "") + ":".join(out)
        pieces[i] = REF_RE.sub(sub, pieces[i])
    return "".join(pieces)


def shift_formula(raw, dr, dc):
    return rewrite_refs(raw, lambda sh, r, c, ra, ca, part: (r if ra else r + dr, c if ca else c + dc))


def struct_shift(raw, own_sheet, target, axis, at, n):
    """Insert (n > 0) or delete (n < 0) rows/columns at index `at` of sheet `target`."""
    def f(sh, r, c, ra, ca, part):
        s = own_sheet if sh is None else sh
        if s.lower() != target.lower():
            return r, c
        v = r if axis == 0 else c
        if n > 0:
            if v >= at:
                v += n
        else:
            d = -n
            if at <= v < at + d:
                if part == 1:
                    v = at
                elif part == 2:
                    v = at - 1
                else:
                    return None
            elif v >= at + d:
                v -= d
        return (v, c) if axis == 0 else (r, v)
    return rewrite_refs(raw, f)


def rename_sheet_refs(raw, old, new):
    if not raw.startswith("="):
        return raw
    q = lambda nm: nm if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", nm) else "'" + nm.replace("'", "''") + "'"  # noqa: E731
    pieces = STR_SPLIT.split(raw)
    for i in range(0, len(pieces), 2):
        pieces[i] = re.sub(r"(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)(?=!)",
                           lambda m: q(new) if (m.group(0)[1:-1].replace("''", "'") if m.group(0).startswith("'")
                                                else m.group(0)).lower() == old.lower() else m.group(0), pieces[i])
    return "".join(pieces)


# ----------------------------------------------------------------------------
# Excel: files
# ----------------------------------------------------------------------------
DEF_COLW = 72
DEF_ROWH = 22
DEF_FONT = "Calibri"
DEF_SIZE = 11.0


def _argb(c):
    """openpyxl color -> '#RRGGBB' or None."""
    try:
        if c is None or c.type != "rgb" or not isinstance(c.rgb, str):
            return None
        v = c.rgb
        if len(v) == 8:
            if v[:2] == "00" and v[2:] == "000000":
                return None
            v = v[2:]
        return "#" + v.upper()
    except AttributeError:
        return None


THEME_FILLS = {0: "#FFFFFF", 1: "#000000", 2: "#E7E6E6", 3: "#44546A", 4: "#4472C4", 5: "#ED7D31", 6: "#A5A5A5",
               7: "#FFC000", 8: "#5B9BD5", 9: "#70AD47"}


def _theme_color(c):
    try:
        if c is not None and c.type == "theme" and c.theme in THEME_FILLS:
            base = THEME_FILLS[c.theme]
            t = c.tint or 0
            r, g, b = (int(base[i:i + 2], 16) for i in (1, 3, 5))
            if t > 0:
                r, g, b = (int(x + (255 - x) * t) for x in (r, g, b))
            elif t < 0:
                r, g, b = (int(x * (1 + t)) for x in (r, g, b))
            return f"#{r:02X}{g:02X}{b:02X}"
    except AttributeError:
        pass
    return None


def _ref_box(f):
    """"'Sheet'!$A$1:$B$4" -> (sheet, r1, c1, r2, c2)."""
    if not f:
        return None
    sh, rest = split_sheet(f.replace("$", ""))
    a, _, b = rest.partition(":")
    pa = parse_cell_name(a)
    pb = parse_cell_name(b) if b else pa
    if not pa or not pb:
        return None
    return sh, pa[0], pa[1], pb[0], pb[1]


def load_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=False)
    book = Book()
    book.sheets = []
    for ws in wb.worksheets:
        sh = Sheet(ws.title)
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                f = {}
                fo = cell.font
                if fo is not None:
                    if fo.b:
                        f["b"] = True
                    if fo.i:
                        f["i"] = True
                    if fo.u:
                        f["u"] = True
                    if fo.strike:
                        f["s"] = True
                    if fo.name and fo.name != DEF_FONT:
                        f["font"] = fo.name
                    if fo.sz and float(fo.sz) != DEF_SIZE:
                        f["size"] = float(fo.sz)
                    col = _argb(fo.color) or _theme_color(fo.color)
                    if col and col not in ("#000000",):
                        f["fg"] = col
                fill = cell.fill
                if fill is not None and fill.fill_type == "solid":
                    col = _argb(fill.fgColor) or _theme_color(fill.fgColor)
                    if col and col != "#FFFFFF":
                        f["bg"] = col
                al = cell.alignment
                if al is not None:
                    if al.horizontal in ("left", "center", "right"):
                        f["align"] = al.horizontal
                    if al.wrap_text:
                        f["wrap"] = True
                bd = cell.border
                if bd is not None and any(getattr(bd, s) is not None and getattr(bd, s).style for s in ("left", "right", "top", "bottom")):
                    f["border"] = True
                nf = cell.number_format
                if nf and nf != "General":
                    f["num"] = nf
                if v is None and not f:
                    continue
                if v is None:
                    raw = ""
                elif isinstance(v, str):
                    if cell.data_type == "f" or (v.startswith("=") and cell.data_type != "s"):
                        raw = v
                    elif v.startswith("=") or (parse_input(v)[1] != v):
                        raw = "'" + v
                    else:
                        raw = v
                elif isinstance(v, bool):
                    raw = "TRUE" if v else "FALSE"
                elif isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
                    raw = general(to_serial(v))
                    if "num" not in f:
                        f["num"] = "m/d/yyyy"
                elif isinstance(v, (int, float)):
                    raw = general(float(v)) if not (isinstance(v, float) and not v.is_integer()) else repr(float(v))
                elif hasattr(v, "text"):   # array formulas
                    raw = str(v.text)
                    raw = raw if raw.startswith("=") else "=" + raw
                else:
                    raw = str(v)
                sh.cells[(cell.row - 1, cell.column - 1)] = {"raw": raw, "f": f}
        for letter, dim in ws.column_dimensions.items():
            if dim.width and not dim.hidden:
                try:
                    lo, hi = col_index(letter), col_index(letter)
                    if dim.min and dim.max:
                        lo, hi = dim.min - 1, min(dim.max - 1, lo + 200)
                    for c in range(lo, hi + 1):
                        sh.colw[c] = max(12, int(dim.width * 7 + 5))
                except (ValueError, TypeError):
                    pass
        for r, dim in ws.row_dimensions.items():
            if dim.height:
                sh.rowh[r - 1] = max(8, int(dim.height * 4 / 3))
        if ws.freeze_panes:
            p = parse_cell_name(ws.freeze_panes)
            if p:
                sh.freeze = p
        for ch in getattr(ws, "_charts", []):
            try:
                sh.charts.append(read_chart(ch))
            except Exception:
                pass
        sh.touch()
        book.sheets.append(sh)
    if not book.sheets:
        book.sheets = [Sheet("Sheet1")]
    try:
        book.active = wb.worksheets.index(wb.active)
    except (ValueError, AttributeError):
        book.active = 0
    return book


def read_chart(ch):
    tag = ch.tagname
    typ = {"barChart": "col", "bar3DChart": "col", "lineChart": "line", "line3DChart": "line", "pieChart": "pie",
           "pie3DChart": "pie", "doughnutChart": "pie", "areaChart": "line", "scatterChart": "line"}.get(tag, "col")
    if typ == "col" and getattr(ch, "type", "col") == "bar":
        typ = "bar"
    boxes = []
    for s in ch.series:
        for ref in (s.val.numRef.f if s.val and s.val.numRef else None,
                    (s.cat.strRef.f if s.cat and s.cat.strRef else s.cat.numRef.f if s.cat and s.cat.numRef else None),
                    s.tx.strRef.f if s.tx and s.tx.strRef else None):
            bx = _ref_box(ref)
            if bx:
                boxes.append(bx)
    if not boxes:
        raise ValueError("no data")
    r1 = min(b[1] for b in boxes)
    c1 = min(b[2] for b in boxes)
    r2 = max(b[3] for b in boxes)
    c2 = max(b[4] for b in boxes)
    title = None
    try:
        title = "".join(r.t for p in ch.title.tx.rich.p for r in (p.r or []))
    except AttributeError:
        pass
    fr = ch.anchor._from
    w = int((ch.width or 15) * 37.8)
    h = int((ch.height or 7.5) * 37.8)
    ext = getattr(ch.anchor, "ext", None)
    to = getattr(ch.anchor, "to", None)
    if ext is not None and getattr(ext, "cx", None):
        w, h = int(ext.cx / 9525), int(ext.cy / 9525)
    elif to is not None:
        w = int((to.col - fr.col) * DEF_COLW + ((to.colOff or 0) - (fr.colOff or 0)) / 9525)
        h = int((to.row - fr.row) * DEF_ROWH + ((to.rowOff or 0) - (fr.rowOff or 0)) / 9525)
    return {"type": typ, "range": (r1, c1, r2, c2), "title": title, "row": fr.row, "col": fr.col,
            "w": w or 480, "h": h or 288, "sheet": boxes[0][0]}


def save_xlsx(book, path, active=0):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    thin = Side(style="thin", color="FF000000")
    styles = {}
    for sh in book.sheets:
        ws = wb.create_sheet(sh.name[:31])
        for (r, c), cell in sorted(sh.cells.items()):
            raw, f = cell["raw"], cell["f"]
            p = book.parsed(cell)
            xc = ws.cell(row=r + 1, column=c + 1)
            if raw.startswith("=") and len(raw) > 1:
                xc.value = raw
            elif raw.startswith("'"):
                xc.value = raw[1:]
                xc.data_type = "s"
            elif p[0] == "v":
                v = p[1]
                if isinstance(v, XlErr):
                    xc.value = v.code
                    xc.data_type = "e"
                elif isinstance(v, str):
                    xc.value = v
                    if v.startswith("="):
                        xc.data_type = "s"
                else:
                    xc.value = v
                    if isinstance(v, float) and v.is_integer() and abs(v) < 2 ** 53:
                        xc.value = int(v)
            if not f:
                continue
            key = tuple(sorted(f.items()))
            st = styles.get(key)
            if st is None:
                font = Font(name=f.get("font", DEF_FONT), size=f.get("size", DEF_SIZE), bold=f.get("b", False),
                            italic=f.get("i", False), underline="single" if f.get("u") else None, strike=f.get("s", False),
                            color=("FF" + f["fg"][1:]) if f.get("fg") else None)
                fill = PatternFill("solid", fgColor="FF" + f["bg"][1:]) if f.get("bg") else None
                al = Alignment(horizontal=f.get("align"), wrap_text=f.get("wrap") or None) if (f.get("align") or f.get("wrap")) else None
                bd = Border(left=thin, right=thin, top=thin, bottom=thin) if f.get("border") else None
                st = styles[key] = (font, fill, al, bd, f.get("num"))
            font, fill, al, bd, num = st
            xc.font = font
            if fill:
                xc.fill = fill
            if al:
                xc.alignment = al
            if bd:
                xc.border = bd
            if num:
                xc.number_format = num
        for c, w in sh.colw.items():
            ws.column_dimensions[col_name(c)].width = round(max(1, (w - 5) / 7), 2)
        for r, h in sh.rowh.items():
            ws.row_dimensions[r + 1].height = round(h * 0.75, 2)
        if sh.freeze != (0, 0):
            ws.freeze_panes = cell_name(sh.freeze[0], sh.freeze[1])
        for chd in sh.charts:
            try:
                ws.add_chart(make_chart(book, sh, ws, chd, BarChart, LineChart, PieChart, Reference),
                             cell_name(chd["row"], chd["col"]))
            except Exception:
                pass
    if book.sheets:
        wb.active = max(0, min(active, len(book.sheets) - 1))
    wb.save(path)


def chart_layout(book, sh, chd):
    """Split a chart range into header row / category column / series columns."""
    r1, c1, r2, c2 = chd["range"]
    src = book.sheet_named(chd.get("sheet") or sh.name) or sh
    val = lambda r, c: book.value(src, r, c)  # noqa: E731
    is_text = lambda v: isinstance(v, str) and v != ""  # noqa: E731
    head = r2 > r1 and any(is_text(val(r1, c)) for c in range(c1 + (1 if c2 > c1 else 0), c2 + 1))
    cats_col = c2 > c1 and (any(is_text(val(r, c1)) for r in range(r1 + (1 if head else 0), r2 + 1)) or
                            (head and not is_text(val(r1, c1)) and val(r1, c1) in (None, "")))
    dr1 = r1 + (1 if head else 0)
    sc1 = c1 + (1 if cats_col else 0)
    return src, head, cats_col, dr1, sc1


def make_chart(book, sh, ws, chd, BarChart, LineChart, PieChart, Reference):
    src, head, cats_col, dr1, sc1 = chart_layout(book, sh, chd)
    r1, c1, r2, c2 = chd["range"]
    wsrc = ws.parent[src.name[:31]]
    t = chd["type"]
    if t in ("col", "bar"):
        ch = BarChart()
        ch.type = t
    elif t == "line":
        ch = LineChart()
    else:
        ch = PieChart()
    if chd.get("title"):
        ch.title = chd["title"]
    data = Reference(wsrc, min_col=sc1 + 1, max_col=c2 + 1, min_row=(r1 if head else dr1) + 1, max_row=r2 + 1)
    ch.add_data(data, titles_from_data=head)
    if cats_col:
        ch.set_categories(Reference(wsrc, min_col=c1 + 1, max_col=c1 + 1, min_row=dr1 + 1, max_row=r2 + 1))
    ch.width = chd["w"] / 37.8
    ch.height = chd["h"] / 37.8
    return ch


def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        text = f.read()
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    book = Book()
    sh = book.sheets[0]
    sh.name = os.path.splitext(os.path.basename(path))[0][:31] or "Sheet1"
    for r, row in enumerate(csv.reader(io.StringIO(text), dialect)):
        for c, v in enumerate(row):
            if v != "":
                sh.cells[(r, c)] = {"raw": v, "f": {}}
    sh.touch()
    book.active = 0
    return book


def save_csv(book, sh, path):
    mr, mc = sh.extent()
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        for r in range(mr + 1):
            w.writerow([display_text(book, sh, r, c) for c in range(mc + 1)])


def display_text(book, sh, r, c):
    v = book.value(sh, r, c)
    return format_value(v, sh.fmt(r, c).get("num"))


# ----------------------------------------------------------------------------
# Excel: window
# ----------------------------------------------------------------------------
gi.require_version("PangoCairo", "1.0")
from gi.repository import PangoCairo  # noqa: E402

APP_ID = "io.github.piexcel"
DEFAULTS = {"width": 1280, "height": 820, "max": False, "zoom": 100, "recent": [], "grid": True}
HW, HH = 46, 22
FONTS = ["Calibri", "Cambria", "Arial", "Times New Roman", "Courier New", "Georgia", "Verdana", "DejaVu Sans",
         "Liberation Sans", "Noto Sans", "Open Sans"]
FONT_SUB = {"Calibri": "Carlito", "Cambria": "Caladea", "Arial": "Liberation Sans", "Times New Roman": "Liberation Serif",
            "Courier New": "Liberation Mono", "Aptos": "Carlito", "Aptos Narrow": "Carlito"}
SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
NUM_FORMATS = [("General", "General"), ("Number", "0.00"), ("Currency", '"$"#,##0.00'),
               ("Accounting", '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'),
               ("Short date", "m/d/yyyy"), ("Long date", "dddd, mmmm d, yyyy"), ("Time", "h:mm:ss AM/PM"),
               ("Percentage", "0.00%"), ("Fraction-free whole", "#,##0"), ("Scientific", "0.00E+00"), ("Text", "@")]
PALETTE = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E"]
TEXT_COLORS = ["#000000", "#C00000", "#FF0000", "#FFC000", "#00B050", "#0070C0", "#002060", "#7030A0", "#7F7F7F", "#FFFFFF"]
FILL_COLORS = [("Yellow", "#FFFF00"), ("Light yellow", "#FFF2CC"), ("Light green", "#E2EFDA"), ("Green", "#C6EFCE"),
               ("Light blue", "#DDEBF7"), ("Blue", "#BDD7EE"), ("Light orange", "#FCE4D6"), ("Light red", "#FFC7CE"),
               ("Light gray", "#EDEDED"), ("Gray", "#D9D9D9"), ("Dark blue", "#1F4E78"), ("Black", "#000000")]
ACCENT_X = "#21A366"
SEL_GREEN = (0.13, 0.45, 0.27)

CSS = """
.ribbon-tabs { background-color: #202020; padding: 0 8px; }
button.rtab { background: transparent; border: none; box-shadow: none; border-radius: 0; padding: 6px 12px; min-height: 26px; }
button.rtab label { font-size: 12px; color: rgba(255,255,255,0.8); }
button.rtab:hover { background-color: rgba(255,255,255,0.05); }
button.rtab.sel { border-bottom: 2px solid #3FC380; }
button.rtab.sel label { color: #ffffff; font-weight: 600; }
.ribbon { background-color: #2C2C2C; padding: 4px 8px; border-bottom: 1px solid rgba(0,0,0,0.4); min-height: 40px; }
.rsep { background-color: rgba(255,255,255,0.1); min-width: 1px; margin: 6px 6px; }
button.rb { background: transparent; border: none; box-shadow: none; border-radius: 4px; min-width: 30px; min-height: 30px; padding: 0 6px; }
button.rb:hover { background-color: rgba(255,255,255,0.07); }
button.rb.on { background-color: rgba(63,195,128,0.25); }
button.rb label { font-size: 12px; color: #ffffff; }
.fbar { background-color: #292929; padding: 4px 8px; border-bottom: 1px solid rgba(0,0,0,0.5); }
entry.namebox, entry.fx { background-color: #1F1F1F; color: #ffffff; border: 1px solid rgba(255,255,255,0.1); border-radius: 4px; min-height: 26px; }
entry.fx { font-family: "Carlito", "DejaVu Sans"; }
label.fxl { color: rgba(255,255,255,0.6); font-style: italic; font-family: serif; margin: 0 6px; }
entry.celled { background-color: #ffffff; color: #000000; border: 2px solid #217346; border-radius: 0; box-shadow: none;
               min-height: 0; padding: 0 3px; caret-color: #000000; }
.sheetbar { background-color: #202020; padding: 0 6px; border-top: 1px solid rgba(0,0,0,0.5); }
button.stab { background: transparent; border: none; box-shadow: none; border-radius: 0; padding: 4px 14px; min-height: 26px; }
button.stab label { color: rgba(255,255,255,0.75); font-size: 12px; }
button.stab:hover { background-color: rgba(255,255,255,0.06); }
button.stab.sel { background-color: #2C2C2C; border-bottom: 2px solid #3FC380; }
button.stab.sel label { color: #3FC380; font-weight: 600; }
.status { background-color: #2C2C2C; }
.findbar { background-color: #2C2C2C; border: 1px solid rgba(0,0,0,0.5); border-radius: 8px; padding: 6px; }
combobox.fontbox button, combobox.sizebox button, combobox.numbox button { min-height: 26px; padding: 0 6px; }
"""


def draw_logo(cr, s):
    cr.save()
    cr.scale(s / 16.0, s / 16.0)
    _rr(cr, 5, 1.5, 10, 13, 1.5)
    cr.set_source_rgb(0.20, 0.66, 0.42)
    cr.fill()
    cr.set_source_rgba(1, 1, 1, 0.55)
    cr.set_line_width(0.8)
    for y in (5, 8, 11):
        cr.move_to(10, y)
        cr.line_to(15, y)
        cr.stroke()
    cr.move_to(12.5, 1.5)
    cr.line_to(12.5, 14.5)
    cr.stroke()
    g = cairo.LinearGradient(0, 0, 16, 16)
    g.add_color_stop_rgb(0, 0.13, 0.55, 0.32)
    g.add_color_stop_rgb(1, 0.06, 0.36, 0.20)
    _rr(cr, 1, 3.5, 9, 9, 1.4)
    cr.set_source(g)
    cr.fill()
    cr.set_source_rgb(1, 1, 1)
    cr.set_line_width(1.3)
    cr.set_line_cap(cairo.LINE_CAP_ROUND)
    cr.move_to(3.4, 5.6)
    cr.line_to(7.6, 10.4)
    cr.stroke()
    cr.move_to(7.6, 5.6)
    cr.line_to(3.4, 10.4)
    cr.stroke()
    cr.restore()


_orig_extra = draw_glyph_extra


def draw_glyph_extra(cr, name, L, pi):  # noqa: F811
    if name in ("al-left", "al-center", "al-right"):
        for w, y in ((12, 3.5), (8, 6.5), (12, 9.5), (8, 12.5)):
            x0 = {"al-left": 2, "al-center": 8 - w / 2, "al-right": 14 - w}[name]
            L((x0, y), (x0 + w, y))
    elif name == "fcolor":
        L((4, 11), (8, 2.5), (12, 11))
        L((5.6, 7.8), (10.4, 7.8))
        cr.rectangle(2.5, 12.8, 11, 2)
        cr.fill()
    elif name == "fill":
        L((3, 8), (8, 3), (12.5, 7.5), (7.5, 12.5), (3, 8))
        L((6, 1.5), (8, 3))
        cr.move_to(13.5, 9)
        cr.curve_to(15, 11, 14.5, 12.5, 13.5, 12.5)
        cr.curve_to(12.5, 12.5, 12, 11, 13.5, 9)
        cr.fill()
        cr.rectangle(2.5, 13.6, 11, 1.8)
        cr.fill()
    elif name == "borders":
        _rr(cr, 2, 2, 12, 12, 0.5)
        cr.stroke()
        cr.set_dash([1, 1.5])
        L((8, 2), (8, 14))
        L((2, 8), (14, 8))
        cr.set_dash([])
    elif name == "sigma":
        L((12, 3.5), (12, 2.5), (3.5, 2.5), (8.5, 8), (3.5, 13.5), (12, 13.5), (12, 12.5))
    elif name in ("sort-az", "sort-za"):
        cr.set_font_size(5.5)
        top, bot = ("A", "Z") if name == "sort-az" else ("Z", "A")
        cr.move_to(1.5, 6.5)
        cr.show_text(top)
        cr.move_to(1.5, 14)
        cr.show_text(bot)
        L((11.5, 2), (11.5, 14))
        L((9, 11.5), (11.5, 14), (14, 11.5))
    elif name == "wraptext":
        L((2, 3.5), (14, 3.5))
        cr.move_to(2, 8)
        cr.line_to(12, 8)
        cr.curve_to(15, 8, 15, 12.5, 12, 12.5)
        cr.line_to(8, 12.5)
        cr.stroke()
        L((10, 10.5), (8, 12.5), (10, 14.5))
        L((2, 12.5), (5, 12.5))
    elif name == "dollar":
        cr.set_font_size(11)
        cr.move_to(4.5, 12.5)
        cr.show_text("$")
    elif name == "percent":
        cr.set_font_size(10.5)
        cr.move_to(2.5, 12.5)
        cr.show_text("%")
    elif name == "comma":
        cr.set_font_size(12)
        cr.move_to(5.5, 10)
        cr.show_text(",")
    elif name in ("dec-more", "dec-less"):
        cr.set_font_size(5)
        cr.move_to(1, 7)
        cr.show_text(".0" if name == "dec-more" else ".00")
        cr.move_to(4 if name == "dec-more" else 6, 14)
        cr.show_text(".00" if name == "dec-more" else ".0")
        if name == "dec-more":
            L((2, 11.5), (4.5, 11.5))
            L((3, 10.3), (1.8, 11.5), (3, 12.7))
        else:
            L((2, 11.5), (5, 11.5))
            L((4, 10.3), (5.2, 11.5), (4, 12.7))
    elif name in ("ins-row", "del-row", "ins-col", "del-col"):
        if name.endswith("row"):
            _rr(cr, 2, 5.5, 12, 5, 0.5)
        else:
            _rr(cr, 5.5, 2, 5, 12, 0.5)
        cr.stroke()
        if name.startswith("ins"):
            L((8, 6.2), (8, 9.8))
            L((6.2, 8), (9.8, 8))
        else:
            L((6.5, 6.5), (9.5, 9.5))
            L((9.5, 6.5), (6.5, 9.5))
    elif name == "ch-col":
        for x, h in ((2.5, 6), (6.5, 10), (10.5, 4)):
            cr.rectangle(x, 14 - h, 3, h)
            cr.fill()
        L((1.5, 14.5), (14.5, 14.5))
    elif name == "ch-bar":
        for y, w in ((2.5, 7), (6.5, 11), (10.5, 5)):
            cr.rectangle(2, y, w, 3)
            cr.fill()
        L((1.5, 1.5), (1.5, 14.5))
    elif name == "ch-line":
        L((1.5, 1.5), (1.5, 14.5), (14.5, 14.5))
        L((3, 11), (6.5, 6.5), (9.5, 9), (14, 3))
    elif name == "ch-pie":
        cr.move_to(8, 8)
        cr.arc(8, 8, 6, -pi / 2, pi * 0.2)
        cr.close_path()
        cr.fill()
        cr.arc(8, 8, 6, 0, 2 * pi)
        cr.stroke()
    elif name == "fx":
        cr.set_font_size(9)
        cr.select_font_face("serif", cairo.FONT_SLANT_ITALIC, cairo.FONT_WEIGHT_NORMAL)
        cr.move_to(2.5, 12)
        cr.show_text("fx")
    elif name == "freeze":
        _rr(cr, 2, 2, 12, 12, 0.5)
        cr.stroke()
        cr.set_line_width(2)
        L((2, 6), (14, 6))
        L((6, 2), (6, 14))
    elif name == "grid":
        _rr(cr, 2, 2, 12, 12, 0.5)
        cr.stroke()
        L((2, 6), (14, 6))
        L((2, 10), (14, 10))
        L((6, 2), (6, 14))
        L((10, 2), (10, 14))
    elif name == "pdf":
        L((4, 1.5), (10, 1.5), (13, 4.5), (13, 14.5), (4, 14.5), (4, 1.5))
        L((10, 1.5), (10, 4.5), (13, 4.5))
    else:
        _orig_extra(cr, name, L, pi)


def font_for(name):
    return FONT_SUB.get(name, name)


def nice_ticks(lo, hi, n=5):
    if hi == lo:
        hi = lo + 1
    span = hi - lo
    step = 10 ** math.floor(math.log10(span / n))
    for m in (1, 2, 2.5, 5, 10):
        if span / (step * m) <= n:
            step *= m
            break
    a = math.floor(lo / step) * step
    b = math.ceil(hi / step) * step
    out = []
    v = a
    while v <= b + step / 2:
        out.append(round(v, 10))
        v += step
    return out


class ExcelWindow(Gtk.ApplicationWindow):
    def __init__(self, app, path=None):
        super().__init__(application=app)
        self.app = app
        self.st = app.settings
        self.path = None
        self.book = Book()
        self.si = 0                      # active sheet index
        self.sel = (0, 0, 0, 0)          # r1, c1, r2, c2
        self.act = (0, 0)                # active cell
        self.anchor = (0, 0)
        self.zoom = self.st.get("zoom", 100) / 100
        self.show_grid = self.st.get("grid", True)
        self.show_heads = True
        self.show_formulas = False
        self.vrows, self.vcols = 200, 40
        self.editing = None              # None | "enter" | "edit"
        self.undo_stack, self.redo_stack = [], []
        self.dirty = False
        self.clip = None
        self.drag = None
        self.chart_sel = None
        self.updating = False
        self.set_icon_name("pi-excel")
        add_cls(self, "fx")
        self.set_default_size(self.st["width"], self.st["height"])
        if self.st.get("max"):
            self.maximize()
        self.connect("delete-event", self._on_close)
        self.connect("configure-event", self._cfg)
        self.connect("key-press-event", self._win_key)

        tb = WinTitlebar(self)
        t, self.title_lb = app_title("Book1 - Excel", draw_logo)
        tb.pack_start(t)
        qs = Gtk.Box()
        qs.set_margin_start(12)
        qs.pack_start(icon_button("save", lambda w: self.save(), tip="Save (Ctrl+S)"), False, False, 0)
        qs.pack_start(icon_button("undo", lambda w: self.undo(), tip="Undo (Ctrl+Z)"), False, False, 0)
        qs.pack_start(icon_button("redo", lambda w: self.redo(), tip="Redo (Ctrl+Y)"), False, False, 0)
        tb.pack_start(qs)
        tb.show_all()

        root = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.tabs = Gtk.Box()
        add_cls(self.tabs, "ribbon-tabs")
        self.ribbon = Gtk.Stack()
        self.tab_btns = {}
        fb = Gtk.Button(label="File")
        add_cls(fb, "rtab")
        fb.connect("clicked", self.file_menu)
        self.tabs.pack_start(fb, False, False, 0)
        for key, name, builder in (("home", "Home", self.rib_home), ("insert", "Insert", self.rib_insert),
                                   ("formulas", "Formulas", self.rib_formulas), ("view", "View", self.rib_view)):
            b = Gtk.Button(label=name)
            add_cls(b, "rtab")
            b.connect("clicked", lambda w, k=key: self.show_tab(k))
            self.tabs.pack_start(b, False, False, 0)
            self.tab_btns[key] = b
            box = Gtk.Box(spacing=2)
            add_cls(box, "ribbon")
            builder(box)
            self.ribbon.add_named(box, key)
        root.pack_start(self.tabs, False, False, 0)
        root.pack_start(self.ribbon, False, False, 0)

        fbar = Gtk.Box(spacing=4)
        add_cls(fbar, "fbar")
        self.namebox = Gtk.Entry()
        add_cls(self.namebox, "namebox")
        self.namebox.set_width_chars(10)
        self.namebox.connect("activate", self._goto_name)
        fbar.pack_start(self.namebox, False, False, 0)
        fx = Gtk.Label(label="fx")
        add_cls(fx, "fxl")
        fbar.pack_start(fx, False, False, 0)
        self.fxentry = Gtk.Entry()
        add_cls(self.fxentry, "fx")
        self.fxentry.connect("activate", lambda e: self._fx_commit())
        self.fxentry.connect("key-press-event", self._fx_key)
        self.fxentry.connect("changed", self._fx_changed)
        fbar.pack_start(self.fxentry, True, True, 0)
        root.pack_start(fbar, False, False, 0)

        grid_box = Gtk.Grid()
        self.ov = Gtk.Overlay()
        self.da = Gtk.DrawingArea()
        self.da.set_can_focus(True)
        self.da.set_hexpand(True)
        self.da.set_vexpand(True)
        self.da.add_events(Gdk.EventMask.BUTTON_PRESS_MASK | Gdk.EventMask.BUTTON_RELEASE_MASK |
                           Gdk.EventMask.POINTER_MOTION_MASK | Gdk.EventMask.SCROLL_MASK | Gdk.EventMask.SMOOTH_SCROLL_MASK |
                           Gdk.EventMask.KEY_PRESS_MASK)
        self.da.connect("draw", self._draw)
        self.da.connect("button-press-event", self._press)
        self.da.connect("button-release-event", self._release)
        self.da.connect("motion-notify-event", self._motion)
        self.da.connect("scroll-event", self._scroll)
        self.da.connect("key-press-event", self._grid_key)
        self.da.connect("size-allocate", lambda w, a: self._update_adj())
        self.ov.add(self.da)
        self.editor = Gtk.Entry()
        add_cls(self.editor, "celled")
        self.editor.set_halign(Gtk.Align.START)
        self.editor.set_valign(Gtk.Align.START)
        self.editor.set_has_frame(False)
        self.editor.connect("key-press-event", self._ed_key)
        self.editor.connect("changed", self._ed_changed)
        self.editor.set_no_show_all(True)
        self.ov.add_overlay(self.editor)
        self.findbar = self._build_find()
        self.ov.add_overlay(self.findbar)
        self.hadj = Gtk.Adjustment(0, 0, 1, 20, 200, 1)
        self.vadj = Gtk.Adjustment(0, 0, 1, 20, 200, 1)
        self.hadj.connect("value-changed", self._scrolled)
        self.vadj.connect("value-changed", self._scrolled)
        vs = Gtk.Scrollbar(orientation=Gtk.Orientation.VERTICAL, adjustment=self.vadj)
        grid_box.attach(self.ov, 0, 0, 1, 1)
        grid_box.attach(vs, 1, 0, 1, 1)
        root.pack_start(grid_box, True, True, 0)

        sb = Gtk.Box()
        add_cls(sb, "sheetbar")
        self.sheet_tabs = Gtk.Box()
        sb.pack_start(self.sheet_tabs, False, False, 0)
        sb.pack_start(icon_button("plus", lambda w: self.add_sheet(), tip="New sheet"), False, False, 4)
        hs = Gtk.Scrollbar(orientation=Gtk.Orientation.HORIZONTAL, adjustment=self.hadj)
        hs.set_valign(Gtk.Align.CENTER)
        sb.pack_start(hs, True, True, 12)
        root.pack_start(sb, False, False, 0)

        st = Gtk.Box(spacing=16)
        add_cls(st, "status")
        self.st_mode = label("Ready", "sec")
        st.pack_start(self.st_mode, False, False, 0)
        self.zoom_lb = label("", "sec")
        st.pack_end(self.zoom_lb, False, False, 0)
        self.zoom_sc = Gtk.Scale.new_with_range(Gtk.Orientation.HORIZONTAL, 50, 200, 10)
        self.zoom_sc.set_draw_value(False)
        self.zoom_sc.set_size_request(140, -1)
        self.zoom_sc.set_value(self.zoom * 100)
        self.zoom_sc.connect("value-changed", lambda w: self.set_zoom(w.get_value()))
        st.pack_end(self.zoom_sc, False, False, 0)
        self.st_stats = label("", "sec")
        st.pack_end(self.st_stats, False, False, 12)
        root.pack_start(st, False, False, 0)
        self.add(root)
        root.show_all()
        self.findbar.hide()
        self.show_tab("home")
        self.zoom_lb.set_text(f"{int(self.zoom * 100)}%")
        self.load_book(Book(), None)
        if path:
            GLib.idle_add(lambda: (self.open_path(path), False)[1])
        GLib.idle_add(lambda: (self.da.grab_focus(), False)[1])

    # ================================================================ geometry
    @property
    def sh(self):
        return self.book.sheets[self.si]

    def cw(self, c):
        return self.sh.colw.get(c, DEF_COLW) * self.zoom

    def rh(self, r):
        return self.sh.rowh.get(r, DEF_ROWH) * self.zoom

    def col_x(self, c):
        z = self.zoom
        return (c * DEF_COLW + sum(w - DEF_COLW for k, w in self.sh.colw.items() if k < c)) * z

    def row_y(self, r):
        z = self.zoom
        return (r * DEF_ROWH + sum(h - DEF_ROWH for k, h in self.sh.rowh.items() if k < r)) * z

    def col_at(self, x):
        c = max(0, int(x / (DEF_COLW * self.zoom)))
        while c > 0 and self.col_x(c) > x:
            c -= 1
        while self.col_x(c + 1) <= x:
            c += 1
        return c

    def row_at(self, y):
        r = max(0, int(y / (DEF_ROWH * self.zoom)))
        while r > 0 and self.row_y(r) > y:
            r -= 1
        while self.row_y(r + 1) <= y:
            r += 1
        return r

    def hw(self):
        return HW if self.show_heads else 0

    def hh(self):
        return HH if self.show_heads else 0

    def frz(self):
        fr, fc = self.sh.freeze
        return fr, fc, self.col_x(fc), self.row_y(fr)

    def sx(self, c):
        fr, fc, FX, FY = self.frz()
        return self.hw() + self.col_x(c) - (0 if c < fc else self.hadj.get_value())

    def sy(self, r):
        fr, fc, FX, FY = self.frz()
        return self.hh() + self.row_y(r) - (0 if r < fr else self.vadj.get_value())

    def cell_at(self, x, y, clamp=True):
        fr, fc, FX, FY = self.frz()
        xx, yy = x - self.hw(), y - self.hh()
        if clamp:
            xx, yy = max(0, xx), max(0, yy)
        c = self.col_at(xx) if xx < FX else self.col_at(xx + self.hadj.get_value())
        r = self.row_at(yy) if yy < FY else self.row_at(yy + self.vadj.get_value())
        return r, c

    def cell_rect(self, r, c):
        return self.sx(c), self.sy(r), self.cw(c), self.rh(r)

    def _update_adj(self):
        a = self.da.get_allocation()
        fr, fc, FX, FY = self.frz()
        mr, mc = self.sh.extent()
        self.vrows = max(self.vrows, mr + 50)
        self.vcols = max(self.vcols, mc + 10)
        pw = max(1, a.width - self.hw() - FX)
        ph = max(1, a.height - self.hh() - FY)
        tw = self.col_x(self.vcols) - FX
        th = self.row_y(self.vrows) - FY
        self.updating = True
        self.hadj.configure(min(self.hadj.get_value(), max(0, tw - pw)), 0, max(tw, pw), 40 * self.zoom, pw * 0.9, pw)
        self.vadj.configure(min(self.vadj.get_value(), max(0, th - ph)), 0, max(th, ph), 3 * DEF_ROWH * self.zoom, ph * 0.9, ph)
        self.updating = False

    def _scrolled(self, adj):
        if self.updating:
            return
        # grow the sheet as the user scrolls to the end
        if self.vadj.get_value() + self.vadj.get_page_size() >= self.vadj.get_upper() - 2 * DEF_ROWH:
            self.vrows += 100
            self._update_adj()
        if self.hadj.get_value() + self.hadj.get_page_size() >= self.hadj.get_upper() - 2 * DEF_COLW:
            self.vcols += 10
            self._update_adj()
        self._place_editor()
        self.da.queue_draw()

    def ensure_visible(self, r, c):
        a = self.da.get_allocation()
        fr, fc, FX, FY = self.frz()
        while r >= self.vrows - 5:
            self.vrows += 100
        while c >= self.vcols - 3:
            self.vcols += 10
        self._update_adj()
        if c >= fc:
            x0 = self.col_x(c) - FX
            x1 = x0 + self.cw(c)
            pw = a.width - self.hw() - FX
            if x0 < self.hadj.get_value():
                self.hadj.set_value(x0)
            elif x1 > self.hadj.get_value() + pw:
                self.hadj.set_value(x1 - pw)
        if r >= fr:
            y0 = self.row_y(r) - FY
            y1 = y0 + self.rh(r)
            ph = a.height - self.hh() - FY
            if y0 < self.vadj.get_value():
                self.vadj.set_value(y0)
            elif y1 > self.vadj.get_value() + ph:
                self.vadj.set_value(y1 - ph)

    # ================================================================ drawing
    def _draw(self, w, cr):
        a = w.get_allocation()
        W, H = a.width, a.height
        hw, hh = self.hw(), self.hh()
        cr.set_source_rgb(1, 1, 1)
        cr.paint()
        fr, fc, FX, FY = self.frz()
        sxv, syv = self.hadj.get_value(), self.vadj.get_value()
        c_first = self.col_at(FX + sxv)
        r_first = self.row_at(FY + syv)
        c_last = c_first
        while self.sx(c_last) < W and c_last < c_first + 400:
            c_last += 1
        r_last = r_first
        while self.sy(r_last) < H and r_last < r_first + 600:
            r_last += 1
        regions = [((r_first, r_last), (c_first, c_last), (hw + FX, hh + FY, W, H))]
        if fr:
            regions.append(((0, fr - 1), (c_first, c_last), (hw + FX, hh, W, hh + FY)))
        if fc:
            regions.append(((r_first, r_last), (0, fc - 1), (hw, hh + FY, hw + FX, H)))
        if fr and fc:
            regions.append(((0, fr - 1), (0, fc - 1), (hw, hh, hw + FX, hh + FY)))
        for rows, cols, clip in regions:
            cr.save()
            cr.rectangle(clip[0], clip[1], clip[2] - clip[0], clip[3] - clip[1])
            cr.clip()
            self._draw_cells(cr, rows, cols)
            cr.restore()
        # selection, copy marquee, charts
        cr.save()
        cr.rectangle(hw, hh, W - hw, H - hh)
        cr.clip()
        self._draw_selection(cr)
        self._draw_charts(cr)
        cr.restore()
        if fr or fc:
            cr.set_source_rgb(0.6, 0.6, 0.6)
            cr.set_line_width(1)
            if fr:
                cr.move_to(hw, hh + FY + 0.5)
                cr.line_to(W, hh + FY + 0.5)
            if fc:
                cr.move_to(hw + FX + 0.5, hh)
                cr.line_to(hw + FX + 0.5, H)
            cr.stroke()
        if self.show_heads:
            self._draw_headers(cr, W, H, (c_first, c_last), (r_first, r_last))

    def _font_desc(self, f):
        d = Pango.FontDescription()
        d.set_family(font_for(f.get("font", DEF_FONT)))
        d.set_size(int(f.get("size", DEF_SIZE) * self.zoom * Pango.SCALE))
        if f.get("b"):
            d.set_weight(Pango.Weight.BOLD)
        if f.get("i"):
            d.set_style(Pango.Style.ITALIC)
        return d

    def _layout(self, cr, text, f):
        lay = PangoCairo.create_layout(cr)
        lay.set_font_description(self._font_desc(f))
        attrs = Pango.AttrList()
        if f.get("u"):
            attrs.insert(Pango.attr_underline_new(Pango.Underline.SINGLE))
        if f.get("s"):
            attrs.insert(Pango.attr_strikethrough_new(True))
        lay.set_attributes(attrs)
        lay.set_text(text, -1)
        return lay

    def cell_text(self, r, c):
        sh = self.sh
        cell = sh.cells.get((r, c))
        if cell is None:
            return "", None
        if self.show_formulas and cell["raw"].startswith("="):
            return cell["raw"], "f"
        v = self.book.value(sh, r, c)
        return format_value(v, cell["f"].get("num")), v

    def _draw_cells(self, cr, rows, cols):
        sh = self.sh
        r0, r1 = rows
        c0, c1 = cols
        xs = {c: self.sx(c) for c in range(c0, c1 + 2)}
        ys = {r: self.sy(r) for r in range(r0, r1 + 2)}
        # fills
        for (r, c), cell in sh.cells.items():
            if r0 <= r <= r1 and c0 <= c <= c1 and cell["f"].get("bg"):
                cr.set_source_rgba(*hexc(cell["f"]["bg"]))
                cr.rectangle(xs[c], ys[r], xs[c + 1] - xs[c], ys[r + 1] - ys[r])
                cr.fill()
        # gridlines
        if self.show_grid:
            cr.set_source_rgb(0.86, 0.86, 0.86)
            cr.set_line_width(1)
            for c in range(c0, c1 + 2):
                x = int(xs[c]) - 0.5
                cr.move_to(x, ys[r0])
                cr.line_to(x, ys[r1 + 1])
            for r in range(r0, r1 + 2):
                y = int(ys[r]) - 0.5
                cr.move_to(xs[c0], y)
                cr.line_to(xs[c1 + 1], y)
            cr.stroke()
        # text
        pad = 3 * self.zoom
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                cell = sh.cells.get((r, c))
                if cell is None:
                    continue
                text, v = self.cell_text(r, c)
                if text == "":
                    continue
                f = cell["f"]
                x, y = xs[c], ys[r]
                w, h = xs[c + 1] - x, ys[r + 1] - y
                lay = self._layout(cr, text, f)
                isnum = isinstance(v, (int, float)) and not isinstance(v, bool) and v != "f"
                align = f.get("align") or ("right" if isnum else "center" if isinstance(v, (bool, XlErr)) else "left")
                if f.get("wrap"):
                    lay.set_width(int(max(1, w - 2 * pad) * Pango.SCALE))
                    lay.set_wrap(Pango.WrapMode.WORD_CHAR)
                    lay.set_alignment({"left": Pango.Alignment.LEFT, "center": Pango.Alignment.CENTER,
                                       "right": Pango.Alignment.RIGHT}[align])
                tw, th = lay.get_pixel_size()
                cr.save()
                if isnum and tw > w - 2 * pad and not f.get("wrap"):
                    lay.set_text("#" * max(1, int((w - 2 * pad) / max(1, tw / max(1, len(text))))), -1)
                    tw, th = lay.get_pixel_size()
                # overflow into empty neighbours
                x_end = x + w
                if not isnum and not f.get("wrap") and align == "left" and tw > w - 2 * pad:
                    cc = c + 1
                    while x_end - x < tw + 2 * pad and cc <= c1 + 30 and (r, cc) not in sh.cells:
                        nxt = self.sx(cc) + self.cw(cc)
                        cr.set_source_rgb(1, 1, 1)
                        cr.rectangle(x_end - 1, y, nxt - x_end + 1, h - 1)
                        cr.fill()
                        x_end = nxt
                        cc += 1
                cr.rectangle(x, y, x_end - x - 1, h)
                cr.clip()
                if f.get("wrap"):
                    tx = x + pad
                else:
                    tx = {"left": x + pad, "right": x + w - pad - tw, "center": x + (w - tw) / 2}[align]
                ty = y + h - th - 2 * self.zoom if th < h else y
                cr.set_source_rgba(*hexc(f.get("fg", "#000000")))
                if isinstance(v, XlErr) and not f.get("fg"):
                    cr.set_source_rgb(0.75, 0.0, 0.0)
                cr.move_to(tx, ty)
                PangoCairo.show_layout(cr, lay)
                cr.restore()
        # borders
        cr.set_source_rgb(0, 0, 0)
        cr.set_line_width(1)
        for (r, c), cell in sh.cells.items():
            if r0 <= r <= r1 and c0 <= c <= c1 and cell["f"].get("border"):
                cr.rectangle(int(xs[c]) - 0.5, int(ys[r]) - 0.5, int(xs[c + 1]) - int(xs[c]), int(ys[r + 1]) - int(ys[r]))
        cr.stroke()

    def _draw_headers(self, cr, W, H, cols, rows):
        fr, fc, FX, FY = self.frz()
        r1s, c1s, r2s, c2s = self.sel
        cr.set_source_rgb(0.96, 0.96, 0.96)
        cr.rectangle(0, 0, W, HH)
        cr.rectangle(0, 0, HW, H)
        cr.fill()
        lay = PangoCairo.create_layout(cr)
        lay.set_font_description(Pango.FontDescription("Carlito 10"))
        for part, (lo, hi), clip in (("c", (0, fc - 1), (HW, HW + FX)), ("c", cols, (HW + FX, W))):
            cr.save()
            cr.rectangle(clip[0], 0, clip[1] - clip[0], HH)
            cr.clip()
            for c in range(lo, hi + 1):
                x, w = self.sx(c), self.cw(c)
                insel = c1s <= c <= c2s
                if insel:
                    cr.set_source_rgb(0.82, 0.88, 0.85)
                    cr.rectangle(x, 0, w, HH)
                    cr.fill()
                    cr.set_source_rgb(*SEL_GREEN)
                    cr.rectangle(x, HH - 2, w, 2)
                    cr.fill()
                cr.set_source_rgb(0.78, 0.78, 0.78)
                cr.rectangle(int(x + w) - 0.5, 3, 1, HH - 6)
                cr.fill()
                lay.set_text(col_name(c), -1)
                tw, th = lay.get_pixel_size()
                cr.set_source_rgb(*SEL_GREEN) if insel else cr.set_source_rgb(0.3, 0.3, 0.3)
                cr.move_to(x + (w - tw) / 2, (HH - th) / 2)
                PangoCairo.show_layout(cr, lay)
            cr.restore()
        for (lo, hi), clip in (((0, fr - 1), (HH, HH + FY)), (rows, (HH + FY, H))):
            cr.save()
            cr.rectangle(0, clip[0], HW, clip[1] - clip[0])
            cr.clip()
            for r in range(lo, hi + 1):
                y, h = self.sy(r), self.rh(r)
                insel = r1s <= r <= r2s
                if insel:
                    cr.set_source_rgb(0.82, 0.88, 0.85)
                    cr.rectangle(0, y, HW, h)
                    cr.fill()
                    cr.set_source_rgb(*SEL_GREEN)
                    cr.rectangle(HW - 2, y, 2, h)
                    cr.fill()
                cr.set_source_rgb(0.78, 0.78, 0.78)
                cr.rectangle(3, int(y + h) - 0.5, HW - 6, 1)
                cr.fill()
                lay.set_text(str(r + 1), -1)
                tw, th = lay.get_pixel_size()
                cr.set_source_rgb(*SEL_GREEN) if insel else cr.set_source_rgb(0.3, 0.3, 0.3)
                cr.move_to(HW - tw - 6, y + (h - th) / 2)
                PangoCairo.show_layout(cr, lay)
            cr.restore()
        cr.set_source_rgb(0.75, 0.75, 0.75)
        cr.set_line_width(1)
        cr.move_to(0, HH - 0.5)
        cr.line_to(W, HH - 0.5)
        cr.move_to(HW - 0.5, 0)
        cr.line_to(HW - 0.5, H)
        cr.stroke()
        cr.set_source_rgb(0.7, 0.7, 0.7)
        cr.move_to(HW - 4, HH - 14)
        cr.line_to(HW - 4, HH - 4)
        cr.line_to(HW - 14, HH - 4)
        cr.close_path()
        cr.fill()

    def sel_rect(self, r1, c1, r2, c2):
        mr = min(r2, self.vrows + 2)
        mc = min(c2, self.vcols + 2)
        x1, y1 = self.sx(c1), self.sy(r1)
        x2, y2 = self.sx(mc) + self.cw(mc), self.sy(mr) + self.rh(mr)
        return x1, y1, x2 - x1, y2 - y1

    def _draw_selection(self, cr):
        r1, c1, r2, c2 = self.sel
        x, y, w, h = self.sel_rect(r1, c1, r2, c2)
        ar, ac = self.act
        if (r1, c1) != (r2, c2):
            cr.set_source_rgba(0.13, 0.45, 0.27, 0.12)
            cr.rectangle(x, y, w, h)
            ax, ay, aw, ah = self.cell_rect(ar, ac)
            cr.rectangle(ax + aw, ay, -aw, ah)  # active cell stays clear (reverse winding)
            cr.set_fill_rule(cairo.FILL_RULE_EVEN_ODD)
            cr.fill()
            cr.set_fill_rule(cairo.FILL_RULE_WINDING)
        cr.set_source_rgb(*SEL_GREEN)
        cr.set_line_width(2)
        cr.rectangle(int(x) - 0.5, int(y) - 0.5, int(w), int(h))
        cr.stroke()
        if not self.editing:
            cr.rectangle(x + w - 4, y + h - 4, 7, 7)
            cr.set_source_rgb(1, 1, 1)
            cr.fill_preserve()
            cr.set_source_rgb(*SEL_GREEN)
            cr.set_line_width(1)
            cr.rectangle(x + w - 3, y + h - 3, 5, 5)
            cr.fill()
        if self.drag and self.drag[0] == "fill" and self.drag[2]:
            fr1, fc1, fr2, fc2 = self.drag[2]
            fx, fy, fw, fh = self.sel_rect(fr1, fc1, fr2, fc2)
            cr.set_source_rgb(0.4, 0.4, 0.4)
            cr.set_dash([3, 3])
            cr.set_line_width(1)
            cr.rectangle(int(fx) + 0.5, int(fy) + 0.5, int(fw) - 1, int(fh) - 1)
            cr.stroke()
            cr.set_dash([])
        if self.clip and self.clip["sheet"] is self.sh and not self.clip.get("done"):
            cx, cy, cw_, ch_ = self.sel_rect(*self.clip["box"])
            cr.set_source_rgb(*SEL_GREEN)
            cr.set_dash([4, 3])
            cr.set_line_width(1.5)
            cr.rectangle(int(cx) + 0.5, int(cy) + 0.5, int(cw_) - 1, int(ch_) - 1)
            cr.stroke()
            cr.set_dash([])
        if self.editing and self.editor.get_text().startswith("="):
            # color the cells the formula refers to
            cols = [(0.27, 0.45, 0.77), (0.78, 0.29, 0.12), (0.44, 0.19, 0.63), (0.13, 0.55, 0.3), (0.75, 0.55, 0.0)]
            for i, m in enumerate(REF_RE.finditer(self.editor.get_text())):
                if m.group(1):
                    continue
                a = parse_cell_name(m.group(2) + m.group(3) + m.group(4) + m.group(5))
                b = parse_cell_name(m.group(6) + m.group(7) + m.group(8) + m.group(9)) if m.group(7) else a
                if not a or not b:
                    continue
                rx, ry, rw, rh = self.sel_rect(min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))
                col = cols[i % len(cols)]
                cr.set_source_rgba(*col, 0.08)
                cr.rectangle(rx, ry, rw, rh)
                cr.fill()
                cr.set_source_rgb(*col)
                cr.set_line_width(1.5)
                cr.rectangle(int(rx) + 0.5, int(ry) + 0.5, int(rw) - 1, int(rh) - 1)
                cr.stroke()

    # ---------------------------------------------------------------- charts
    def chart_rect(self, chd):
        z = self.zoom
        return self.sx(chd["col"]), self.sy(chd["row"]), chd["w"] * z, chd["h"] * z

    def chart_data(self, chd):
        src, head, cats_col, dr1, sc1 = chart_layout(self.book, self.sh, chd)
        r1, c1, r2, c2 = chd["range"]
        mr, mc = src.extent()
        r2, c2 = min(r2, max(mr, r1)), min(c2, max(mc, c1))
        cats = [format_value(self.book.value(src, r, c1), src.fmt(r, c1).get("num")) if cats_col else str(r - dr1 + 1)
                for r in range(dr1, r2 + 1)]
        series = []
        for c in range(sc1, c2 + 1):
            name = to_str_safe(self.book.value(src, r1, c)) if head else f"Series{c - sc1 + 1}"
            vals = []
            for r in range(dr1, r2 + 1):
                v = self.book.value(src, r, c)
                vals.append(float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0)
            series.append((name, vals))
        return cats, series

    def _draw_charts(self, cr):
        for i, chd in enumerate(self.sh.charts):
            x, y, w, h = self.chart_rect(chd)
            cr.save()
            cr.translate(x, y)
            draw_chart(cr, chd, *self.chart_data(chd), w, h, self.zoom)
            if self.chart_sel == i:
                cr.set_source_rgb(*SEL_GREEN)
                cr.set_line_width(2)
                cr.rectangle(0, 0, w, h)
                cr.stroke()
                for hx, hy in ((0, 0), (w, 0), (0, h), (w, h), (w / 2, 0), (w / 2, h), (0, h / 2), (w, h / 2)):
                    cr.arc(hx, hy, 4, 0, 2 * math.pi)
                    cr.set_source_rgb(1, 1, 1)
                    cr.fill_preserve()
                    cr.set_source_rgb(*SEL_GREEN)
                    cr.set_line_width(1)
                    cr.stroke()
            cr.restore()

    def chart_hit(self, x, y):
        for i in range(len(self.sh.charts) - 1, -1, -1):
            cx, cy, w, h = self.chart_rect(self.sh.charts[i])
            if cx - 4 <= x <= cx + w + 4 and cy - 4 <= y <= cy + h + 4:
                corner = x >= cx + w - 12 and y >= cy + h - 12
                return i, corner
        return None, False

    # ================================================================ selection
    def select(self, r1, c1, r2=None, c2=None, act=None, keep_anchor=False):
        r2 = r1 if r2 is None else r2
        c2 = c1 if c2 is None else c2
        self.sel = (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
        if act is not None:
            self.act = act
        elif not keep_anchor:
            self.act = (r1, c1)
        if not keep_anchor:
            self.anchor = self.act
        self.chart_sel = None
        self._sel_changed()

    def _sel_changed(self):
        r, c = self.act
        r1, c1, r2, c2 = self.sel
        if (r1, c1) == (r2, c2):
            self.namebox.set_text(cell_name(r, c))
        elif r1 == 0 and r2 >= 10 ** 6:
            self.namebox.set_text(f"{col_name(c1)}:{col_name(c2)}")
        elif c1 == 0 and c2 >= 10 ** 4:
            self.namebox.set_text(f"{r1 + 1}:{r2 + 1}")
        else:
            self.namebox.set_text(f"{r2 - r1 + 1}R x {c2 - c1 + 1}C")
        if not self.editing:
            self.updating = True
            self.fxentry.set_text(self.sh.raw(r, c))
            self.updating = False
        self._update_ribbon_state()
        self._update_stats()
        self.da.queue_draw()

    def _update_stats(self):
        r1, c1, r2, c2 = self.clamped_sel()
        if (r1, c1) == (r2, c2):
            self.st_stats.set_text("")
            return
        nums, count = [], 0
        for (r, c) in list(self.sh.cells):
            if r1 <= r <= r2 and c1 <= c <= c2:
                v = self.book.value(self.sh, r, c)
                if v is None or v == "":
                    continue
                count += 1
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    nums.append(v)
        parts = []
        if nums:
            fmt = self.sh.fmt(*self.act).get("num")
            parts.append(f"Average: {format_value(sum(nums) / len(nums), fmt)}")
        if count:
            parts.append(f"Count: {count}")
        if nums:
            parts.append(f"Sum: {format_value(sum(nums), self.sh.fmt(*self.act).get('num'))}")
        self.st_stats.set_text("     ".join(parts))

    def clamped_sel(self):
        r1, c1, r2, c2 = self.sel
        mr, mc = self.sh.extent()
        return r1, c1, min(r2, max(mr, r1)), min(c2, max(mc, c1))

    def move(self, dr, dc, extend=False, jump=False):
        if extend:
            r, c = self.sel_far()
        else:
            r, c = self.act
        if jump:
            r, c = self.jump(r, c, dr, dc)
        else:
            r, c = max(0, r + dr), max(0, c + dc)
        if extend:
            ar, ac = self.anchor
            self.sel = (min(ar, r), min(ac, c), max(ar, r), max(ac, c))
            self._far = (r, c)
            self.ensure_visible(r, c)
            self._sel_changed()
        else:
            self.select(r, c)
            self.ensure_visible(r, c)

    def sel_far(self):
        far = getattr(self, "_far", None)
        r1, c1, r2, c2 = self.sel
        if far and r1 <= far[0] <= r2 and c1 <= far[1] <= c2:
            return far
        ar, ac = self.anchor
        return (r2 if ar == r1 else r1, c2 if ac == c1 else c1)

    def jump(self, r, c, dr, dc):
        """Ctrl+arrow: edge of the data block."""
        has = lambda rr, cc: (rr, cc) in self.sh.cells and self.sh.raw(rr, cc) != ""  # noqa: E731
        mr, mc = self.sh.extent()
        lim_r, lim_c = max(mr + 1, r), max(mc + 1, c)
        nr, nc = r + dr, c + dc
        if nr < 0 or nc < 0:
            return r, c
        if has(r, c) and has(nr, nc):
            while 0 <= nr + dr and 0 <= nc + dc and has(nr + dr, nc + dc):
                nr, nc = nr + dr, nc + dc
            return nr, nc
        while 0 <= nr <= lim_r and 0 <= nc <= lim_c and not has(nr, nc):
            nr, nc = nr + dr, nc + dc
        if not (0 <= nr <= lim_r and 0 <= nc <= lim_c):
            return (0 if dr < 0 else r if dr == 0 else lim_r), (0 if dc < 0 else c if dc == 0 else lim_c)
        return nr, nc

    def _goto_name(self, e):
        t = e.get_text().strip().upper()
        if ":" in t:
            a, b = t.split(":", 1)
            pa, pb = parse_cell_name(a), parse_cell_name(b)
            if pa and pb:
                self.select(pa[0], pa[1], pb[0], pb[1])
                self.ensure_visible(*pa)
        else:
            p = parse_cell_name(t)
            if p:
                self.select(*p)
                self.ensure_visible(*p)
        self.da.grab_focus()

    # ================================================================ mouse
    def _press(self, w, ev):
        self.da.grab_focus() if not self.editing else None
        x, y = ev.x, ev.y
        hw, hh = self.hw(), self.hh()
        if ev.type == Gdk.EventType._2BUTTON_PRESS and ev.button == 1:
            i, corner = self.chart_hit(x, y)
            if i is not None:
                self.edit_chart_title(i)
                return True
            edge = self._header_edge(x, y)
            if edge:
                self.autofit(*edge)
                return True
            if x > hw and y > hh:
                self.start_edit(keep=True)
            return True
        if ev.button == 3:
            i, _ = self.chart_hit(x, y)
            if i is not None:
                self.chart_sel = i
                self.da.queue_draw()
                self.chart_menu(ev)
                return True
            if x > hw and y > hh:
                r, c = self.cell_at(x, y)
                r1, c1, r2, c2 = self.sel
                if not (r1 <= r <= r2 and c1 <= c <= c2):
                    self.select(r, c)
            self.cell_menu(ev)
            return True
        if ev.button != 1:
            return False
        # formula point mode: clicking a cell while typing a formula inserts its reference
        if self.editing and self.ref_insertable():
            if x > hw and y > hh:
                r, c = self.cell_at(x, y)
                self._ref_start = (r, c)
                self._ref_pos = self.editor.get_position()
                self._insert_ref(r, c, r, c)
                self.drag = ("ref",)
                return True
        if self.editing:
            self.commit_edit()
        i, corner = self.chart_hit(x, y)
        if i is not None:
            self.chart_sel = i
            chd = self.sh.charts[i]
            self.push_undo()
            self.drag = ("chart-size" if corner else "chart", i, x, y, chd["row"], chd["col"], chd["w"], chd["h"])
            self.da.queue_draw()
            return True
        self.chart_sel = None
        edge = self._header_edge(x, y)
        if edge:
            self.push_undo()
            kind, idx = edge
            self.drag = ("resize", kind, idx, x, y, (self.sh.colw.get(idx, DEF_COLW) if kind == "c" else self.sh.rowh.get(idx, DEF_ROWH)))
            return True
        if y < hh and x > hw:
            r, c = self.cell_at(x, y)
            if ev.state & Gdk.ModifierType.SHIFT_MASK:
                c0 = self.anchor[1]
                self.select(0, min(c0, c), 10 ** 7, max(c0, c), keep_anchor=True)
            else:
                self.select(0, c, 10 ** 7, c, act=(0, c))
            self.drag = ("cols", c)
            return True
        if x < hw and y > hh:
            r, c = self.cell_at(x, y)
            self.select(r, 0, r, 10 ** 5, act=(r, 0))
            self.drag = ("rows", r)
            return True
        if x < hw and y < hh:
            self.select(0, 0, 10 ** 7, 10 ** 5, act=(0, 0))
            return True
        # fill handle
        sx_, sy_, sw, shh = self.sel_rect(*self.sel)
        if abs(x - (sx_ + sw)) <= 5 and abs(y - (sy_ + shh)) <= 5:
            self.drag = ("fill", self.sel, None)
            return True
        r, c = self.cell_at(x, y)
        if ev.state & Gdk.ModifierType.SHIFT_MASK:
            ar, ac = self.anchor
            self.sel = (min(ar, r), min(ac, c), max(ar, r), max(ac, c))
            self._sel_changed()
        else:
            self.select(r, c)
        self.drag = ("sel",)
        return True

    def _header_edge(self, x, y):
        hw, hh = self.hw(), self.hh()
        if not self.show_heads:
            return None
        if y < hh and x > hw:
            r, c = self.cell_at(x, y)
            right = self.sx(c) + self.cw(c)
            if abs(x - right) <= 4:
                return ("c", c)
            if c > 0 and abs(x - self.sx(c)) <= 4:
                return ("c", c - 1)
        if x < hw and y > hh:
            r, c = self.cell_at(x, y)
            bot = self.sy(r) + self.rh(r)
            if abs(y - bot) <= 3:
                return ("r", r)
            if r > 0 and abs(y - self.sy(r)) <= 3:
                return ("r", r - 1)
        return None

    def _motion(self, w, ev):
        x, y = ev.x, ev.y
        win = self.da.get_window()
        if self.drag is None:
            edge = self._header_edge(x, y)
            sx_, sy_, sw, shh = self.sel_rect(*self.sel)
            name = None
            if edge:
                name = "col-resize" if edge[0] == "c" else "row-resize"
            elif abs(x - (sx_ + sw)) <= 5 and abs(y - (sy_ + shh)) <= 5:
                name = "crosshair"
            else:
                i, corner = self.chart_hit(x, y)
                if i is not None:
                    name = "nwse-resize" if corner else "move"
            win.set_cursor(Gdk.Cursor.new_from_name(self.da.get_display(), name) if name else None)
            return False
        k = self.drag[0]
        if k == "sel":
            r, c = self.cell_at(x, y)
            ar, ac = self.anchor
            self.sel = (min(ar, r), min(ac, c), max(ar, r), max(ac, c))
            self._far = (r, c)
            self._autoscroll(x, y)
            self._sel_changed()
        elif k == "cols":
            r, c = self.cell_at(x, y)
            c0 = self.drag[1]
            self.sel = (0, min(c0, c), 10 ** 7, max(c0, c))
            self._sel_changed()
        elif k == "rows":
            r, c = self.cell_at(x, y)
            r0 = self.drag[1]
            self.sel = (min(r0, r), 0, max(r0, r), 10 ** 5)
            self._sel_changed()
        elif k == "resize":
            _, kind, idx, x0, y0, base = self.drag
            if kind == "c":
                self.sh.colw[idx] = max(8, int(base + (x - x0) / self.zoom))
            else:
                self.sh.rowh[idx] = max(8, int(base + (y - y0) / self.zoom))
            self.mark_dirty()
            self._update_adj()
            self.da.queue_draw()
        elif k == "fill":
            r, c = self.cell_at(x, y)
            r1, c1, r2, c2 = self.drag[1]
            if r > r2 and r - r2 >= abs(c - c2 if c > c2 else c1 - c if c < c1 else 0):
                tgt = (r1, c1, r, c2)
            elif r < r1 and r1 - r >= abs(c - c2 if c > c2 else c1 - c if c < c1 else 0):
                tgt = (r, c1, r2, c2)
            elif c > c2:
                tgt = (r1, c1, r2, c)
            elif c < c1:
                tgt = (r1, c, r2, c2)
            else:
                tgt = None
            self.drag = ("fill", self.drag[1], tgt)
            self._autoscroll(x, y)
            self.da.queue_draw()
        elif k == "ref":
            r, c = self.cell_at(x, y)
            r0, c0 = self._ref_start
            self._insert_ref(min(r0, r), min(c0, c), max(r0, r), max(c0, c))
        elif k in ("chart", "chart-size"):
            _, i, x0, y0, row0, col0, w0, h0 = self.drag
            chd = self.sh.charts[i]
            if k == "chart":
                ox, oy = self.sx(col0) + (x - x0), self.sy(row0) + (y - y0)
                r, c = self.cell_at(ox + 1, oy + 1)
                chd["row"], chd["col"] = r, c
            else:
                chd["w"] = max(160, int(w0 + (x - x0) / self.zoom))
                chd["h"] = max(120, int(h0 + (y - y0) / self.zoom))
            self.mark_dirty()
            self.da.queue_draw()
        return True

    def _autoscroll(self, x, y):
        a = self.da.get_allocation()
        if y > a.height - 10:
            self.vadj.set_value(self.vadj.get_value() + DEF_ROWH * self.zoom)
        elif y < self.hh() + 4 and self.vadj.get_value() > 0:
            self.vadj.set_value(max(0, self.vadj.get_value() - DEF_ROWH * self.zoom))
        if x > a.width - 10:
            self.hadj.set_value(self.hadj.get_value() + DEF_COLW * self.zoom / 2)
        elif x < self.hw() + 4 and self.hadj.get_value() > 0:
            self.hadj.set_value(max(0, self.hadj.get_value() - DEF_COLW * self.zoom / 2))

    def _release(self, w, ev):
        d = self.drag
        self.drag = None
        if d is None:
            return False
        if d[0] == "fill" and d[2]:
            self.fill_series(d[1], d[2])
        elif d[0] == "resize":
            self._update_adj()
        elif d[0] == "ref":
            self.editor.grab_focus_without_selecting()
            self.editor.set_position(-1 if self._ref_pos is None else self._ref_end)
        return True

    def _scroll(self, w, ev):
        ok, dx, dy = ev.get_scroll_deltas()
        if not ok:
            dx, dy = {Gdk.ScrollDirection.UP: (0, -1), Gdk.ScrollDirection.DOWN: (0, 1), Gdk.ScrollDirection.LEFT: (-1, 0),
                      Gdk.ScrollDirection.RIGHT: (1, 0)}.get(ev.direction, (0, 0))
        if ev.state & Gdk.ModifierType.CONTROL_MASK:
            self.zoom_sc.set_value(self.zoom_sc.get_value() - 10 * (1 if dy > 0 else -1 if dy < 0 else 0))
            return True
        if ev.state & Gdk.ModifierType.SHIFT_MASK:
            dx, dy = dy, dx
        if dy:
            self.vadj.set_value(min(self.vadj.get_upper() - self.vadj.get_page_size(),
                                    max(0, self.vadj.get_value() + dy * 3 * DEF_ROWH * self.zoom)))
        if dx:
            self.hadj.set_value(min(self.hadj.get_upper() - self.hadj.get_page_size(),
                                    max(0, self.hadj.get_value() + dx * DEF_COLW * self.zoom)))
        return True

    # ================================================================ editing
    def ref_insertable(self):
        t = self.editor.get_text()
        if not t.startswith("="):
            return False
        if self.drag and self.drag[0] == "ref":
            return True
        pos = self.editor.get_position()
        before = t[:pos].rstrip()
        if getattr(self, "_ref_end", None) is not None and pos == self._ref_end and self._ref_pos is not None:
            return True
        return before.endswith(("=", "(", ",", "+", "-", "*", "/", "^", "&", "<", ">", ":", ";"))

    def _insert_ref(self, r1, c1, r2, c2):
        t = self.editor.get_text()
        start = self._ref_pos
        end = getattr(self, "_ref_end", None)
        if end is None or end < start or self.drag is None:
            end = start
        ref = cell_name(r1, c1) + ("" if (r1, c1) == (r2, c2) else ":" + cell_name(r2, c2))
        new = t[:start] + ref + t[end:]
        self.editor.set_text(new)
        self._ref_end = start + len(ref)
        self.editor.set_position(self._ref_end)
        self.da.queue_draw()

    def start_edit(self, text=None, keep=False):
        if self.chart_sel is not None:
            return
        r, c = self.act
        self.ensure_visible(r, c)
        self.editing = "edit" if keep else "enter"
        self._ref_pos = None
        self._ref_end = None
        f = self.sh.fmt(r, c)
        self.editor.override_font(self._font_desc(f))
        self.updating = True
        self.editor.set_text(self.sh.raw(r, c) if keep else (text or ""))
        self.updating = False
        self.fxentry.set_text(self.editor.get_text())
        self._place_editor()
        self.editor.show()
        self.editor.grab_focus_without_selecting()
        self.editor.set_position(-1)
        self.st_mode.set_text("Edit" if keep else "Enter")
        self.da.queue_draw()

    def _place_editor(self):
        if not self.editing:
            return
        r, c = self.act
        x, y, w, h = self.cell_rect(r, c)
        lay = self.editor.create_pango_layout(self.editor.get_text() + "  ")
        tw = lay.get_pixel_size()[0] + 10
        a = self.da.get_allocation()
        w = min(max(w, tw), a.width - x - 4)
        self.editor.set_margin_start(max(0, int(x) - 1))
        self.editor.set_margin_top(max(0, int(y) - 1))
        self.editor.set_size_request(int(w) + 2, int(h) + 2)

    def _ed_changed(self, e):
        if self.updating:
            return
        self.updating = True
        self.fxentry.set_text(e.get_text())
        self.updating = False
        if self.drag is None:
            self._ref_pos = None
            self._ref_end = None
        self._place_editor()
        self.da.queue_draw()

    def _ed_key(self, w, ev):
        k = ev.keyval
        shift = ev.state & Gdk.ModifierType.SHIFT_MASK
        ctrl = ev.state & Gdk.ModifierType.CONTROL_MASK
        if k in (Gdk.KEY_Return, Gdk.KEY_KP_Enter):
            if ev.state & Gdk.ModifierType.MOD1_MASK:
                pos = self.editor.get_position()
                self.editor.insert_text("\n", pos)
                self.editor.set_position(pos + 1)
                return True
            if ctrl:
                self.commit_edit(fill_selection=True)
                return True
            self.commit_edit()
            self.move(-1 if shift else 1, 0)
            return True
        if k in (Gdk.KEY_Tab, Gdk.KEY_ISO_Left_Tab):
            self.commit_edit()
            self.move(0, -1 if shift or k == Gdk.KEY_ISO_Left_Tab else 1)
            return True
        if k == Gdk.KEY_Escape:
            self.cancel_edit()
            return True
        if k == Gdk.KEY_F2:
            self.editing = "edit" if self.editing == "enter" else "enter"
            self.st_mode.set_text("Edit" if self.editing == "edit" else "Enter")
            return True
        if self.editing == "enter" and k in (Gdk.KEY_Up, Gdk.KEY_Down, Gdk.KEY_Left, Gdk.KEY_Right):
            dr, dc = {Gdk.KEY_Up: (-1, 0), Gdk.KEY_Down: (1, 0), Gdk.KEY_Left: (0, -1), Gdk.KEY_Right: (0, 1)}[k]
            if self.ref_insertable():
                # arrow keys point at cells while typing a formula
                if self._ref_pos is None:
                    self._ref_pos = self.editor.get_position()
                    self._kref = self.act
                r, c = self._kref
                r, c = max(0, r + dr), max(0, c + dc)
                self._kref = (r, c)
                self.drag = ("kref",)
                self._insert_ref(r, c, r, c)
                self.drag = None
                self._ref_pos_keep = True
                return True
            self.commit_edit()
            self.move(dr, dc)
            return True
        return False

    def commit_edit(self, fill_selection=False):
        if not self.editing:
            return
        text = self.editor.get_text()
        self.editing = None
        self.editor.hide()
        self.st_mode.set_text("Ready")
        r, c = self.act
        if text.startswith("=") and len(text) > 1:
            text = self._autoclose(text)
        if fill_selection:
            r1, c1, r2, c2 = self.clamped_sel()
            self.push_undo()
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    self.set_raw(rr, cc, shift_formula(text, rr - r, cc - c) if text.startswith("=") else text)
            self.changed()
        elif text != self.sh.raw(r, c):
            self.push_undo()
            self.set_raw(r, c, text)
            self.changed()
        self.da.grab_focus()
        self._sel_changed()

    @staticmethod
    def _autoclose(text):
        depth = 0
        instr = False
        for ch in text:
            if ch == '"':
                instr = not instr
            elif not instr:
                depth += (ch == "(") - (ch == ")")
        if instr:
            text += '"'
        return text + ")" * max(0, depth)

    def cancel_edit(self):
        self.editing = None
        self.editor.hide()
        self.st_mode.set_text("Ready")
        self.da.grab_focus()
        self._sel_changed()

    def _fx_changed(self, e):
        if self.updating:
            return
        if self.editing:
            self.updating = True
            self.editor.set_text(e.get_text())
            self.updating = False
            self._place_editor()
        elif e.has_focus():
            self.editing = "edit"
            self._ref_pos = None
            r, c = self.act
            self.editor.override_font(self._font_desc(self.sh.fmt(r, c)))
            self.updating = True
            self.editor.set_text(e.get_text())
            self.updating = False
            self._place_editor()
            self.editor.show()
            self.st_mode.set_text("Edit")
        self.da.queue_draw()

    def _fx_commit(self):
        if self.editing:
            self.updating = True
            self.editor.set_text(self.fxentry.get_text())
            self.updating = False
            self.commit_edit()
        self.move(1, 0)

    def _fx_key(self, w, ev):
        if ev.keyval == Gdk.KEY_Escape:
            self.cancel_edit()
            return True
        if ev.keyval == Gdk.KEY_Tab:
            self._fx_commit()
            return True
        return False

    # ================================================================ cell changes
    def set_raw(self, r, c, raw, fmt=None, sheet=None):
        sh = sheet or self.sh
        cell = sh.cells.get((r, c))
        f = dict(cell["f"]) if cell else {}
        if fmt is not None:
            f = dict(fmt)
        if raw and not raw.startswith(("=", "'")) and f.get("num") in (None, "General") and fmt is None:
            p = parse_number(raw)
            if p is not None and p[1]:
                f["num"] = p[1]
            elif p is None:
                d = parse_date(raw)
                if d is not None:
                    f["num"] = d[1]
        if raw == "" and not f:
            sh.cells.pop((r, c), None)
        else:
            sh.cells[(r, c)] = {"raw": raw, "f": f}
        sh.touch()

    def set_fmt(self, key, value, toggle=False):
        r1, c1, r2, c2 = self.fmt_range()
        self.push_undo()
        on = value
        if toggle:
            on = not all(self.sh.fmt(r, c).get(key) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1))
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.sh.cells.get((r, c))
                if cell is None:
                    if not on:
                        continue
                    cell = self.sh.cells[(r, c)] = {"raw": "", "f": {}}
                if on in (None, False):
                    cell["f"].pop(key, None)
                else:
                    cell["f"][key] = on
                if not cell["raw"] and not cell["f"]:
                    del self.sh.cells[(r, c)]
        self.sh.touch()
        if key in ("wrap", "size", "font") :
            for r in range(r1, r2 + 1):
                if r not in self.sh.rowh or key == "wrap":
                    self.autofit("r", r, quiet=True)
        self.changed(recalc=False)

    def fmt_range(self):
        r1, c1, r2, c2 = self.sel
        mr, mc = self.sh.extent()
        if r2 >= 10 ** 6:
            r2 = max(mr, r1)
        if c2 >= 10 ** 4:
            c2 = max(mc, c1)
        return r1, c1, r2, c2

    def changed(self, recalc=True):
        if recalc:
            self.book.recalc()
        self.mark_dirty()
        self._update_adj()
        self._sel_changed()

    def mark_dirty(self):
        if not self.dirty:
            self.dirty = True
            self.update_title()

    # ================================================================ undo
    def snapshot(self):
        return {"sheets": [(s.name, {k: {"raw": v["raw"], "f": dict(v["f"])} for k, v in s.cells.items()}, dict(s.colw),
                            dict(s.rowh), [dict(ch) for ch in s.charts], s.freeze) for s in self.book.sheets],
                "si": self.si, "sel": self.sel, "act": self.act}

    def restore(self, snap):
        sheets = []
        for name, cells, colw, rowh, charts, freeze in snap["sheets"]:
            s = Sheet(name)
            s.cells = {k: {"raw": v["raw"], "f": dict(v["f"])} for k, v in cells.items()}
            s.colw, s.rowh, s.charts, s.freeze = dict(colw), dict(rowh), [dict(c) for c in charts], freeze
            sheets.append(s)
        self.book.sheets = sheets
        self.si = min(snap["si"], len(sheets) - 1)
        self.sel, self.act = snap["sel"], snap["act"]
        self.anchor = self.act
        self.chart_sel = None
        self.book.recalc()
        self.build_sheet_tabs()
        self.mark_dirty()
        self._update_adj()
        self._sel_changed()

    def push_undo(self):
        self.undo_stack.append(self.snapshot())
        del self.undo_stack[:-100]
        self.redo_stack.clear()

    def undo(self):
        if self.editing:
            self.cancel_edit()
            return
        if self.undo_stack:
            self.redo_stack.append(self.snapshot())
            self.restore(self.undo_stack.pop())

    def redo(self):
        if self.redo_stack:
            self.undo_stack.append(self.snapshot())
            self.restore(self.redo_stack.pop())

    # ================================================================ keyboard
    def _win_key(self, w, ev):
        ctrl = ev.state & Gdk.ModifierType.CONTROL_MASK
        shift = ev.state & Gdk.ModifierType.SHIFT_MASK
        k = Gdk.keyval_to_lower(ev.keyval)
        if ev.keyval == Gdk.KEY_Escape and self.findbar.get_visible():
            self.findbar.hide()
            self.da.grab_focus()
            return True
        if not ctrl:
            return False
        acts = {Gdk.KEY_s: (self.save_as if shift else self.save), Gdk.KEY_o: self.open_dialog, Gdk.KEY_n: self.app.new_window,
                Gdk.KEY_w: self.close, Gdk.KEY_f: lambda: self.show_find(False), Gdk.KEY_h: lambda: self.show_find(True),
                Gdk.KEY_z: self.undo, Gdk.KEY_y: self.redo, Gdk.KEY_equal: lambda: self.zoom_sc.set_value(self.zoom_sc.get_value() + 10),
                Gdk.KEY_minus: lambda: self.zoom_sc.set_value(self.zoom_sc.get_value() - 10)}
        if self.editing or self.fxentry.has_focus() or self.namebox.has_focus() or self.f_entry.has_focus() or self.r_entry.has_focus():
            if k in (Gdk.KEY_s, Gdk.KEY_o, Gdk.KEY_n, Gdk.KEY_w):
                if self.editing:
                    self.commit_edit()
                acts[k]()
                return True
            return False
        if k in acts:
            acts[k]()
            return True
        return False

    def _grid_key(self, w, ev):
        k = ev.keyval
        ctrl = ev.state & Gdk.ModifierType.CONTROL_MASK
        shift = ev.state & Gdk.ModifierType.SHIFT_MASK
        arrows = {Gdk.KEY_Up: (-1, 0), Gdk.KEY_Down: (1, 0), Gdk.KEY_Left: (0, -1), Gdk.KEY_Right: (0, 1)}
        if self.chart_sel is not None:
            if k in (Gdk.KEY_Delete, Gdk.KEY_BackSpace):
                self.push_undo()
                del self.sh.charts[self.chart_sel]
                self.chart_sel = None
                self.changed(recalc=False)
                return True
            if k == Gdk.KEY_Escape:
                self.chart_sel = None
                self.da.queue_draw()
                return True
        if k in arrows:
            dr, dc = arrows[k]
            self.move(dr, dc, extend=bool(shift), jump=bool(ctrl))
            return True
        if k in (Gdk.KEY_Return, Gdk.KEY_KP_Enter, Gdk.KEY_Tab, Gdk.KEY_ISO_Left_Tab):
            down = k in (Gdk.KEY_Return, Gdk.KEY_KP_Enter)
            back = shift or k == Gdk.KEY_ISO_Left_Tab
            r1, c1, r2, c2 = self.sel
            if (r1, c1) != (r2, c2):
                # move inside the selection
                r, c = self.act
                if down:
                    r += -1 if back else 1
                    if r > r2:
                        r, c = r1, c + 1 if c < c2 else c1
                    elif r < r1:
                        r, c = r2, c - 1 if c > c1 else c2
                else:
                    c += -1 if back else 1
                    if c > c2:
                        c, r = c1, r + 1 if r < r2 else r1
                    elif c < c1:
                        c, r = c2, r - 1 if r > r1 else r2
                self.act = (r, c)
                self._sel_changed()
                return True
            self.move(*((-1 if back else 1, 0) if down else (0, -1 if back else 1)))
            return True
        if k == Gdk.KEY_Home:
            self.select(0 if ctrl else self.act[0], 0)
            self.ensure_visible(*self.act)
            return True
        if k == Gdk.KEY_End and ctrl:
            mr, mc = self.sh.extent()
            self.select(max(0, mr), max(0, mc))
            self.ensure_visible(*self.act)
            return True
        if k in (Gdk.KEY_Page_Down, Gdk.KEY_Page_Up):
            n = max(1, int(self.vadj.get_page_size() / (DEF_ROWH * self.zoom)))
            self.move(n if k == Gdk.KEY_Page_Down else -n, 0, extend=bool(shift))
            return True
        if k == Gdk.KEY_Delete:
            self.clear_contents()
            return True
        if k == Gdk.KEY_BackSpace:
            self.push_undo()
            self.set_raw(*self.act, "")
            self.changed()
            self.start_edit("")
            return True
        if k == Gdk.KEY_F2:
            self.start_edit(keep=True)
            return True
        if ctrl:
            low = Gdk.keyval_to_lower(k)
            acts = {Gdk.KEY_c: self.copy, Gdk.KEY_x: lambda: self.copy(cut=True), Gdk.KEY_v: self.paste,
                    Gdk.KEY_b: lambda: self.set_fmt("b", True, toggle=True), Gdk.KEY_i: lambda: self.set_fmt("i", True, toggle=True),
                    Gdk.KEY_u: lambda: self.set_fmt("u", True, toggle=True), Gdk.KEY_d: self.fill_down,
                    Gdk.KEY_r: self.fill_right, Gdk.KEY_a: lambda: self.select(0, 0, 10 ** 7, 10 ** 5, act=self.act),
                    Gdk.KEY_semicolon: lambda: self._type_now("date"), Gdk.KEY_colon: lambda: self._type_now("time"),
                    Gdk.KEY_space: lambda: self.select(0, self.act[1], 10 ** 7, self.act[1], act=self.act),
                    Gdk.KEY_grave: self.toggle_formulas, Gdk.KEY_Home: lambda: self.select(0, 0)}
            if low in acts:
                acts[low]()
                return True
            return False
        if shift and k == Gdk.KEY_space:
            self.select(self.act[0], 0, self.act[0], 10 ** 5, act=self.act)
            return True
        if ev.state & Gdk.ModifierType.MOD1_MASK and k == Gdk.KEY_equal:
            self.autosum()
            return True
        ch = Gdk.keyval_to_unicode(k)
        if ch and ch >= 32 and not (ev.state & Gdk.ModifierType.MOD1_MASK):
            self.start_edit(chr(ch))
            return True
        return False

    def _type_now(self, kind):
        now = datetime.datetime.now()
        r, c = self.act
        self.push_undo()
        if kind == "date":
            self.set_raw(r, c, general(float(to_serial(now.date()))), fmt=dict(self.sh.fmt(r, c), num="m/d/yyyy"))
        else:
            self.set_raw(r, c, general(to_serial(now.time())), fmt=dict(self.sh.fmt(r, c), num="h:mm AM/PM"))
        self.changed()

    def toggle_formulas(self):
        self.show_formulas = not self.show_formulas
        self.da.queue_draw()
        self._update_ribbon_state()

    # ================================================================ clipboard
    def copy(self, cut=False):
        r1, c1, r2, c2 = self.clamped_sel()
        cells = {}
        lines = []
        for r in range(r1, r2 + 1):
            row = []
            for c in range(c1, c2 + 1):
                cl = self.sh.cells.get((r, c))
                if cl:
                    cells[(r - r1, c - c1)] = {"raw": cl["raw"], "f": dict(cl["f"])}
                row.append(display_text(self.book, self.sh, r, c))
            lines.append("\t".join(row))
        text = "\n".join(lines)
        self.clip = {"sheet": self.sh, "box": (r1, c1, r2, c2), "cells": cells, "text": text, "cut": cut,
                     "size": (r2 - r1 + 1, c2 - c1 + 1)}
        Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD).set_text(text, -1)
        self.da.queue_draw()

    def paste(self, values_only=False):
        text = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD).wait_for_text()
        r0, c0 = self.sel[0], self.sel[1]
        self.push_undo()
        if self.clip and text == self.clip["text"]:
            cl = self.clip
            nr, nc = cl["size"]
            r1, c1, r2, c2 = self.clamped_sel()
            # one copied block pasted over a bigger selection repeats to fill it
            reps_r = max(1, (r2 - r1 + 1) // nr) if (r2 - r1 + 1) % nr == 0 else 1
            reps_c = max(1, (c2 - c1 + 1) // nc) if (c2 - c1 + 1) % nc == 0 else 1
            sr, sc = cl["box"][0], cl["box"][1]
            if cl["cut"]:
                for (i, j) in list(cl["cells"]):
                    cl["sheet"].cells.pop((sr + i, sc + j), None)
                cl["sheet"].touch()
            for a in range(reps_r):
                for b in range(reps_c):
                    tr, tc = r0 + a * nr, c0 + b * nc
                    for i in range(nr):
                        for j in range(nc):
                            src = cl["cells"].get((i, j))
                            if src is None:
                                self.sh.cells.pop((tr + i, tc + j), None)
                                continue
                            raw = src["raw"]
                            if values_only:
                                v = self.book.value(cl["sheet"], sr + i, sc + j) if not cl["cut"] else None
                                raw = general(v) if isinstance(v, float) else to_str_safe(v) if v is not None else raw
                                self.set_raw(tr + i, tc + j, raw)
                                continue
                            if raw.startswith("=") and not cl["cut"]:
                                raw = shift_formula(raw, tr + i - (sr + i), tc + j - (sc + j))
                            self.set_raw(tr + i, tc + j, raw, fmt=src["f"])
            self.select(r0, c0, r0 + nr * reps_r - 1, c0 + nc * reps_c - 1)
            if cl["cut"]:
                self.clip = None
        elif text:
            rows = [ln.split("\t") for ln in text.rstrip("\n").split("\n")]
            for i, row in enumerate(rows):
                for j, v in enumerate(row):
                    self.set_raw(r0 + i, c0 + j, v.rstrip("\r"))
            self.select(r0, c0, r0 + len(rows) - 1, c0 + max(len(r) for r in rows) - 1)
        self.changed()

    def clear_contents(self):
        r1, c1, r2, c2 = self.fmt_range()
        self.push_undo()
        for (r, c) in list(self.sh.cells):
            if r1 <= r <= r2 and c1 <= c <= c2:
                cl = self.sh.cells[(r, c)]
                if cl["f"]:
                    cl["raw"] = ""
                    cl.pop("p", None)
                else:
                    del self.sh.cells[(r, c)]
        self.sh.touch()
        self.changed()

    def clear_all(self):
        r1, c1, r2, c2 = self.fmt_range()
        self.push_undo()
        for (r, c) in list(self.sh.cells):
            if r1 <= r <= r2 and c1 <= c <= c2:
                del self.sh.cells[(r, c)]
        self.sh.touch()
        self.changed()

    # ================================================================ fill
    def fill_down(self):
        r1, c1, r2, c2 = self.clamped_sel()
        if r1 == r2:
            if r1 == 0:
                return
            r1 -= 1
        self.fill_series((r1, c1, r1, c2), (r1, c1, r2, c2), copy_only=True)

    def fill_right(self):
        r1, c1, r2, c2 = self.clamped_sel()
        if c1 == c2:
            if c1 == 0:
                return
            c1 -= 1
        self.fill_series((r1, c1, r2, c1), (r1, c1, r2, c2), copy_only=True)

    def fill_series(self, src, dst, copy_only=False):
        sr1, sc1, sr2, sc2 = src
        dr1, dc1, dr2, dc2 = dst
        mr, mc = self.sh.extent()
        sr2, sc2 = min(sr2, max(mr, sr1)), min(sc2, max(mc, sc1))
        self.push_undo()
        vertical = dc1 == sc1 and dc2 == sc2
        if vertical:
            lines = [[(r, c) for r in range(sr1, sr2 + 1)] for c in range(sc1, sc2 + 1)]
            targets = [[(r, c) for r in (range(sr2 + 1, dr2 + 1) if dr2 > sr2 else range(sr1 - 1, dr1 - 1, -1))]
                       for c in range(sc1, sc2 + 1)]
        else:
            lines = [[(r, c) for c in range(sc1, sc2 + 1)] for r in range(sr1, sr2 + 1)]
            targets = [[(r, c) for c in (range(sc2 + 1, dc2 + 1) if dc2 > sc2 else range(sc1 - 1, dc1 - 1, -1))]
                       for r in range(sr1, sr2 + 1)]
        for src_cells, tgt in zip(lines, targets):
            backwards = tgt and (tgt[0][0] < src_cells[0][0] or tgt[0][1] < src_cells[0][1])
            seq = list(reversed(src_cells)) if backwards else src_cells
            raws = [self.sh.raw(*p) for p in seq]
            fmts = [self.sh.fmt(*p) for p in seq]
            nums = [parse_input(x)[1] if x and not x.startswith("=") else None for x in raws]
            numeric = all(isinstance(v, float) for v in nums) and raws and all(raws)
            step = None
            if not copy_only and numeric:
                if len(nums) >= 2:
                    step = (nums[-1] - nums[0]) / (len(nums) - 1)
                elif is_date_fmt(fmts[0].get("num") or "General") and fmts[0].get("num"):
                    step = 1.0
                if step is not None and backwards:
                    pass
            for k, (r, c) in enumerate(tgt):
                i = k % len(seq)
                raw, f = raws[i], fmts[i]
                if step is not None:
                    v = nums[-1] + step * (k + 1)
                    raw = general(v)
                elif raw.startswith("="):
                    sr_, sc_ = seq[i]
                    raw = shift_formula(raw, r - sr_, c - sc_)
                elif not copy_only and len(seq) == 1 and raw and not numeric:
                    m = re.match(r"^(.*?)(\d+)$", raw)
                    if m:
                        n = int(m.group(2)) + (k + 1) * (-1 if backwards else 1)
                        raw = m.group(1) + str(max(0, n)).zfill(len(m.group(2)))
                self.set_raw(r, c, raw, fmt=f)
        self.select(dr1, dc1, dr2, dc2, act=self.act)
        self.changed()

    # ================================================================ rows / columns
    def insert_rc(self, axis, count=None):
        r1, c1, r2, c2 = self.sel
        at = r1 if axis == 0 else c1
        n = count or ((r2 - r1 + 1) if axis == 0 and r2 < 10 ** 6 else (c2 - c1 + 1) if axis == 1 and c2 < 10 ** 4 else 1)
        self.push_undo()
        self._shift_struct(axis, at, n)
        self.changed()

    def delete_rc(self, axis):
        r1, c1, r2, c2 = self.sel
        at = r1 if axis == 0 else c1
        n = (r2 - r1 + 1) if axis == 0 else (c2 - c1 + 1)
        mr, mc = self.sh.extent()
        n = min(n, (mr if axis == 0 else mc) - at + 1) if n > 1000 else n
        if n <= 0:
            return
        self.push_undo()
        self._shift_struct(axis, at, -n)
        self.changed()

    def _shift_struct(self, axis, at, n):
        sh = self.sh
        new = {}
        for (r, c), cell in sh.cells.items():
            v = r if axis == 0 else c
            if n < 0 and at <= v < at - n:
                continue
            if v >= at:
                v += n
            new[(v, c) if axis == 0 else (r, v)] = cell
        sh.cells = new
        dims = sh.rowh if axis == 0 else sh.colw
        nd = {}
        for k, val in dims.items():
            if n < 0 and at <= k < at - n:
                continue
            nd[k + n if k >= at else k] = val
        dims.clear()
        dims.update(nd)
        for s in self.book.sheets:
            for cell in s.cells.values():
                if cell["raw"].startswith("="):
                    nr = struct_shift(cell["raw"], s.name, sh.name, axis, at, n)
                    if nr != cell["raw"]:
                        cell["raw"] = nr
                        cell.pop("p", None)
            for chd in s.charts:
                if (chd.get("sheet") or s.name).lower() == sh.name.lower():
                    r1, c1, r2, c2 = chd["range"]
                    if axis == 0:
                        r1, r2 = (r1 + n if r1 >= at else r1), (r2 + n if r2 >= at else r2)
                    else:
                        c1, c2 = (c1 + n if c1 >= at else c1), (c2 + n if c2 >= at else c2)
                    chd["range"] = (max(0, r1), max(0, c1), max(0, r2), max(0, c2))
        for chd in sh.charts:
            if axis == 0 and chd["row"] >= at:
                chd["row"] = max(0, chd["row"] + n)
            if axis == 1 and chd["col"] >= at:
                chd["col"] = max(0, chd["col"] + n)
        sh.touch()

    def autofit(self, kind, idx, quiet=False):
        if not quiet:
            self.push_undo()
        surf = cairo.ImageSurface(cairo.FORMAT_ARGB32, 1, 1)
        cr = cairo.Context(surf)
        z = self.zoom
        self.zoom = 1.0
        best = 0
        for (r, c), cell in self.sh.cells.items():
            if (kind == "c" and c == idx) or (kind == "r" and r == idx):
                text, v = self.cell_text(r, c)
                if not text:
                    continue
                lay = self._layout(cr, text, cell["f"])
                if kind == "r" and cell["f"].get("wrap"):
                    lay.set_width(int((self.sh.colw.get(c, DEF_COLW) - 6) * Pango.SCALE))
                    lay.set_wrap(Pango.WrapMode.WORD_CHAR)
                tw, th = lay.get_pixel_size()
                best = max(best, tw + 10 if kind == "c" else th + 6)
        self.zoom = z
        if kind == "c":
            if best:
                self.sh.colw[idx] = max(20, best)
            else:
                self.sh.colw.pop(idx, None)
        else:
            if best > DEF_ROWH:
                self.sh.rowh[idx] = best
            else:
                self.sh.rowh.pop(idx, None)
        if not quiet:
            self.changed(recalc=False)

    # ================================================================ sort / autosum
    def current_region(self, r, c):
        has = lambda rr, cc: self.sh.raw(rr, cc) != ""  # noqa: E731
        r1 = r2 = r
        c1 = c2 = c
        grew = True
        while grew:
            grew = False
            if r1 > 0 and any(has(r1 - 1, cc) for cc in range(max(0, c1 - 1), c2 + 2)):
                r1 -= 1
                grew = True
            if any(has(r2 + 1, cc) for cc in range(max(0, c1 - 1), c2 + 2)):
                r2 += 1
                grew = True
            if c1 > 0 and any(has(rr, c1 - 1) for rr in range(max(0, r1 - 1), r2 + 2)):
                c1 -= 1
                grew = True
            if any(has(rr, c2 + 1) for rr in range(max(0, r1 - 1), r2 + 2)):
                c2 += 1
                grew = True
        return r1, c1, r2, c2

    def sort(self, desc=False):
        r1, c1, r2, c2 = self.clamped_sel()
        if (r1, c1) == (r2, c2):
            r1, c1, r2, c2 = self.current_region(r1, c1)
            # a text first row over numbers is a header
            first = [self.book.value(self.sh, r1, c) for c in range(c1, c2 + 1)]
            second = [self.book.value(self.sh, r1 + 1, c) for c in range(c1, c2 + 1)] if r2 > r1 else []
            if r2 > r1 and all(isinstance(v, str) and v for v in first if v is not None) and \
                    any(not isinstance(v, str) for v in second if v is not None):
                r1 += 1
            elif r2 > r1 and all(isinstance(v, str) for v in first) and any(self.sh.fmt(r1, c).get("b") for c in range(c1, c2 + 1)):
                r1 += 1
            # leave total rows (formulas over a range) at the bottom alone
            while r2 > r1 and any(self.sh.raw(r2, c).startswith("=") and ":" in self.sh.raw(r2, c) for c in range(c1, c2 + 1)):
                r2 -= 1
        if r2 <= r1:
            return
        key_c = min(max(self.act[1], c1), c2)
        rows = list(range(r1, r2 + 1))

        def keyf(r):
            v = self.book.value(self.sh, r, key_c)
            if v is None or v == "":
                return (1, 0, 0)
            return (0, type_rank(v), v.lower() if isinstance(v, str) else (v.code if isinstance(v, XlErr) else v))
        blanks = [r for r in rows if keyf(r)[0] == 1]
        filled = sorted([r for r in rows if keyf(r)[0] == 0], key=keyf, reverse=desc)
        order = filled + blanks
        self.push_undo()
        old = {(r, c): self.sh.cells.get((r, c)) for r in rows for c in range(c1, c2 + 1)}
        for i, src_r in enumerate(order):
            dst_r = r1 + i
            for c in range(c1, c2 + 1):
                cell = old[(src_r, c)]
                if cell is None:
                    self.sh.cells.pop((dst_r, c), None)
                else:
                    raw = shift_formula(cell["raw"], dst_r - src_r, 0) if cell["raw"].startswith("=") else cell["raw"]
                    self.sh.cells[(dst_r, c)] = {"raw": raw, "f": dict(cell["f"])}
        self.sh.touch()
        self.select(r1, c1, r2, c2, act=(r1, key_c))
        self.changed()

    def autosum(self, func="SUM"):
        r1, c1, r2, c2 = self.clamped_sel()
        isnum = lambda r, c: isinstance(self.book.value(self.sh, r, c), (int, float)) and \
            not isinstance(self.book.value(self.sh, r, c), bool)  # noqa: E731
        self.push_undo()
        if (r1, c1) != (r2, c2):
            # totals under each selected column
            for c in range(c1, c2 + 1):
                self.set_raw(r2 + 1, c, f"={func}({cell_name(r1, c)}:{cell_name(r2, c)})")
            self.changed()
            return
        r, c = r1, c1
        rr = r - 1
        while rr >= 0 and isnum(rr, c):
            rr -= 1
        if rr < r - 1:
            ref = f"{cell_name(rr + 1, c)}:{cell_name(r - 1, c)}"
        else:
            cc = c - 1
            while cc >= 0 and isnum(r, cc):
                cc -= 1
            ref = f"{cell_name(r, cc + 1)}:{cell_name(r, c - 1)}" if cc < c - 1 else ""
        self.undo_stack.pop()
        self.start_edit(f"={func}({ref})")
        if ref:
            self.editor.select_region(len(func) + 2, len(func) + 2 + len(ref))

    # ================================================================ sheets
    def build_sheet_tabs(self):
        for ch in self.sheet_tabs.get_children():
            ch.destroy()
        for i, s in enumerate(self.book.sheets):
            b = Gtk.Button(label=s.name)
            add_cls(b, "stab")
            if i == self.si:
                add_cls(b, "sel")
            b.connect("button-press-event", self._tab_press, i)
            self.sheet_tabs.pack_start(b, False, False, 0)
        self.sheet_tabs.show_all()

    def _tab_press(self, w, ev, i):
        if ev.type == Gdk.EventType._2BUTTON_PRESS:
            self.rename_sheet(i)
            return True
        if ev.button == 3:
            self.switch_sheet(i)
            items = [{"label": "Insert sheet", "icon": "plus", "cb": self.add_sheet},
                     {"label": "Duplicate", "icon": "copy", "cb": lambda: self.dup_sheet(i)},
                     {"label": "Rename", "icon": "rename", "cb": lambda: self.rename_sheet(i)},
                     {"label": "Move left", "cb": lambda: self.move_sheet(i, -1), "sensitive": i > 0},
                     {"label": "Move right", "cb": lambda: self.move_sheet(i, 1), "sensitive": i < len(self.book.sheets) - 1},
                     None, {"label": "Delete", "icon": "delete", "cb": lambda: self.del_sheet(i),
                            "sensitive": len(self.book.sheets) > 1}]
            popup_menu(build_menu(items), w, ev)
            return True
        if ev.button == 1:
            self.switch_sheet(i)
        return False

    def switch_sheet(self, i):
        if self.editing:
            self.commit_edit()
        self.si = i
        self.hadj.set_value(0)
        self.vadj.set_value(0)
        self.vrows, self.vcols = 200, 40
        self.select(0, 0)
        self.build_sheet_tabs()
        self._update_adj()
        self.da.queue_draw()

    def add_sheet(self):
        n = 1
        names = {s.name.lower() for s in self.book.sheets}
        while f"sheet{n}" in names:
            n += 1
        self.push_undo()
        self.book.sheets.insert(self.si + 1, Sheet(f"Sheet{n}"))
        self.mark_dirty()
        self.switch_sheet(self.si + 1)

    def dup_sheet(self, i):
        self.push_undo()
        src = self.book.sheets[i]
        s = Sheet(self._unique(src.name + " (2)"))
        s.cells = {k: {"raw": v["raw"], "f": dict(v["f"])} for k, v in src.cells.items()}
        s.colw, s.rowh, s.freeze = dict(src.colw), dict(src.rowh), src.freeze
        s.charts = [dict(c) for c in src.charts]
        self.book.sheets.insert(i + 1, s)
        self.book.recalc()
        self.mark_dirty()
        self.switch_sheet(i + 1)

    def _unique(self, name):
        names = {s.name.lower() for s in self.book.sheets}
        base, n = name, 2
        while name.lower() in names:
            name = f"{base} ({n})"
            n += 1
        return name[:31]

    def move_sheet(self, i, d):
        j = i + d
        if 0 <= j < len(self.book.sheets):
            self.push_undo()
            s = self.book.sheets
            s[i], s[j] = s[j], s[i]
            self.mark_dirty()
            self.switch_sheet(j)

    def del_sheet(self, i):
        if len(self.book.sheets) < 2:
            return

        def go(ok):
            if not ok:
                return
            self.push_undo()
            del self.book.sheets[i]
            self.book.recalc()
            self.mark_dirty()
            self.switch_sheet(max(0, min(i, len(self.book.sheets) - 1)))
        confirm(self, "Delete sheet", f"Delete \"{self.book.sheets[i].name}\"? This can be undone with Ctrl+Z.", "Delete", go)

    def rename_sheet(self, i):
        e = Gtk.Entry()
        e.set_text(self.book.sheets[i].name)
        e.set_max_length(31)
        dlg = WinDialog(self, "Rename sheet", primary="Rename", secondary="Cancel", width=340, extra=e)
        e.connect("activate", lambda w: dlg.done(True))

        def go(ok):
            new = e.get_text().strip()
            if not ok or not new or re.search(r"[\\/?*\[\]:]", new):
                return
            old = self.book.sheets[i].name
            if new.lower() != old.lower() and self.book.sheet_named(new):
                info_dialog(self, "Rename sheet", "That name is already used by another sheet.")
                return
            self.push_undo()
            for s in self.book.sheets:
                for cell in s.cells.values():
                    if cell["raw"].startswith("="):
                        nr = rename_sheet_refs(cell["raw"], old, new)
                        if nr != cell["raw"]:
                            cell["raw"] = nr
                            cell.pop("p", None)
                for chd in s.charts:
                    if (chd.get("sheet") or "").lower() == old.lower():
                        chd["sheet"] = new
            self.book.sheets[i].name = new
            self.book.recalc()
            self.mark_dirty()
            self.build_sheet_tabs()
            self.da.queue_draw()
        dlg.ask(go)

    # ================================================================ charts
    def insert_chart(self, typ):
        r1, c1, r2, c2 = self.clamped_sel()
        if (r1, c1) == (r2, c2):
            if self.sh.raw(r1, c1) == "":
                info_dialog(self, "Insert chart", "Select the data for the chart first (for example a column of names "
                                                  "next to a column of numbers), then pick a chart.")
                return
            r1, c1, r2, c2 = self.current_region(r1, c1)
        self.push_undo()
        title = None
        src, head, cats_col, dr1, sc1 = chart_layout(self.book, self.sh, {"range": (r1, c1, r2, c2)})
        if head and c2 == sc1:
            title = to_str_safe(self.book.value(self.sh, r1, sc1))
        chd = {"type": typ, "range": (r1, c1, r2, c2), "title": title, "row": r1 + 1, "col": c2 + 2, "w": 480, "h": 288,
               "sheet": self.sh.name}
        self.sh.charts.append(chd)
        self.chart_sel = len(self.sh.charts) - 1
        self.ensure_visible(r1, min(c2 + 6, c2 + 2))
        self.changed(recalc=False)
        self.chart_sel = len(self.sh.charts) - 1
        self.da.queue_draw()

    def chart_menu(self, ev):
        i = self.chart_sel
        chd = self.sh.charts[i]

        def set_type(t):
            self.push_undo()
            chd["type"] = t
            self.changed(recalc=False)
        items = [{"label": n, "icon": ic, "checked": chd["type"] == t, "cb": lambda t=t: set_type(t)}
                 for t, n, ic in (("col", "Column", "ch-col"), ("bar", "Bar", "ch-bar"), ("line", "Line", "ch-line"),
                                  ("pie", "Pie", "ch-pie"))]
        items += [None, {"label": "Edit title...", "icon": "rename", "cb": lambda: self.edit_chart_title(i)},
                  {"label": "Delete chart", "icon": "delete", "cb": lambda: self._del_chart(i)}]
        popup_menu(build_menu(items), self.da, ev)

    def _del_chart(self, i):
        self.push_undo()
        del self.sh.charts[i]
        self.chart_sel = None
        self.changed(recalc=False)

    def edit_chart_title(self, i):
        chd = self.sh.charts[i]
        e = Gtk.Entry()
        e.set_text(chd.get("title") or "")
        dlg = WinDialog(self, "Chart title", primary="OK", secondary="Cancel", width=360, extra=e)
        e.connect("activate", lambda w: dlg.done(True))

        def go(ok):
            if ok:
                self.push_undo()
                chd["title"] = e.get_text().strip() or None
                self.changed(recalc=False)
        dlg.ask(go)

    # ================================================================ context menu
    def cell_menu(self, ev):
        items = [{"label": "Cut", "icon": "cut", "accel": "Ctrl+X", "cb": lambda: self.copy(cut=True)},
                 {"label": "Copy", "icon": "copy", "accel": "Ctrl+C", "cb": self.copy},
                 {"label": "Paste", "icon": "paste", "accel": "Ctrl+V", "cb": self.paste},
                 {"label": "Paste values", "cb": lambda: self.paste(values_only=True)}, None,
                 {"label": "Insert rows above", "icon": "ins-row", "cb": lambda: self.insert_rc(0)},
                 {"label": "Insert columns left", "icon": "ins-col", "cb": lambda: self.insert_rc(1)},
                 {"label": "Delete rows", "icon": "del-row", "cb": lambda: self.delete_rc(0)},
                 {"label": "Delete columns", "icon": "del-col", "cb": lambda: self.delete_rc(1)}, None,
                 {"label": "Clear contents", "accel": "Delete", "cb": self.clear_contents},
                 {"label": "Clear all", "icon": "delete", "cb": self.clear_all}, None,
                 {"label": "Sort A to Z", "icon": "sort-az", "cb": lambda: self.sort(False)},
                 {"label": "Sort Z to A", "icon": "sort-za", "cb": lambda: self.sort(True)}]
        popup_menu(build_menu(items), self.da, ev)

    # ================================================================ ribbon
    def show_tab(self, key):
        for k, b in self.tab_btns.items():
            (b.get_style_context().add_class if k == key else b.get_style_context().remove_class)("sel")
        self.ribbon.set_visible_child_name(key)

    def rsep(self, box):
        s = Gtk.Box()
        add_cls(s, "rsep")
        box.pack_start(s, False, False, 0)

    def rbtn(self, box, icon, tip, cb, text=None, markup=None):
        b = Gtk.Button()
        add_cls(b, "rb")
        h = Gtk.Box(spacing=6)
        if icon:
            h.pack_start(Icon(icon, 16), False, False, 0)
        if markup:
            lb = Gtk.Label()
            lb.set_markup(markup)
            h.pack_start(lb, False, False, 0)
        if text:
            h.pack_start(Gtk.Label(label=text), False, False, 0)
        b.add(h)
        b.set_tooltip_text(tip)
        if cb:
            b.connect("clicked", lambda w: (cb(), self.da.grab_focus() if not self.editing else None))
        box.pack_start(b, False, False, 0)
        return b

    def menu_btn(self, box, icon, tip, items_fn, text=None):
        b = self.rbtn(box, icon, tip, None, text=text)
        b.connect("clicked", lambda w: popup_menu(build_menu(items_fn()), w))
        return b

    def rib_home(self, box):
        self.rbtn(box, "paste", "Paste (Ctrl+V)", self.paste)
        self.rbtn(box, "cut", "Cut (Ctrl+X)", lambda: self.copy(cut=True))
        self.rbtn(box, "copy", "Copy (Ctrl+C)", self.copy)
        self.rsep(box)
        self.font_cb = Gtk.ComboBoxText.new_with_entry()
        add_cls(self.font_cb, "fontbox")
        for f in FONTS:
            self.font_cb.append_text(f)
        self.font_cb.get_child().set_width_chars(12)
        self.font_cb.connect("changed", self._font_changed)
        box.pack_start(self.font_cb, False, False, 0)
        self.size_cb = Gtk.ComboBoxText.new_with_entry()
        add_cls(self.size_cb, "sizebox")
        for s in SIZES:
            self.size_cb.append_text(f"{s:g}")
        self.size_cb.get_child().set_width_chars(3)
        self.size_cb.connect("changed", self._size_changed)
        box.pack_start(self.size_cb, False, False, 4)
        self.fmt_btns = {}
        for key, mk, tip in (("b", "<b>B</b>", "Bold (Ctrl+B)"), ("i", "<i>I</i>", "Italic (Ctrl+I)"),
                             ("u", "<u>U</u>", "Underline (Ctrl+U)")):
            self.fmt_btns[key] = self.rbtn(box, None, tip, lambda k=key: self.set_fmt(k, True, toggle=True), markup=mk)
        self.menu_btn(box, "borders", "Borders", lambda: [
            {"label": "All borders", "icon": "borders", "cb": lambda: self.set_fmt("border", True)},
            {"label": "No border", "cb": lambda: self.set_fmt("border", None)}])
        self.menu_btn(box, "fill", "Fill color", lambda: [{"label": "No fill", "cb": lambda: self.set_fmt("bg", None)}, None] +
                      [{"label": n, "pixbuf": self._swatch(c), "cb": lambda c=c: self.set_fmt("bg", c)} for n, c in FILL_COLORS])
        self.menu_btn(box, "fcolor", "Font color", lambda: [{"label": "Automatic", "cb": lambda: self.set_fmt("fg", None)}, None] +
                      [{"label": c, "pixbuf": self._swatch(c), "cb": lambda c=c: self.set_fmt("fg", c)} for c in TEXT_COLORS])
        self.rsep(box)
        self.align_btns = {}
        for a in ("left", "center", "right"):
            self.align_btns[a] = self.rbtn(box, "al-" + a, f"Align {a}", lambda a=a: self.set_fmt(
                "align", None if self.sh.fmt(*self.act).get("align") == a else a))
        self.wrap_btn = self.rbtn(box, "wraptext", "Wrap text", lambda: self.set_fmt("wrap", True, toggle=True))
        self.rsep(box)
        self.num_cb = Gtk.ComboBoxText()
        add_cls(self.num_cb, "numbox")
        for name, code in NUM_FORMATS:
            self.num_cb.append(code, name)
        self.num_cb.append("__custom", "Custom")
        self.num_cb.connect("changed", self._num_changed)
        box.pack_start(self.num_cb, False, False, 0)
        self.rbtn(box, "dollar", "Currency format", lambda: self.set_fmt("num", '"$"#,##0.00'))
        self.rbtn(box, "percent", "Percent style", lambda: self.set_fmt("num", "0%"))
        self.rbtn(box, "comma", "Comma style", lambda: self.set_fmt("num", "#,##0.00"))
        self.rbtn(box, "dec-more", "Increase decimal", lambda: self.change_decimals(1))
        self.rbtn(box, "dec-less", "Decrease decimal", lambda: self.change_decimals(-1))
        self.rsep(box)
        self.menu_btn(box, "ins-row", "Insert", lambda: [
            {"label": "Insert rows", "icon": "ins-row", "cb": lambda: self.insert_rc(0)},
            {"label": "Insert columns", "icon": "ins-col", "cb": lambda: self.insert_rc(1)},
            {"label": "Insert sheet", "icon": "plus", "cb": self.add_sheet}])
        self.menu_btn(box, "del-row", "Delete", lambda: [
            {"label": "Delete rows", "icon": "del-row", "cb": lambda: self.delete_rc(0)},
            {"label": "Delete columns", "icon": "del-col", "cb": lambda: self.delete_rc(1)},
            {"label": "Delete sheet", "icon": "delete", "cb": lambda: self.del_sheet(self.si)}])
        self.rsep(box)
        self.menu_btn(box, "sigma", "AutoSum (Alt+=)", lambda: [
            {"label": n.title(), "cb": lambda n=n: self.autosum(n)} for n in ("SUM", "AVERAGE", "COUNT", "MAX", "MIN")])
        self.rbtn(box, "sort-az", "Sort A to Z", lambda: self.sort(False))
        self.rbtn(box, "sort-za", "Sort Z to A", lambda: self.sort(True))
        self.rbtn(box, "search", "Find (Ctrl+F)", lambda: self.show_find(False))

    def rib_insert(self, box):
        for t, name, ic in (("col", "Column", "ch-col"), ("bar", "Bar", "ch-bar"), ("line", "Line", "ch-line"),
                            ("pie", "Pie", "ch-pie")):
            self.rbtn(box, ic, f"Insert a {name.lower()} chart of the selected data", lambda t=t: self.insert_chart(t), text=name)
        self.rsep(box)
        self.menu_btn(box, "fx", "Insert a function", self._func_items, text="Function")
        self.rbtn(box, "history", "Insert today's date (Ctrl+;)", lambda: self._type_now("date"), text="Date")
        self.rbtn(box, "ins-row", "Insert rows", lambda: self.insert_rc(0), text="Rows")
        self.rbtn(box, "ins-col", "Insert columns", lambda: self.insert_rc(1), text="Columns")

    def _func_items(self):
        return [{"label": g, "sub": [{"label": f, "cb": lambda f=f: self.insert_function(f)} for f in fs]} for g, fs in FUNC_GROUPS]

    def rib_formulas(self, box):
        self.menu_btn(box, "sigma", "AutoSum", lambda: [
            {"label": n.title(), "cb": lambda n=n: self.autosum(n)} for n in ("SUM", "AVERAGE", "COUNT", "MAX", "MIN")], text="AutoSum")
        for g, fs in FUNC_GROUPS:
            self.menu_btn(box, None, f"{g} functions", lambda fs=fs: [{"label": f, "cb": lambda f=f: self.insert_function(f)} for f in fs],
                          text=g)
        self.rsep(box)
        self.showf_btn = self.rbtn(box, None, "Show formulas instead of results (Ctrl+`)", self.toggle_formulas, text="Show formulas")
        self.rbtn(box, "refresh", "Recalculate", lambda: (self.book.recalc(), self.da.queue_draw()), text="Calculate")

    def rib_view(self, box):
        self.grid_btn = self.rbtn(box, "grid", "Show gridlines", self.toggle_grid, text="Gridlines")
        self.head_btn = self.rbtn(box, None, "Show row and column headings", self.toggle_heads, text="Headings")
        self.rsep(box)
        self.menu_btn(box, "freeze", "Freeze panes", lambda: [
            {"label": "Freeze panes", "cb": lambda: self.set_freeze(*self.act)},
            {"label": "Freeze top row", "cb": lambda: self.set_freeze(1, 0)},
            {"label": "Freeze first column", "cb": lambda: self.set_freeze(0, 1)},
            {"label": "Unfreeze panes", "cb": lambda: self.set_freeze(0, 0)}], text="Freeze panes")
        self.rsep(box)
        self.rbtn(box, "zoomout", "Zoom out", lambda: self.zoom_sc.set_value(self.zoom_sc.get_value() - 10))
        self.rbtn(box, "zoomin", "Zoom in", lambda: self.zoom_sc.set_value(self.zoom_sc.get_value() + 10))
        self.rbtn(box, None, "Zoom to 100%", lambda: self.zoom_sc.set_value(100), text="100%")

    def insert_function(self, name):
        if not self.editing:
            self.start_edit("=" + name + "(")
        else:
            pos = self.editor.get_position()
            t = self.editor.get_text()
            ins = name + "(" if t.startswith("=") else "=" + name + "("
            self.editor.set_text(t[:pos] + ins + t[pos:])
            self.editor.set_position(pos + len(ins))
            self.editor.grab_focus_without_selecting()

    def toggle_grid(self):
        self.show_grid = not self.show_grid
        self.st["grid"] = self.show_grid
        self._update_ribbon_state()
        self.da.queue_draw()

    def toggle_heads(self):
        self.show_heads = not self.show_heads
        self._update_ribbon_state()
        self._update_adj()
        self.da.queue_draw()

    def set_freeze(self, r, c):
        self.push_undo()
        self.sh.freeze = (r, c)
        self.hadj.set_value(0)
        self.vadj.set_value(0)
        self.mark_dirty()
        self._update_adj()
        self.da.queue_draw()

    def change_decimals(self, d):
        num = self.sh.fmt(*self.act).get("num") or "General"
        if num == "General":
            v = self.book.value(self.sh, *self.act)
            cur = 0
            if isinstance(v, float) and not v.is_integer():
                cur = min(9, len(general(v).split(".")[-1]))
            new = "0" + ("." + "0" * (cur + d) if cur + d > 0 else "")
        else:
            m = re.search(r"([0#])(\.0*)?", num)
            if not m:
                return
            cur = len(m.group(2) or ".") - 1
            n = max(0, cur + d)
            rep = m.group(1) + ("." + "0" * n if n else "")
            new = num[:m.start()] + rep + num[m.end():]
            # keep every section of the format in step
            if ";" in num:
                new = ";".join(re.sub(r"([0#])(\.0*)?", lambda mm: mm.group(1) + ("." + "0" * n if n else ""), s, count=1)
                               for s in num.split(";"))
        self.set_fmt("num", new)

    def _font_changed(self, cb):
        if self.updating or cb.get_active() < 0:
            return
        self.set_fmt("font", None if cb.get_active_text() == DEF_FONT else cb.get_active_text())
        self.da.grab_focus()

    def _size_changed(self, cb):
        if self.updating:
            return
        try:
            v = float((cb.get_active_text() or "").strip())
        except ValueError:
            return
        if cb.get_active() < 0 and not cb.get_child().has_focus():
            return
        if 1 <= v <= 409:
            self.set_fmt("size", None if v == DEF_SIZE else v)
        if cb.get_active() >= 0:
            self.da.grab_focus()

    def _num_changed(self, cb):
        if self.updating:
            return
        code = cb.get_active_id()
        if code and code != "__custom":
            self.set_fmt("num", None if code == "General" else code)
            self.da.grab_focus()

    @staticmethod
    def _swatch(c):
        pb = GdkPixbuf.Pixbuf.new(GdkPixbuf.Colorspace.RGB, False, 8, 16, 16)
        r, g, b = (int(c[i:i + 2], 16) for i in (1, 3, 5))
        pb.fill((r << 24) | (g << 16) | (b << 8) | 0xFF)
        return pb

    def _update_ribbon_state(self):
        f = self.sh.fmt(*self.act)
        for k, b in self.fmt_btns.items():
            (b.get_style_context().add_class if f.get(k) else b.get_style_context().remove_class)("on")
        for a, b in self.align_btns.items():
            (b.get_style_context().add_class if f.get("align") == a else b.get_style_context().remove_class)("on")
        (self.wrap_btn.get_style_context().add_class if f.get("wrap") else self.wrap_btn.get_style_context().remove_class)("on")
        (self.grid_btn.get_style_context().add_class if self.show_grid else self.grid_btn.get_style_context().remove_class)("on")
        (self.head_btn.get_style_context().add_class if self.show_heads else self.head_btn.get_style_context().remove_class)("on")
        (self.showf_btn.get_style_context().add_class if self.show_formulas else
         self.showf_btn.get_style_context().remove_class)("on")
        self.updating = True
        self.font_cb.get_child().set_text(f.get("font", DEF_FONT))
        self.size_cb.get_child().set_text(f"{f.get('size', DEF_SIZE):g}")
        num = f.get("num") or "General"
        if not self.num_cb.set_active_id(num):
            self.num_cb.set_active_id("__custom")
        self.updating = False

    def set_zoom(self, pct):
        pct = int(round(pct / 10) * 10)
        self.zoom = pct / 100
        self.st["zoom"] = pct
        self.zoom_lb.set_text(f"{pct}%")
        self._update_adj()
        self._place_editor()
        self.da.queue_draw()

    # ================================================================ find
    def _build_find(self):
        bar = Gtk.Box(spacing=6)
        add_cls(bar, "findbar")
        bar.set_halign(Gtk.Align.END)
        bar.set_valign(Gtk.Align.START)
        bar.set_margin_top(28)
        bar.set_margin_end(24)
        v = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        r1 = Gtk.Box(spacing=4)
        self.f_entry = Gtk.Entry()
        add_cls(self.f_entry, "field")
        self.f_entry.set_placeholder_text("Find")
        self.f_entry.set_width_chars(24)
        self.f_entry.connect("activate", lambda e: self.find_next())
        r1.pack_start(self.f_entry, True, True, 0)
        r1.pack_start(icon_button("up", lambda w: self.find_next(back=True), tip="Previous"), False, False, 0)
        r1.pack_start(icon_button("chevdown", lambda w: self.find_next(), tip="Next"), False, False, 0)
        r1.pack_start(icon_button("close", lambda w: (bar.hide(), self.da.grab_focus()), tip="Close"), False, False, 0)
        v.pack_start(r1, False, False, 0)
        self.r_row = Gtk.Box(spacing=4)
        self.r_entry = Gtk.Entry()
        add_cls(self.r_entry, "field")
        self.r_entry.set_placeholder_text("Replace with")
        self.r_row.pack_start(self.r_entry, True, True, 0)
        b1 = Gtk.Button(label="Replace")
        add_cls(b1, "secondary")
        b1.connect("clicked", lambda w: self.replace_one())
        b2 = Gtk.Button(label="All")
        add_cls(b2, "secondary")
        b2.connect("clicked", lambda w: self.replace_all())
        self.r_row.pack_start(b1, False, False, 0)
        self.r_row.pack_start(b2, False, False, 0)
        v.pack_start(self.r_row, False, False, 0)
        self.f_info = label("", "ter")
        v.pack_start(self.f_info, False, False, 0)
        bar.pack_start(v, True, True, 0)
        bar.show_all()
        return bar

    def show_find(self, replace):
        self.findbar.show()
        self.r_row.set_visible(replace)
        self.f_entry.grab_focus()

    def _matches(self):
        q = self.f_entry.get_text().lower()
        if not q:
            return []
        out = []
        for (r, c), cell in self.sh.cells.items():
            if q in cell["raw"].lower() or q in display_text(self.book, self.sh, r, c).lower():
                out.append((r, c))
        return sorted(out)

    def find_next(self, back=False):
        m = self._matches()
        if not m:
            self.f_info.set_text("No results")
            return False
        cur = self.act
        if back:
            prev = [p for p in m if p < cur]
            nxt = prev[-1] if prev else m[-1]
        else:
            after = [p for p in m if p > cur]
            nxt = after[0] if after else m[0]
        self.select(*nxt)
        self.ensure_visible(*nxt)
        self.f_info.set_text(f"{m.index(nxt) + 1} of {len(m)}")
        return True

    def replace_one(self):
        q, rep = self.f_entry.get_text(), self.r_entry.get_text()
        if not q:
            return
        raw = self.sh.raw(*self.act)
        if q.lower() in raw.lower():
            self.push_undo()
            self.set_raw(*self.act, re.sub(re.escape(q), rep.replace("\\", "\\\\"), raw, flags=re.I))
            self.changed()
        self.find_next()

    def replace_all(self):
        q, rep = self.f_entry.get_text(), self.r_entry.get_text()
        if not q:
            return
        self.push_undo()
        n = 0
        for (r, c), cell in list(self.sh.cells.items()):
            if q.lower() in cell["raw"].lower():
                self.set_raw(r, c, re.sub(re.escape(q), rep.replace("\\", "\\\\"), cell["raw"], flags=re.I))
                n += 1
        self.changed()
        self.f_info.set_text(f"Replaced {n}")

    # ================================================================ files
    def file_menu(self, w):
        recent = [p for p in self.st.get("recent", []) if os.path.exists(p)][:8]
        items = [{"label": "New", "icon": "new", "accel": "Ctrl+N", "cb": lambda: self.app.new_window()},
                 {"label": "Open...", "icon": "open", "accel": "Ctrl+O", "cb": self.open_dialog},
                 {"label": "Save", "icon": "save", "accel": "Ctrl+S", "cb": self.save},
                 {"label": "Save as...", "icon": "save", "accel": "Ctrl+Shift+S", "cb": self.save_as},
                 {"label": "Export as PDF...", "icon": "pdf", "cb": self.export_pdf}, None]
        if recent:
            items.append({"label": "Recent", "sub": [{"label": os.path.basename(p), "cb": lambda p=p: self.open_path(p)} for p in recent]})
            items.append(None)
        items.append({"label": "Close", "icon": "close", "accel": "Ctrl+W", "cb": self.close})
        popup_menu(build_menu(items), w)

    def load_book(self, book, path):
        self.book = book
        self.si = getattr(book, "active", 0)
        self.path = path
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.clip = None
        self.dirty = False
        self.book.recalc()
        self.build_sheet_tabs()
        self.switch_sheet(self.si)
        self.update_title()

    def is_blank(self):
        return self.path is None and not self.dirty and all(not s.cells for s in self.book.sheets)

    def update_title(self):
        name = os.path.basename(self.path) if self.path else "Book1"
        mod = "*" if self.dirty else ""
        self.set_title(f"{mod}{name} - Excel")
        self.title_lb.set_text(f"{mod}{name} - Excel")

    def open_dialog(self):
        dlg = Gtk.FileChooserNative.new("Open", self, Gtk.FileChooserAction.OPEN, "Open", "Cancel")
        for name, pats in (("Spreadsheets (.xlsx, .xlsm, .csv)", ["*.xlsx", "*.xlsm", "*.csv", "*.tsv"]), ("All files", ["*"])):
            f = Gtk.FileFilter()
            f.set_name(name)
            for p in pats:
                f.add_pattern(p)
            dlg.add_filter(f)
        dlg.set_current_folder(os.path.dirname(self.path) if self.path else os.path.join(HOME, "Documents"))
        if dlg.run() == Gtk.ResponseType.ACCEPT:
            path = dlg.get_filename()
            if self.is_blank():
                self.open_path(path)
            else:
                self.app.new_window(path)

    def open_path(self, path):
        try:
            if path.lower().endswith((".csv", ".tsv", ".txt")):
                book = load_csv(path)
            else:
                book = load_xlsx(path)
        except Exception as e:
            info_dialog(self, "Couldn't open this file", f"{os.path.basename(path)}\n{e}")
            return
        self.load_book(book, path)
        self._remember(path)

    def _remember(self, path):
        r = [p for p in self.st.get("recent", []) if p != path]
        self.st["recent"] = [path] + r[:14]
        self.st.save()
        try:
            Gtk.RecentManager.get_default().add_item(GLib.filename_to_uri(path, None))
        except GLib.Error:
            pass

    def save(self):
        if self.editing:
            self.commit_edit()
        if not self.path or not self.path.lower().endswith((".xlsx", ".xlsm", ".csv")):
            return self.save_as()
        return self.write(self.path)

    def save_as(self):
        if self.editing:
            self.commit_edit()
        dlg = Gtk.FileChooserNative.new("Save as", self, Gtk.FileChooserAction.SAVE, "Save", "Cancel")
        dlg.set_do_overwrite_confirmation(True)
        f = Gtk.FileFilter()
        f.set_name("Excel workbook (.xlsx)")
        f.add_pattern("*.xlsx")
        dlg.add_filter(f)
        f2 = Gtk.FileFilter()
        f2.set_name("CSV, current sheet only (.csv)")
        f2.add_pattern("*.csv")
        dlg.add_filter(f2)
        dlg.set_current_folder(os.path.dirname(self.path) if self.path else os.path.join(HOME, "Documents"))
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "Book1"
        dlg.set_current_name(base + ".xlsx")
        if dlg.run() != Gtk.ResponseType.ACCEPT:
            return False
        path = dlg.get_filename()
        if dlg.get_filter() is f2 and not path.lower().endswith(".csv"):
            path = os.path.splitext(path)[0] + ".csv"
        elif not path.lower().endswith((".xlsx", ".csv")):
            path += ".xlsx"
        return self.write(path)

    def write(self, path):
        try:
            if path.lower().endswith(".csv"):
                save_csv(self.book, self.sh, path)
            else:
                tmp = path + ".saving"
                save_xlsx(self.book, tmp, self.si)
                os.replace(tmp, path)
        except Exception as e:
            info_dialog(self, "Couldn't save", str(e))
            return False
        self.path = path
        self.dirty = False
        self._remember(path)
        self.update_title()
        return True

    def export_pdf(self):
        if self.editing:
            self.commit_edit()
        dlg = Gtk.FileChooserNative.new("Export as PDF", self, Gtk.FileChooserAction.SAVE, "Export", "Cancel")
        dlg.set_do_overwrite_confirmation(True)
        dlg.set_current_folder(os.path.dirname(self.path) if self.path else os.path.join(HOME, "Documents"))
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "Book1"
        dlg.set_current_name(base + ".pdf")
        if dlg.run() != Gtk.ResponseType.ACCEPT:
            return
        path = dlg.get_filename()
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            self.render_pdf(path)
        except Exception as e:
            info_dialog(self, "Couldn't export", str(e))
            return
        confirm(self, "Exported", os.path.basename(path) + " was saved.", "Open it",
                lambda ok: ok and subprocess.Popen(["xdg-open", path], start_new_session=True))

    def render_pdf(self, path):
        """Print the used range of the current sheet on Letter pages, splitting by rows and columns."""
        mr, mc = self.sh.extent()
        for chd in self.sh.charts:
            mr = max(mr, chd["row"] + int(chd["h"] / DEF_ROWH) + 1)
            c, x = chd["col"], 0
            while x < chd["w"]:
                x += self.sh.colw.get(c, DEF_COLW)
                c += 1
            mc = max(mc, c)
        if mr < 0:
            mr, mc = 0, 0
        W, H, M = 612, 792, 36
        total_w = sum(self.sh.colw.get(c, DEF_COLW) for c in range(mc + 1))
        if total_w * 0.75 > W - 2 * M and total_w > sum(self.sh.rowh.get(r, DEF_ROWH) for r in range(mr + 1)):
            W, H = H, W     # wide sheets print landscape
        # shrink to fit the page width when that keeps the text readable
        scale = max(0.45, min(0.75, (W - 2 * M) / max(1, total_w)))
        save = (self.zoom, self.show_heads, self.sel, self.act, self.chart_sel, self.hadj.get_value(), self.vadj.get_value(),
                self.sh.freeze, self.editing, self.show_grid)
        self.zoom, self.show_heads, self.chart_sel, self.editing, self.show_grid = 1.0, False, None, None, False
        self.sh.freeze = (0, 0)
        self.sel = (-5, -5, -5, -5)
        surf = cairo.PDFSurface(path, W, H)
        cr = cairo.Context(surf)
        try:
            self._update_adj()
            col_breaks = [0]
            x = 0
            for c in range(mc + 1):
                w = self.sh.colw.get(c, DEF_COLW) * scale
                if x + w > W - 2 * M + 0.5 and x > 0:
                    col_breaks.append(c)
                    x = 0
                x += w
            col_breaks.append(mc + 1)
            row_breaks = [0]
            y = 0
            for r in range(mr + 1):
                h = self.sh.rowh.get(r, DEF_ROWH) * scale
                if y + h > H - 2 * M + 0.5 and y > 0:
                    row_breaks.append(r)
                    y = 0
                y += h
            row_breaks.append(mr + 1)
            for ci in range(len(col_breaks) - 1):
                for ri in range(len(row_breaks) - 1):
                    c0, c1 = col_breaks[ci], col_breaks[ci + 1] - 1
                    r0, r1 = row_breaks[ri], row_breaks[ri + 1] - 1
                    self.hadj.set_value(self.col_x(c0))
                    self.vadj.set_value(self.row_y(r0))
                    cr.save()
                    cr.translate(M, M)
                    cr.scale(scale, scale)
                    pw = self.col_x(c1 + 1) - self.col_x(c0)
                    ph = self.row_y(r1 + 1) - self.row_y(r0)
                    cr.rectangle(0, 0, pw + 1, ph + 1)
                    cr.clip()
                    self._draw_cells(cr, (r0, r1), (c0, c1))
                    self._draw_charts(cr)
                    cr.restore()
                    cr.show_page()
        finally:
            surf.finish()
            (self.zoom, self.show_heads, self.sel, self.act, self.chart_sel, hv, vv, self.sh.freeze, self.editing,
             self.show_grid) = save
            self._update_adj()
            self.hadj.set_value(hv)
            self.vadj.set_value(vv)
            self.da.queue_draw()

    # ================================================================ misc
    def _cfg(self, *a):
        if not self.is_maximized():
            self.st["width"], self.st["height"] = self.get_size()
        self.st["max"] = self.is_maximized()
        return False

    def _on_close(self, *a):
        self.st.save()
        if self.editing:
            self.commit_edit()
        if not self.dirty:
            return False
        name = os.path.basename(self.path) if self.path else "Book1"
        dlg = WinDialog(self, "Excel", f"Do you want to save your changes to {name}?", primary="Save", secondary="Don't save")
        cancel = Gtk.Button(label="Cancel")
        add_cls(cancel, "secondary")
        cancel.connect("clicked", lambda w: (setattr(dlg, "cb", None), dlg.destroy()))
        dlg.foot.pack_start(cancel, True, True, 0)

        def done(ok):
            if ok and not self.save():
                return
            self.dirty = False
            self.destroy()
        dlg.ask(done)
        return True


def to_str_safe(v):
    try:
        return to_str(v)
    except XlRaise as e:
        return e.err.code


def draw_chart(cr, chd, cats, series, w, h, z):
    cr.set_source_rgb(1, 1, 1)
    cr.rectangle(0, 0, w, h)
    cr.fill()
    cr.set_source_rgb(0.85, 0.85, 0.85)
    cr.set_line_width(1)
    cr.rectangle(0.5, 0.5, w - 1, h - 1)
    cr.stroke()
    lay = PangoCairo.create_layout(cr)
    small = Pango.FontDescription(f"Carlito {max(6, 9 * z):.1f}")
    top = 10 * z
    if chd.get("title"):
        lay.set_font_description(Pango.FontDescription(f"Carlito {14 * z:.1f}"))
        lay.set_text(chd["title"], -1)
        tw, th = lay.get_pixel_size()
        cr.set_source_rgb(0.35, 0.35, 0.35)
        cr.move_to((w - tw) / 2, 8 * z)
        PangoCairo.show_layout(cr, lay)
        top = 8 * z + th + 8 * z
    lay.set_font_description(small)
    if not series or not cats:
        return
    typ = chd["type"]
    legend = typ == "pie" or len(series) > 1
    bottom = h - (24 * z if legend else 8 * z)
    if typ == "pie":
        vals = [max(0.0, v) for v in series[0][1]]
        tot = sum(vals) or 1
        cx, cy = w / 2, (top + bottom) / 2
        rad = max(10, min(w / 2 - 20 * z, (bottom - top) / 2 - 4 * z))
        a = -math.pi / 2
        for i, v in enumerate(vals):
            da = 2 * math.pi * v / tot
            cr.move_to(cx, cy)
            cr.arc(cx, cy, rad, a, a + da)
            cr.close_path()
            cr.set_source_rgba(*hexc(PALETTE[i % len(PALETTE)]))
            cr.fill_preserve()
            cr.set_source_rgb(1, 1, 1)
            cr.set_line_width(1.5)
            cr.stroke()
            if v / tot > 0.04:
                mid = a + da / 2
                lay.set_text(f"{v / tot * 100:.0f}%", -1)
                tw, th = lay.get_pixel_size()
                cr.move_to(cx + math.cos(mid) * rad * 0.65 - tw / 2, cy + math.sin(mid) * rad * 0.65 - th / 2)
                PangoCairo.show_layout(cr, lay)
            a += da
        names = cats
    else:
        allv = [v for _, vals in series for v in vals] or [0]
        lo, hi = min(0.0, min(allv)), max(0.0, max(allv))
        ticks = nice_ticks(lo, hi)
        lo, hi = ticks[0], ticks[-1]
        lab_w = 0
        for t in ticks:
            lay.set_text(general(t), -1)
            lab_w = max(lab_w, lay.get_pixel_size()[0])
        horiz = typ == "bar"
        if horiz:
            cat_w = 0
            for cname in cats:
                lay.set_text(cname, -1)
                cat_w = max(cat_w, lay.get_pixel_size()[0])
            left = min(w * 0.35, cat_w + 12 * z)
            px0, px1, py0, py1 = left, w - 14 * z, top, bottom - 16 * z
        else:
            px0, px1, py0, py1 = lab_w + 14 * z, w - 12 * z, top, bottom - 18 * z
        span = (hi - lo) or 1
        cr.set_line_width(1)
        for t in ticks:
            if horiz:
                x = px0 + (t - lo) / span * (px1 - px0)
                cr.set_source_rgb(0.88, 0.88, 0.88)
                cr.move_to(int(x) + 0.5, py0)
                cr.line_to(int(x) + 0.5, py1)
                cr.stroke()
                lay.set_text(general(t), -1)
                tw, th = lay.get_pixel_size()
                cr.set_source_rgb(0.4, 0.4, 0.4)
                cr.move_to(x - tw / 2, py1 + 3 * z)
            else:
                y = py1 - (t - lo) / span * (py1 - py0)
                cr.set_source_rgb(0.88, 0.88, 0.88)
                cr.move_to(px0, int(y) + 0.5)
                cr.line_to(px1, int(y) + 0.5)
                cr.stroke()
                lay.set_text(general(t), -1)
                tw, th = lay.get_pixel_size()
                cr.set_source_rgb(0.4, 0.4, 0.4)
                cr.move_to(px0 - tw - 6 * z, y - th / 2)
            PangoCairo.show_layout(cr, lay)
        n = len(cats)
        ns = len(series)
        if horiz:
            band = (py1 - py0) / n
            zero = px0 + (0 - lo) / span * (px1 - px0)
            for i, cname in enumerate(cats):
                y0 = py0 + i * band
                bh = band * 0.7 / ns
                for s, (_, vals) in enumerate(series):
                    x = px0 + (vals[i] - lo) / span * (px1 - px0)
                    cr.set_source_rgba(*hexc(PALETTE[s % len(PALETTE)]))
                    cr.rectangle(min(zero, x), y0 + band * 0.15 + s * bh, abs(x - zero), bh)
                    cr.fill()
                lay.set_text(cname, -1)
                tw, th = lay.get_pixel_size()
                cr.set_source_rgb(0.4, 0.4, 0.4)
                cr.move_to(px0 - tw - 6 * z, y0 + band / 2 - th / 2)
                PangoCairo.show_layout(cr, lay)
        else:
            band = (px1 - px0) / n
            zero = py1 - (0 - lo) / span * (py1 - py0)
            step = max(1, int(n / max(1, (px1 - px0) / (40 * z))))
            for i, cname in enumerate(cats):
                x0 = px0 + i * band
                if typ == "col":
                    bw = band * 0.7 / ns
                    for s, (_, vals) in enumerate(series):
                        y = py1 - (vals[i] - lo) / span * (py1 - py0)
                        cr.set_source_rgba(*hexc(PALETTE[s % len(PALETTE)]))
                        cr.rectangle(x0 + band * 0.15 + s * bw, min(zero, y), bw, abs(zero - y))
                        cr.fill()
                if i % step == 0:
                    lay.set_text(cname, -1)
                    tw, th = lay.get_pixel_size()
                    cr.set_source_rgb(0.4, 0.4, 0.4)
                    cr.move_to(x0 + band / 2 - tw / 2, py1 + 3 * z)
                    PangoCairo.show_layout(cr, lay)
            if typ == "line":
                for s, (_, vals) in enumerate(series):
                    cr.set_source_rgba(*hexc(PALETTE[s % len(PALETTE)]))
                    cr.set_line_width(2.2 * z)
                    cr.set_line_join(cairo.LINE_JOIN_ROUND)
                    for i, v in enumerate(vals):
                        x = px0 + i * band + band / 2
                        y = py1 - (v - lo) / span * (py1 - py0)
                        (cr.move_to if i == 0 else cr.line_to)(x, y)
                    cr.stroke()
                    for i, v in enumerate(vals):
                        x = px0 + i * band + band / 2
                        y = py1 - (v - lo) / span * (py1 - py0)
                        cr.arc(x, y, 3 * z, 0, 2 * math.pi)
                        cr.fill()
        cr.set_source_rgb(0.75, 0.75, 0.75)
        cr.set_line_width(1)
        if horiz:
            cr.move_to(int(zero) + 0.5, py0)
            cr.line_to(int(zero) + 0.5, py1)
        else:
            cr.move_to(px0, int(zero) + 0.5)
            cr.line_to(px1, int(zero) + 0.5)
        cr.stroke()
        names = [s[0] for s in series]
    if legend:
        items = []
        for i, nm in enumerate(names[:12]):
            lay.set_text(nm, -1)
            items.append((nm, lay.get_pixel_size()[0]))
        total = sum(tw + 22 * z for _, tw in items)
        x = max(6, (w - total) / 2)
        y = h - 18 * z
        for i, (nm, tw) in enumerate(items):
            cr.set_source_rgba(*hexc(PALETTE[i % len(PALETTE)]))
            cr.rectangle(x, y + 3 * z, 8 * z, 8 * z)
            cr.fill()
            lay.set_text(nm, -1)
            cr.set_source_rgb(0.35, 0.35, 0.35)
            cr.move_to(x + 11 * z, y)
            PangoCairo.show_layout(cr, lay)
            x += tw + 22 * z


class ExcelApp(Gtk.Application):
    def __init__(self):
        super().__init__(application_id=APP_ID, flags=Gio.ApplicationFlags.HANDLES_COMMAND_LINE)

    def do_startup(self):
        Gtk.Application.do_startup(self)
        self.settings = JsonSettings("pi-excel", DEFAULTS)
        apply_css(CSS)

    def new_window(self, path=None):
        w = ExcelWindow(self, path)
        w.present()
        return w

    def do_command_line(self, cmd):
        files = [a for a in cmd.get_arguments()[1:] if not a.startswith("-")]
        if files:
            for f in files:
                p = cmd.create_file_for_arg(f).get_path()
                wins = [w for w in self.get_windows() if isinstance(w, ExcelWindow) and w.is_blank()]
                if wins:
                    wins[0].open_path(p)
                    wins[0].present()
                else:
                    self.new_window(p)
        else:
            self.new_window()
        return 0


def main():
    GLib.set_prgname("pi-excel")
    GLib.set_application_name("Excel")
    app = ExcelApp()
    signal.signal(signal.SIGINT, signal.SIG_DFL)
    return app.run(sys.argv)


if __name__ == "__main__":
    sys.exit(main())
