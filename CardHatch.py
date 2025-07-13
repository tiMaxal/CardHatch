"""
MIT License

Copyright (c) 2025 tiMaxal `CardHatch`


perplexity-ai; tdwm20250709

prompt 1:
create py app,
 that is able to format a printable pdf from 2 spreadsheet columns as lists,
   to business cards [flashcards] with front and back [the 2 lists] aligned

prompt 2:
provide the code as a gui that accepts values for excel/csv file location,
 column names [incl. extra for 'index' + 'notes'
   - if possible, being able to add configurable bars top + bottom for color-coding is desired],
     cards_per_row, page-size and margins [with defaults prefilled first run,
       and settings file to save current values for next operation];

prompt 3:
apply grid columns like the following example for equi-spacing 'start' and 'exit'
 buttons at the bottom [and similarly, other entry fields in the main app window] -
# Configure columns for centering
frame_buttons.columnconfigure(0, weight=1) # left spacer
frame_buttons.columnconfigure(1, weight=0) # left button [start]
frame_buttons.columnconfigure(2, weight=1) # center spacer
frame_buttons.columnconfigure(3, weight=0) # right button [exit]
frame_buttons.columnconfigure(4, weight=1) # right spacer

prompt 4:
add color pickers for the top\bottom bar options

prompt 5:
- add a field to assign output field [default as origin of excel\csv input]
- also wrap text content of input columns to fit output of card width and available lines
     [alert and stop if text overflow, ie out of room - allow 'truncate' checkbox]
- format output with lines demarking cards, for separation [cutting]

20250710:

create output pdf as front/back pages alternating, for print both sides;
ensure 'back' cells will align with associated 'front' cells when page is flipped - requires:
- radio button selection for 'flip long edge' or 'flip short edge'
- if long, change order of each row
- if short, invert whole page, but ensure cells will align vertically


logic must assess how many cards fit on each page,
 then create a front and a back page for that many cards,
   then continue to the next amount that fit to the next page,
     and create those as front and back,
       forming a final output that alternates front and rear pages,
         with appropriate alignment

grok-ai:
centre cards on page [both vertically n horizontally],
 to ensure alignment when flipped during print

centre the text in each card [both vertically and horizontally];
also provide options to choose Font size, family [and style - bold\italic\etc] by type\dropdown, and colour

provide the complete code, incl full docstrings,
 with good practice logging added to all stages
   [and include logging to file, in cwd]

remove all controls and references to 'index' and 'notes' columns
provide complete code, with full logging to file and docstrings

20250711

 also function with .ods files
csv, excel and ods types all be shown at the same time in the picker,
 instead of needing to use a drop-down
  [and just show 'all' in the drop-down as well]

include a `Various amounts from 'qty' col` checkbox
 to apply diff quantity to certain cards [via column 'qty'],
   and a `Multiple` value input box to multiply *all* cards by a same amount
"""

"""
Install Required Packages:
 bash
pip install pandas reportlab openpyxl

"""

"""
Flashcard PDF Generator

This script creates a Tkinter-based GUI application that generates a double-sided flashcard PDF from a CSV, Excel, or ODS file.
The application allows users to select an input file (showing CSV, Excel, and ODS files simultaneously by default), specify column
names for front and back content, card layout, page size, margins, color bars, flip mode for duplex printing, font size, font family,
font style, and text color. Users can also specify variable card quantities via a 'qty' column and a global quantity multiplier.
Cards are centered on the page both horizontally and vertically, with text centered within each card both horizontally and vertically.
Front and back pages are aligned for duplex printing based on the selected flip mode (long or short edge). Settings are saved to a JSON
file for persistence. Comprehensive logging tracks all stages of the process, with logs written to both console and a file in the
current working directory.

Dependencies:
- pandas
- reportlab
- openpyxl
- odfpy

Install requirements: pip install pandas reportlab openpyxl odfpy
"""

import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import json
import os
import math
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.pdfbase.pdfmetrics import stringWidth
import logging
import uuid

# ---------- Logging Setup ----------

# Configure logging to output to both console and a file in the current working directory
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(),  # Console output
        logging.FileHandler("cardhatch.log"),  # File output in cwd
    ],
)
logger = logging.getLogger(__name__)

# ---------- Settings Management ----------

DEFAULT_SETTINGS = {
    "file_path": "",
    "output_file": "",
    "front_column": "Front",
    "back_column": "Back",
    "cards_per_row": 3,
    "card_width": 85.6,  # mm, default credit card size
    "card_height": 55.0,  # mm, default credit card size
    "page_size": "210x297",  # A4 in mm
    "margins": "10,10,10,10",  # top,bottom,left,right in mm
    "color_bar_top": False,
    "color_bar_bottom": False,
    "color_bar_top_color": "#FF0000",
    "color_bar_bottom_color": "#0000FF",
    "truncate": False,
    "flip_mode": "long",
    "font_size": 12,
    "font_family": "Helvetica",
    "font_style": "Normal",
    "text_color": "#000000",
    "use_qty_column": False,
    "quantity_multiplier": 1,
}

SETTINGS_FILE = "cardhatch_settings.json"


def load_settings():
    """
    Load settings from a JSON file or return default settings if the file doesn't exist.

    Returns:
        dict: Loaded or default settings.
    """
    logger.info(f"Loading settings from {SETTINGS_FILE}")
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r") as f:
                settings = json.load(f)
                logger.info("Settings loaded successfully")
                return settings
        else:
            logger.info("Settings file not found, using default settings")
            return DEFAULT_SETTINGS.copy()
    except Exception as e:
        logger.error(f"Failed to load settings: {e}")
        return DEFAULT_SETTINGS.copy()


settings = load_settings()

# ---------- Utility Functions ----------


def mm(val):
    """
    Convert millimeters to PDF points (1 mm = 2.83465 points).

    Args:
        val (float): Value in millimeters.

    Returns:
        float: Value in PDF points.
    """
    return float(val) * 2.83465


def wrap_text(text, font_name, font_size, max_width_pt, max_lines, truncate=False):
    """
    Wrap text to fit within a maximum width and number of lines.

    Args:
        text (str): Text to wrap.
        font_name (str): Font name for text measurement.
        font_size (float): Font size in points.
        max_width_pt (float): Maximum width in points.
        max_lines (int): Maximum number of lines.
        truncate (bool): If True, truncate text with ellipsis; otherwise, indicate overflow.

    Returns:
        tuple: (list of wrapped lines, bool indicating if text overflowed).
    """
    logger.debug(
        f"Wrapping text: {text[:50]}... (max_width={max_width_pt}, max_lines={max_lines})"
    )
    words = str(text).split()
    lines = []
    current_line = ""
    for word in words:
        test_line = (current_line + " " + word).strip()
        if stringWidth(test_line, font_name, font_size) <= max_width_pt:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word
            if len(lines) == max_lines:
                if truncate:
                    lines[-1] = lines[-1][: int(max_width_pt / font_size)] + "..."
                    logger.debug("Text truncated to fit")
                    return lines, False
                logger.warning("Text overflow detected")
                return lines, True
    if current_line:
        lines.append(current_line)
    if len(lines) > max_lines:
        if truncate:
            lines = lines[:max_lines]
            lines[-1] = lines[-1][: int(max_width_pt / font_size)] + "..."
            logger.debug("Text truncated to fit")
            return lines, False
        logger.warning("Text overflow detected")
        return lines[:max_lines], True
    logger.debug(f"Text wrapped into {len(lines)} lines")
    return lines, False


def draw_cut_lines(
    c,
    page_w_pt,
    page_h_pt,
    offset_left_pt,
    offset_top_pt,
    card_width_pt,
    card_height_pt,
    cards_per_row,
    cards_per_col,
):
    """
    Draw cut lines on the PDF for card separation.

    Args:
        c (Canvas): ReportLab canvas object.
        page_w_pt (float): Page width in points.
        page_h_pt (float): Page height in points.
        offset_left_pt (float): Left offset for card grid in points.
        offset_top_pt (float): Top offset for card grid in points.
        card_width_pt (float): Card width in points.
        card_height_pt (float): Card height in points.
        cards_per_row (int): Number of cards per row.
        cards_per_col (int): Number of cards per column.
    """
    logger.debug(f"Drawing cut lines for {cards_per_row}x{cards_per_col} grid")
    c.setStrokeColor(HexColor("#888888"))
    c.setLineWidth(0.5)
    # Vertical lines
    for i in range(cards_per_row + 1):
        x = offset_left_pt + i * card_width_pt
        c.line(x, offset_top_pt, x, page_h_pt - offset_top_pt)
    # Horizontal lines
    for j in range(cards_per_col + 1):
        y = page_h_pt - offset_top_pt - j * card_height_pt
        c.line(offset_left_pt, y, page_w_pt - offset_left_pt, y)
    logger.debug("Cut lines drawn")


def reorder_for_back(page_indices, cards_per_row, cards_per_col, flip_mode):
    """
    Reorder card indices for the back page based on flip mode for duplex printing.

    Args:
        page_indices (list): List of card indices for the page.
        cards_per_row (int): Number of cards per row.
        cards_per_col (int): Number of cards per column.
        flip_mode (str): 'long' for long edge flip, 'short' for short edge flip.

    Returns:
        list: Reordered indices for the back page.
    """
    logger.debug(f"Reordering indices for back page with flip_mode={flip_mode}")
    reordered = []
    if flip_mode == "long":
        for r in range(cards_per_col):
            row = page_indices[r * cards_per_row : (r + 1) * cards_per_row]
            reordered.extend(row[::-1])
    else:  # short edge
        for r in reversed(range(cards_per_col)):
            row = page_indices[r * cards_per_row : (r + 1) * cards_per_row]
            reordered.extend(row)
    logger.debug(f"Reordered indices: {reordered}")
    return reordered


# ---------- Main Application ----------


class FlashcardApp(tk.Tk):
    """
    Tkinter GUI application for generating printable, double-sided flashcard PDFs from spreadsheet data.

    The application provides fields for input file selection (showing CSV, Excel, and ODS files simultaneously by default),
    column names for front and back content, card layout, page size, margins, color bars, flip mode, font size, font family,
    font style, text color, variable quantities via a 'qty' column, and a global quantity multiplier. Cards are centered on
    the page, and text within each card is centered both horizontally and vertically. Front/back pages are aligned for duplex
    printing based on the flip mode. Settings are saved to a JSON file for persistence. Logging tracks all user interactions
    and processing steps.
    """

    def __init__(self):
        """Initialize the GUI with input fields prefilled from saved or default settings."""
        super().__init__()
        logger.info("Initializing FlashcardApp GUI")
        self.title("Flashcard PDF Generator")
        self.geometry("750x580")  # Adjusted for new fields

        frame_main = tk.Frame(self)
        frame_main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # File selection
        tk.Label(frame_main, text="CSV/Excel/ODS File Path:").grid(
            row=0, column=0, sticky=tk.W
        )
        self.entry_file = tk.Entry(frame_main)
        self.entry_file.grid(row=0, column=1, sticky=tk.EW)
        self.entry_file.insert(0, settings.get("file_path", ""))
        btn_browse = tk.Button(frame_main, text="Browse...", command=self.browse_file)
        btn_browse.grid(row=0, column=2)

        # Output file
        tk.Label(frame_main, text="Output PDF File:").grid(row=1, column=0, sticky=tk.W)
        self.entry_output = tk.Entry(frame_main)
        self.entry_output.grid(row=1, column=1, sticky=tk.EW)
        default_output = (
            os.path.splitext(settings.get("file_path", ""))[0] + ".pdf"
            if settings.get("file_path", "")
            else "flashcards.pdf"
        )
        self.entry_output.insert(0, settings.get("output_file", default_output))
        btn_output_browse = tk.Button(
            frame_main, text="Browse...", command=self.browse_output_file
        )
        btn_output_browse.grid(row=1, column=2)

        # Column names
        tk.Label(frame_main, text="Front Column Name:").grid(
            row=2, column=0, sticky=tk.W
        )
        self.entry_front = tk.Entry(frame_main)
        self.entry_front.grid(row=2, column=1, sticky=tk.EW)
        self.entry_front.insert(0, settings.get("front_column", "Front"))

        tk.Label(frame_main, text="Back Column Name:").grid(
            row=3, column=0, sticky=tk.W
        )
        self.entry_back = tk.Entry(frame_main)
        self.entry_back.grid(row=3, column=1, sticky=tk.EW)
        self.entry_back.insert(0, settings.get("back_column", "Back"))

        # Cards per row
        tk.Label(frame_main, text="Cards per Row:").grid(row=4, column=0, sticky=tk.W)
        self.entry_cards_per_row = tk.Entry(frame_main)
        self.entry_cards_per_row.grid(row=4, column=1, sticky=tk.EW)
        self.entry_cards_per_row.insert(0, str(settings.get("cards_per_row", 3)))

        # Card size
        tk.Label(frame_main, text="Card Width (mm):").grid(row=5, column=0, sticky=tk.W)
        self.entry_card_width = tk.Entry(frame_main)
        self.entry_card_width.grid(row=5, column=1, sticky=tk.EW)
        self.entry_card_width.insert(0, str(settings.get("card_width", 85.6)))

        tk.Label(frame_main, text="Card Height (mm):").grid(
            row=6, column=0, sticky=tk.W
        )
        self.entry_card_height = tk.Entry(frame_main)
        self.entry_card_height.grid(row=6, column=1, sticky=tk.EW)
        self.entry_card_height.insert(0, str(settings.get("card_height", 55.0)))

        # Page size
        tk.Label(frame_main, text="Page Size (WxH mm):").grid(
            row=7, column=0, sticky=tk.W
        )
        self.entry_page_size = tk.Entry(frame_main)
        self.entry_page_size.grid(row=7, column=1, sticky=tk.EW)
        self.entry_page_size.insert(0, settings.get("page_size", "210x297"))

        # Margins
        tk.Label(frame_main, text="Margins (Top,Bottom,Left,Right mm):").grid(
            row=8, column=0, sticky=tk.W
        )
        self.entry_margins = tk.Entry(frame_main)
        self.entry_margins.grid(row=8, column=1, sticky=tk.EW)
        self.entry_margins.insert(0, settings.get("margins", "10,10,10,10"))

        # Font settings
        tk.Label(frame_main, text="Font Size:").grid(row=9, column=0, sticky=tk.W)
        self.entry_font_size = tk.Entry(frame_main)
        self.entry_font_size.grid(row=9, column=1, sticky=tk.EW)
        self.entry_font_size.insert(0, str(settings.get("font_size", 12)))

        tk.Label(frame_main, text="Font Family:").grid(row=10, column=0, sticky=tk.W)
        self.font_family_var = tk.StringVar(
            value=settings.get("font_family", "Helvetica")
        )
        font_families = ["Helvetica", "Times-Roman", "Courier"]
        tk.OptionMenu(frame_main, self.font_family_var, *font_families).grid(
            row=10, column=1, sticky=tk.EW
        )

        tk.Label(frame_main, text="Font Style:").grid(row=11, column=0, sticky=tk.W)
        self.font_style_var = tk.StringVar(value=settings.get("font_style", "Normal"))
        font_styles = ["Normal", "Bold", "Italic", "BoldItalic"]
        tk.OptionMenu(frame_main, self.font_style_var, *font_styles).grid(
            row=11, column=1, sticky=tk.EW
        )

        tk.Label(frame_main, text="Text Color:").grid(row=12, column=0, sticky=tk.W)
        self.text_color_var = tk.StringVar(value=settings.get("text_color", "#000000"))
        btn_text_color = tk.Button(
            frame_main, text="Pick Text Color", command=self.pick_text_color
        )
        btn_text_color.grid(row=12, column=1, sticky=tk.W)
        self.lbl_text_color = tk.Label(
            frame_main,
            textvariable=self.text_color_var,
            bg=self.text_color_var.get(),
            width=10,
        )
        self.lbl_text_color.grid(row=12, column=2, sticky=tk.W)

        # Color bars with pickers
        self.color_bar_top_var = tk.BooleanVar(
            value=settings.get("color_bar_top", False)
        )
        self.color_bar_bottom_var = tk.BooleanVar(
            value=settings.get("color_bar_bottom", False)
        )
        self.color_bar_top_color = tk.StringVar(
            value=settings.get("color_bar_top_color", "#FF0000")
        )
        self.color_bar_bottom_color = tk.StringVar(
            value=settings.get("color_bar_bottom_color", "#0000FF")
        )

        tk.Checkbutton(
            frame_main, text="Add Color Bar Top", variable=self.color_bar_top_var
        ).grid(row=13, column=0, sticky=tk.W)
        btn_top_color = tk.Button(
            frame_main, text="Pick Top Color", command=self.pick_top_color
        )
        btn_top_color.grid(row=13, column=1, sticky=tk.W)
        self.lbl_top_color = tk.Label(
            frame_main,
            textvariable=self.color_bar_top_color,
            bg=self.color_bar_top_color.get(),
            width=10,
        )
        self.lbl_top_color.grid(row=13, column=2, sticky=tk.W)

        tk.Checkbutton(
            frame_main, text="Add Color Bar Bottom", variable=self.color_bar_bottom_var
        ).grid(row=14, column=0, sticky=tk.W)
        btn_bottom_color = tk.Button(
            frame_main, text="Pick Bottom Color", command=self.pick_bottom_color
        )
        btn_bottom_color.grid(row=14, column=1, sticky=tk.W)
        self.lbl_bottom_color = tk.Label(
            frame_main,
            textvariable=self.color_bar_bottom_color,
            bg=self.color_bar_bottom_color.get(),
            width=10,
        )
        self.lbl_bottom_color.grid(row=14, column=2, sticky=tk.W)

        # Truncate checkbox
        self.truncate_var = tk.BooleanVar(value=settings.get("truncate", False))
        tk.Checkbutton(
            frame_main, text="Truncate overflow text", variable=self.truncate_var
        ).grid(row=15, column=0, sticky=tk.W)

        # Quantity settings
        self.use_qty_column_var = tk.BooleanVar(
            value=settings.get("use_qty_column", False)
        )
        tk.Checkbutton(
            frame_main,
            text="Various amounts from 'qty' col",
            variable=self.use_qty_column_var,
        ).grid(row=16, column=0, sticky=tk.W)

        tk.Label(frame_main, text="Multiple (all cards):").grid(
            row=17, column=0, sticky=tk.W
        )
        self.entry_quantity_multiplier = tk.Entry(frame_main)
        self.entry_quantity_multiplier.grid(row=17, column=1, sticky=tk.EW)
        self.entry_quantity_multiplier.insert(
            0, str(settings.get("quantity_multiplier", 1))
        )

        # Flip mode
        self.flip_mode_var = tk.StringVar(value=settings.get("flip_mode", "long"))
        tk.Label(frame_main, text="Flip Mode:").grid(row=18, column=0, sticky=tk.W)
        tk.Radiobutton(
            frame_main,
            text="Flip on Long Edge",
            variable=self.flip_mode_var,
            value="long",
        ).grid(row=18, column=1, sticky=tk.W)
        tk.Radiobutton(
            frame_main,
            text="Flip on Short Edge",
            variable=self.flip_mode_var,
            value="short",
        ).grid(row=18, column=2, sticky=tk.W)

        frame_main.columnconfigure(0, weight=0)
        frame_main.columnconfigure(1, weight=1)
        frame_main.columnconfigure(2, weight=0)

        # Buttons frame
        frame_buttons = tk.Frame(self)
        frame_buttons.pack(fill=tk.BOTH, padx=10, pady=10)
        frame_buttons.columnconfigure(0, weight=1)
        frame_buttons.columnconfigure(1, weight=0)
        frame_buttons.columnconfigure(2, weight=1)
        frame_buttons.columnconfigure(3, weight=0)
        frame_buttons.columnconfigure(4, weight=1)

        btn_start = tk.Button(frame_buttons, text="Start", command=self.start_process)
        btn_start.grid(row=0, column=1, sticky=tk.EW)
        btn_exit = tk.Button(frame_buttons, text="Exit", command=self.quit)
        btn_exit.grid(row=0, column=3, sticky=tk.EW)

        logger.info("GUI initialization complete")

    def browse_file(self):
        """Open a file dialog to select the input CSV, Excel, or ODS file, showing all supported spreadsheet formats by default."""
        logger.info("Opening file dialog for input file selection")
        file_path = filedialog.askopenfilename(
            filetypes=[("Spreadsheets", "*.csv;*.xlsx;*.xls;*.ods"), ("All files", "*")]
        )
        if file_path:
            logger.info(f"Selected input file: {file_path}")
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, file_path)
            default_output = os.path.splitext(file_path)[0] + ".pdf"
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, default_output)
            logger.info(f"Set default output file: {default_output}")
        else:
            logger.info("No input file selected")

    def browse_output_file(self):
        """Open a file dialog to select the output PDF file."""
        logger.info("Opening file dialog for output file selection")
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            logger.info(f"Selected output file: {file_path}")
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, file_path)
        else:
            logger.info("No output file selected")

    def pick_top_color(self):
        """Open a color picker dialog for the top color bar."""
        logger.info("Opening color picker for top bar")
        color = colorchooser.askcolor(title="Pick Top Bar Color")[1]
        if color:
            logger.info(f"Selected top bar color: {color}")
            self.color_bar_top_color.set(color)
            self.lbl_top_color.config(bg=color)
        else:
            logger.info("No top bar color selected")

    def pick_bottom_color(self):
        """Open a color picker dialog for the bottom color bar."""
        logger.info("Opening color picker for bottom bar")
        color = colorchooser.askcolor(title="Pick Bottom Bar Color")[1]
        if color:
            logger.info(f"Selected bottom bar color: {color}")
            self.color_bar_bottom_color.set(color)
            self.lbl_bottom_color.config(bg=color)
        else:
            logger.info("No bottom bar color selected")

    def pick_text_color(self):
        """Open a color picker dialog for the text color."""
        logger.info("Opening color picker for text color")
        color = colorchooser.askcolor(title="Pick Text Color")[1]
        if color:
            logger.info(f"Selected text color: {color}")
            self.text_color_var.set(color)
            self.lbl_text_color.config(bg=color)
        else:
            logger.info("No text color selected")

    def start_process(self):
        """
        Validate inputs, save settings, load data, validate 'qty' column if used, and generate the PDF.

        Saves current GUI inputs to settings, reads the input file (CSV, Excel, or ODS), validates columns
        (including 'qty' if enabled), and calls the PDF generation function. Displays error messages via GUI
        if issues occur.
        """
        logger.info("Starting PDF generation process")
        # Save current settings
        settings["file_path"] = self.entry_file.get()
        settings["output_file"] = self.entry_output.get()
        settings["front_column"] = self.entry_front.get()
        settings["back_column"] = self.entry_back.get()
        settings["use_qty_column"] = self.use_qty_column_var.get()
        try:
            settings["cards_per_row"] = int(self.entry_cards_per_row.get())
            settings["card_width"] = float(self.entry_card_width.get())
            settings["card_height"] = float(self.entry_card_height.get())
            settings["page_size"] = self.entry_page_size.get()
            settings["margins"] = self.entry_margins.get()
            settings["font_size"] = float(self.entry_font_size.get())
            settings["quantity_multiplier"] = int(
                self.entry_quantity_multiplier.get() or 1
            )
            if settings["quantity_multiplier"] <= 0:
                raise ValueError("Quantity multiplier must be a positive integer")
        except ValueError as e:
            logger.error(f"Invalid input in numeric fields: {e}")
            messagebox.showerror("Error", f"Invalid input in numeric fields: {e}")
            return
        settings["color_bar_top"] = self.color_bar_top_var.get()
        settings["color_bar_bottom"] = self.color_bar_bottom_var.get()
        settings["color_bar_top_color"] = self.color_bar_top_color.get()
        settings["color_bar_bottom_color"] = self.color_bar_bottom_color.get()
        settings["truncate"] = self.truncate_var.get()
        settings["flip_mode"] = self.flip_mode_var.get()
        settings["font_family"] = self.font_family_var.get()
        settings["font_style"] = self.font_style_var.get()
        settings["text_color"] = self.text_color_var.get()

        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(settings, f, indent=4)
                logger.info(f"Settings saved to {SETTINGS_FILE}")
        except Exception as e:
            logger.error(f"Failed to save settings: {e}")
            messagebox.showerror("Error", f"Failed to save settings: {e}")
            return

        file_path = settings["file_path"]
        if not file_path:
            logger.error("No input file selected")
            messagebox.showerror("Error", "Please select an input file.")
            return

        logger.info(f"Reading input file: {file_path}")
        try:
            if file_path.lower().endswith(".csv"):
                data = pd.read_csv(file_path)
            elif file_path.lower().endswith((".xlsx", ".xls")):
                data = pd.read_excel(file_path, engine="openpyxl")
            elif file_path.lower().endswith(".ods"):
                data = pd.read_excel(file_path, engine="odf")
            else:
                logger.error("Unsupported file format. Please use CSV, Excel, or ODS.")
                messagebox.showerror(
                    "Error", "Unsupported file format. Please use CSV, Excel, or ODS."
                )
                return
            logger.info(f"Successfully read {len(data)} rows from input file")
        except Exception as e:
            logger.error(f"Failed to read file: {e}")
            messagebox.showerror("Error", f"Failed to read file: {e}")
            return

        # Validate columns
        missing = []
        for col in [settings["front_column"], settings["back_column"]]:
            if col not in data.columns:
                missing.append(col)
        if settings["use_qty_column"] and "qty" not in data.columns:
            missing.append("qty")
        if missing:
            logger.error(f"Missing columns in data: {', '.join(missing)}")
            messagebox.showerror(
                "Error", f"Missing columns in data: {', '.join(missing)}"
            )
            return

        # Validate qty column values if used
        if settings["use_qty_column"]:
            try:
                qty_values = pd.to_numeric(data["qty"], errors="coerce")
                if qty_values.isna().any():
                    logger.error("Non-numeric values found in 'qty' column")
                    messagebox.showerror(
                        "Error", "All values in 'qty' column must be numeric"
                    )
                    return
                if not qty_values.apply(float.is_integer).all():
                    logger.error("Non-integer values found in 'qty' column")
                    messagebox.showerror(
                        "Error", "All values in 'qty' column must be integers"
                    )
                    return
                if (qty_values <= 0).any():
                    logger.error("Non-positive values found in 'qty' column")
                    messagebox.showerror(
                        "Error", "All values in 'qty' column must be positive"
                    )
                    return
                logger.info("Validated 'qty' column successfully")
            except Exception as e:
                logger.error(f"Error validating 'qty' column: {e}")
                messagebox.showerror("Error", f"Error validating 'qty' column: {e}")
                return

        try:
            self.generate_flashcard_pdf(data, settings)
            logger.info(
                f"PDF generated successfully at {os.path.abspath(settings['output_file'])}"
            )
            messagebox.showinfo(
                "Success", f"PDF generated as {settings['output_file']}"
            )
        except Exception as e:
            logger.error(f"PDF generation failed: {e}")
            messagebox.showerror("Error", f"PDF generation failed: {e}")

    def generate_flashcard_pdf(self, data, settings):
        """
        Generate a flashcard PDF with alternating front/back pages, centered on the page, with variable quantities.

        Cards are arranged in a grid, centered both horizontally and vertically on the page. Text within each card
        is centered both horizontally and vertically. Front and back pages are aligned for duplex printing based on
        the flip mode. Includes optional color bars and supports customizable font size, family, style, text color,
        and variable card quantities via a 'qty' column and/or a global multiplier.

        Args:
            data (pandas.DataFrame): Input data with front/back columns and optional 'qty' column.
            settings (dict): Configuration settings for PDF generation.

        Raises:
            ValueError: If the card grid is too large for the page or other validation fails.
            Exception: If text overflow occurs without truncation or other PDF generation errors.
        """
        logger.info("Starting PDF generation")
        # Parse page size and margins
        try:
            page_w, page_h = map(float, settings["page_size"].split("x"))
            margins = list(map(float, settings["margins"].split(",")))
            margin_top, margin_bottom, margin_left, margin_right = margins
        except ValueError as e:
            logger.error(f"Invalid page size or margins format: {e}")
            raise ValueError(f"Invalid page size or margins format: {e}")

        cards_per_row = int(settings["cards_per_row"])
        card_width = float(settings["card_width"])
        card_height = float(settings["card_height"])
        font_size = float(settings["font_size"])
        font_family = settings["font_family"]
        font_style = settings["font_style"]
        text_color = settings["text_color"]
        use_qty_column = settings["use_qty_column"]
        quantity_multiplier = int(settings["quantity_multiplier"])

        # Map font style to ReportLab font name
        font_map = {
            ("Helvetica", "Normal"): "Helvetica",
            ("Helvetica", "Bold"): "Helvetica-Bold",
            ("Helvetica", "Italic"): "Helvetica-Oblique",
            ("Helvetica", "BoldItalic"): "Helvetica-BoldOblique",
            ("Times-Roman", "Normal"): "Times-Roman",
            ("Times-Roman", "Bold"): "Times-Bold",
            ("Times-Roman", "Italic"): "Times-Italic",
            ("Times-Roman", "BoldItalic"): "Times-BoldItalic",
            ("Courier", "Normal"): "Courier",
            ("Courier", "Bold"): "Courier-Bold",
            ("Courier", "Italic"): "Courier-Oblique",
            ("Courier", "BoldItalic"): "Courier-BoldOblique",
        }
        font_name = font_map.get((font_family, font_style), "Helvetica")
        logger.info(f"Using font: {font_name} with size {font_size}")

        page_size_pt = (mm(page_w), mm(page_h))
        card_width_pt = mm(card_width)
        card_height_pt = mm(card_height)

        # Calculate cards per column
        cards_per_col = int((page_h - margin_top - margin_bottom) // card_height)
        cards_per_page = cards_per_row * cards_per_col

        # Validate grid size
        grid_width = cards_per_row * card_width
        grid_height = cards_per_col * card_height
        if (
            grid_width + margin_left + margin_right > page_w
            or grid_height + margin_top + margin_bottom > page_h
        ):
            logger.error("Card grid with margins is too large for the page size")
            raise ValueError("Card grid with margins is too large for the page size.")
        if cards_per_col <= 0 or cards_per_row <= 0:
            logger.error(
                "Card size or margins too large, resulting in zero cards per page"
            )
            raise ValueError("Card size or margins too large for page size.")

        # Calculate centering offsets for the grid
        leftover_width = page_w - grid_width - margin_left - margin_right
        leftover_height = page_h - grid_height - margin_top - margin_bottom
        offset_left = margin_left + leftover_width / 2  # Center horizontally
        offset_top = margin_top + leftover_height / 2  # Center vertically
        offset_left_pt = mm(offset_left)
        offset_top_pt = mm(offset_top)
        logger.info(
            f"Calculated grid: {cards_per_row}x{cards_per_col}, offsets: {offset_left}mm, {offset_top}mm"
        )

        output_file = settings.get("output_file", "flashcards.pdf")
        logger.info(f"Creating PDF at: {os.path.abspath(output_file)}")
        c = canvas.Canvas(output_file, pagesize=page_size_pt)

        # Font settings
        line_height = font_size * 1.2

        # Generate card indices based on quantities
        card_indices = []
        for idx in range(len(data)):
            qty = int(data["qty"].iloc[idx]) if use_qty_column else 1
            qty = qty * quantity_multiplier
            card_indices.extend([idx] * qty)
        num_cards = len(card_indices)
        logger.info(
            f"Total cards to print: {num_cards} (use_qty_column={use_qty_column}, multiplier={quantity_multiplier})"
        )

        num_pages = math.ceil(num_cards / cards_per_page)
        flip_mode = settings.get("flip_mode", "long")
        logger.info(
            f"Generating {num_pages} pages for {num_cards} cards, flip_mode={flip_mode}"
        )

        # Helper for color bars
        def draw_color_bar(x, y, width, height, color_hex):
            c.setFillColor(HexColor(color_hex))
            c.rect(x, y, width, height, fill=1, stroke=0)

        for page in range(num_pages):
            logger.debug(f"Generating front page {page + 1}")
            # --- Front page ---
            page_start = page * cards_per_page
            page_end = min(page_start + cards_per_page, num_cards)
            page_indices = card_indices[page_start:page_end]
            # Pad with None if last page is not full
            while len(page_indices) < cards_per_page:
                page_indices.append(None)

            for pos_on_page, data_idx in enumerate(page_indices):
                card_row = pos_on_page // cards_per_row
                card_col = pos_on_page % cards_per_row
                x = offset_left_pt + card_col * card_width_pt
                y = page_size_pt[1] - offset_top_pt - (card_row + 1) * card_height_pt
                logger.debug(
                    f"Front card at pos {pos_on_page}, data_idx={data_idx}, x={x}, y={y}"
                )

                if data_idx is not None:
                    row = data.iloc[data_idx]
                    # Color bars
                    if settings["color_bar_top"]:
                        draw_color_bar(
                            x,
                            y + card_height_pt - mm(5),
                            card_width_pt,
                            mm(5),
                            settings["color_bar_top_color"],
                        )
                        logger.debug("Added top color bar")
                    if settings["color_bar_bottom"]:
                        draw_color_bar(
                            x,
                            y,
                            card_width_pt,
                            mm(5),
                            settings["color_bar_bottom_color"],
                        )
                        logger.debug("Added bottom color bar")

                    # Main text (front)
                    text = str(row.get(settings["front_column"], ""))
                    max_text_height = card_height_pt - mm(8)
                    if settings["color_bar_top"]:
                        max_text_height -= mm(5)
                    if settings["color_bar_bottom"]:
                        max_text_height -= mm(5)
                    max_lines = int(max_text_height // line_height)
                    lines, overflowed = wrap_text(
                        text,
                        font_name,
                        font_size,
                        card_width_pt - mm(8),
                        max_lines,
                        settings.get("truncate", False),
                    )
                    if overflowed and not settings.get("truncate", False):
                        logger.error(f"Front text overflow at row {data_idx+1}")
                        raise Exception(
                            f"Text does not fit on card at row {data_idx+1}. Enable 'Truncate' or edit your data."
                        )

                    c.setFont(font_name, font_size)
                    c.setFillColor(HexColor(text_color))
                    text_height = len(lines) * line_height
                    text_y = y + (card_height_pt - text_height) / 2  # Center vertically
                    for lidx, line in enumerate(lines):
                        c.drawCentredString(
                            x + card_width_pt / 2,
                            text_y + text_height - (lidx + 1) * line_height,
                            line,
                        )
                    logger.debug(
                        f"Drew {len(lines)} lines of front text, centered at y={text_y}"
                    )

            draw_cut_lines(
                c,
                page_size_pt[0],
                page_size_pt[1],
                offset_left_pt,
                offset_top_pt,
                card_width_pt,
                card_height_pt,
                cards_per_row,
                cards_per_col,
            )
            c.showPage()
            logger.debug(f"Completed front page {page + 1}")

            # --- Back page ---
            logger.debug(f"Generating back page {page + 1}")
            back_order = reorder_for_back(
                page_indices, cards_per_row, cards_per_col, flip_mode
            )
            for pos_on_page, data_idx in enumerate(back_order):
                card_row = pos_on_page // cards_per_row
                card_col = pos_on_page % cards_per_row
                x = offset_left_pt + card_col * card_width_pt
                y = page_size_pt[1] - offset_top_pt - (card_row + 1) * card_height_pt
                logger.debug(
                    f"Back card at pos {pos_on_page}, data_idx={data_idx}, x={x}, y={y}"
                )

                if data_idx is not None:
                    row = data.iloc[data_idx]
                    # Color bars
                    if settings["color_bar_top"]:
                        draw_color_bar(
                            x,
                            y + card_height_pt - mm(5),
                            card_width_pt,
                            mm(5),
                            settings["color_bar_top_color"],
                        )
                        logger.debug("Added top color bar")
                    if settings["color_bar_bottom"]:
                        draw_color_bar(
                            x,
                            y,
                            card_width_pt,
                            mm(5),
                            settings["color_bar_bottom_color"],
                        )
                        logger.debug("Added bottom color bar")

                    # Main text (back)
                    text = str(row.get(settings["back_column"], ""))
                    max_text_height = card_height_pt - mm(8)
                    if settings["color_bar_top"]:
                        max_text_height -= mm(5)
                    if settings["color_bar_bottom"]:
                        max_text_height -= mm(5)
                    max_lines = int(max_text_height // line_height)
                    lines, overflowed = wrap_text(
                        text,
                        font_name,
                        font_size,
                        card_width_pt - mm(8),
                        max_lines,
                        settings.get("truncate", False),
                    )
                    if overflowed and not settings.get("truncate", False):
                        logger.error(f"Back text overflow at row {data_idx+1}")
                        raise Exception(
                            f"Back text does not fit on card at row {data_idx+1}. Enable 'Truncate' or edit your data."
                        )

                    c.setFont(font_name, font_size)
                    c.setFillColor(HexColor(text_color))
                    text_height = len(lines) * line_height
                    text_y = y + (card_height_pt - text_height) / 2  # Center vertically
                    for lidx, line in enumerate(lines):
                        c.drawCentredString(
                            x + card_width_pt / 2,
                            text_y + text_height - (lidx + 1) * line_height,
                            line,
                        )
                    logger.debug(
                        f"Drew {len(lines)} lines of back text, centered at y={text_y}"
                    )
            draw_cut_lines(
                c,
                page_size_pt[0],
                page_size_pt[1],
                offset_left_pt,
                offset_top_pt,
                card_width_pt,
                card_height_pt,
                cards_per_row,
                cards_per_col,
            )
            c.showPage()
            logger.debug(f"Completed back page {page + 1}")

        c.save()
        logger.info(f"PDF saved successfully at {os.path.abspath(output_file)}")


if __name__ == "__main__":
    """
    Entry point for the Flashcard PDF Generator application.

    Initializes and runs the Tkinter GUI application.
    """
    logger.info("Starting Flashcard PDF Generator application")
    app = FlashcardApp()
    app.mainloop()
    logger.info("Application closed")
