'''
MIT License

Copyright (c) 2025 tiMaxal


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

'''

'''
Install Required Packages:
 bash
pip install pandas reportlab openpyxl

'''

import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import json
import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.pdfbase.pdfmetrics import stringWidth

# ---------- Settings Management ----------
DEFAULT_SETTINGS = {
    "file_path": "",
    "output_file": "",
    "front_column": "Front",
    "back_column": "Back",
    "index_column": "Index",
    "notes_column": "Notes",
    "cards_per_row": 3,
    "page_size": "210x297",  # A4 in mm
    "margins": "10,10,10,10",  # top,bottom,left,right in mm
    "color_bar_top": False,
    "color_bar_bottom": False,
    "color_bar_top_color": "#FF0000",
    "color_bar_bottom_color": "#0000FF",
    "truncate": False
}
SETTINGS_FILE = "flashcard_gui_settings.json"

if os.path.exists(SETTINGS_FILE):
    with open(SETTINGS_FILE, "r") as f:
        settings = json.load(f)
else:
    settings = DEFAULT_SETTINGS.copy()

# ---------- Utility Functions ----------
def mm(val):
    return float(val) * 2.83465

def wrap_text(text, font_name, font_size, max_width_pt, max_lines, truncate=False):
    """Wraps text to fit in max_width_pt and max_lines. Returns (lines, overflowed)"""
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
                    lines[-1] = lines[-1][:int(max_width_pt/font_size)] + "..."
                    return lines, False
                else:
                    return lines, True
    if current_line:
        lines.append(current_line)
    if len(lines) > max_lines:
        if truncate:
            lines = lines[:max_lines]
            lines[-1] = lines[-1][:int(max_width_pt/font_size)] + "..."
            return lines, False
        else:
            return lines[:max_lines], True
    return lines, False

def draw_cut_lines(c, page_w_pt, page_h_pt, margin_left_pt, margin_top_pt, card_width_pt, card_height_pt, cards_per_row, cards_per_col):
    c.setStrokeColor(HexColor("#888888"))
    c.setLineWidth(0.5)
    # Vertical lines
    for i in range(cards_per_row + 1):
        x = margin_left_pt + i * card_width_pt
        c.line(x, margin_top_pt, x, page_h_pt - margin_top_pt)
    # Horizontal lines
    for j in range(cards_per_col + 1):
        y = margin_top_pt + j * card_height_pt
        c.line(margin_left_pt, y, page_w_pt - margin_left_pt, y)

# ---------- Main Application ----------
class FlashcardApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Flashcard PDF Generator")
        self.geometry("700x530")

        frame_main = tk.Frame(self)
        frame_main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # File selection
        tk.Label(frame_main, text="Excel/CSV File Path:").grid(row=0, column=0, sticky=tk.W)
        self.entry_file = tk.Entry(frame_main)
        self.entry_file.grid(row=0, column=1, sticky=tk.EW)
        self.entry_file.insert(0, settings.get("file_path", ""))
        btn_browse = tk.Button(frame_main, text="Browse...", command=self.browse_file)
        btn_browse.grid(row=0, column=2)

        # Output file
        tk.Label(frame_main, text="Output PDF File:").grid(row=1, column=0, sticky=tk.W)
        self.entry_output = tk.Entry(frame_main)
        self.entry_output.grid(row=1, column=1, sticky=tk.EW)
        default_output = os.path.splitext(settings.get("file_path", ""))[0] + ".pdf" if settings.get("file_path", "") else "flashcards.pdf"
        self.entry_output.insert(0, settings.get("output_file", default_output))
        btn_output_browse = tk.Button(frame_main, text="Browse...", command=self.browse_output_file)
        btn_output_browse.grid(row=1, column=2)

        # Column names
        tk.Label(frame_main, text="Front Column Name:").grid(row=2, column=0, sticky=tk.W)
        self.entry_front = tk.Entry(frame_main)
        self.entry_front.grid(row=2, column=1, sticky=tk.EW)
        self.entry_front.insert(0, settings.get("front_column", "Front"))

        tk.Label(frame_main, text="Back Column Name:").grid(row=3, column=0, sticky=tk.W)
        self.entry_back = tk.Entry(frame_main)
        self.entry_back.grid(row=3, column=1, sticky=tk.EW)
        self.entry_back.insert(0, settings.get("back_column", "Back"))

        tk.Label(frame_main, text="Index Column Name (optional):").grid(row=4, column=0, sticky=tk.W)
        self.entry_index = tk.Entry(frame_main)
        self.entry_index.grid(row=4, column=1, sticky=tk.EW)
        self.entry_index.insert(0, settings.get("index_column", "Index"))

        tk.Label(frame_main, text="Notes Column Name (optional):").grid(row=5, column=0, sticky=tk.W)
        self.entry_notes = tk.Entry(frame_main)
        self.entry_notes.grid(row=5, column=1, sticky=tk.EW)
        self.entry_notes.insert(0, settings.get("notes_column", "Notes"))

        # Cards per row
        tk.Label(frame_main, text="Cards per Row:").grid(row=6, column=0, sticky=tk.W)
        self.entry_cards_per_row = tk.Entry(frame_main)
        self.entry_cards_per_row.grid(row=6, column=1, sticky=tk.EW)
        self.entry_cards_per_row.insert(0, str(settings.get("cards_per_row", 3)))

        # Page size
        tk.Label(frame_main, text="Page Size (WxH mm):").grid(row=7, column=0, sticky=tk.W)
        self.entry_page_size = tk.Entry(frame_main)
        self.entry_page_size.grid(row=7, column=1, sticky=tk.EW)
        self.entry_page_size.insert(0, settings.get("page_size", "210x297"))

        # Margins
        tk.Label(frame_main, text="Margins (Top,Bottom,Left,Right mm):").grid(row=8, column=0, sticky=tk.W)
        self.entry_margins = tk.Entry(frame_main)
        self.entry_margins.grid(row=8, column=1, sticky=tk.EW)
        self.entry_margins.insert(0, settings.get("margins", "10,10,10,10"))

        # Color bars with pickers
        self.color_bar_top_var = tk.BooleanVar(value=settings.get("color_bar_top", False))
        self.color_bar_bottom_var = tk.BooleanVar(value=settings.get("color_bar_bottom", False))
        self.color_bar_top_color = tk.StringVar(value=settings.get("color_bar_top_color", "#FF0000"))
        self.color_bar_bottom_color = tk.StringVar(value=settings.get("color_bar_bottom_color", "#0000FF"))

        tk.Checkbutton(frame_main, text="Add Color Bar Top", variable=self.color_bar_top_var).grid(row=9, column=0, sticky=tk.W)
        btn_top_color = tk.Button(frame_main, text="Pick Top Color", command=self.pick_top_color)
        btn_top_color.grid(row=9, column=1, sticky=tk.W)
        self.lbl_top_color = tk.Label(frame_main, textvariable=self.color_bar_top_color, bg=self.color_bar_top_color.get(), width=10)
        self.lbl_top_color.grid(row=9, column=2, sticky=tk.W)

        tk.Checkbutton(frame_main, text="Add Color Bar Bottom", variable=self.color_bar_bottom_var).grid(row=10, column=0, sticky=tk.W)
        btn_bottom_color = tk.Button(frame_main, text="Pick Bottom Color", command=self.pick_bottom_color)
        btn_bottom_color.grid(row=10, column=1, sticky=tk.W)
        self.lbl_bottom_color = tk.Label(frame_main, textvariable=self.color_bar_bottom_color, bg=self.color_bar_bottom_color.get(), width=10)
        self.lbl_bottom_color.grid(row=10, column=2, sticky=tk.W)

        # Truncate checkbox
        self.truncate_var = tk.BooleanVar(value=settings.get("truncate", False))
        tk.Checkbutton(frame_main, text="Truncate overflow text", variable=self.truncate_var).grid(row=11, column=0, sticky=tk.W)

        frame_main.columnconfigure(0, weight=0)
        frame_main.columnconfigure(1, weight=1)
        frame_main.columnconfigure(2, weight=0)

        # Buttons frame
        frame_buttons = tk.Frame(self)
        frame_buttons.pack(fill=tk.X, padx=10, pady=10)
        frame_buttons.columnconfigure(0, weight=1)
        frame_buttons.columnconfigure(1, weight=0)
        frame_buttons.columnconfigure(2, weight=1)
        frame_buttons.columnconfigure(3, weight=0)
        frame_buttons.columnconfigure(4, weight=1)

        btn_start = tk.Button(frame_buttons, text="Start", command=self.start_process)
        btn_start.grid(row=0, column=1, sticky=tk.EW)
        btn_exit = tk.Button(frame_buttons, text="Exit", command=self.quit)
        btn_exit.grid(row=0, column=3, sticky=tk.EW)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv"), ("All files", "*")])
        if file_path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, file_path)
            # Set default output file
            default_output = os.path.splitext(file_path)[0] + ".pdf"
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, default_output)

    def browse_output_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, file_path)

    def pick_top_color(self):
        color = colorchooser.askcolor(title="Pick Top Bar Color")[1]
        if color:
            self.color_bar_top_color.set(color)
            self.lbl_top_color.config(bg=color)

    def pick_bottom_color(self):
        color = colorchooser.askcolor(title="Pick Bottom Bar Color")[1]
        if color:
            self.color_bar_bottom_color.set(color)
            self.lbl_bottom_color.config(bg=color)

    def start_process(self):
        # Save current settings
        settings["file_path"] = self.entry_file.get()
        settings["output_file"] = self.entry_output.get()
        settings["front_column"] = self.entry_front.get()
        settings["back_column"] = self.entry_back.get()
        settings["index_column"] = self.entry_index.get()
        settings["notes_column"] = self.entry_notes.get()
        settings["cards_per_row"] = int(self.entry_cards_per_row.get())
        settings["page_size"] = self.entry_page_size.get()
        settings["margins"] = self.entry_margins.get()
        settings["color_bar_top"] = self.color_bar_top_var.get()
        settings["color_bar_bottom"] = self.color_bar_bottom_var.get()
        settings["color_bar_top_color"] = self.color_bar_top_color.get()
        settings["color_bar_bottom_color"] = self.color_bar_bottom_color.get()
        settings["truncate"] = self.truncate_var.get()

        with open(SETTINGS_FILE, "w") as f:
            json.dump(settings, f, indent=4)

        file_path = settings["file_path"]
        if not file_path:
            messagebox.showerror("Error", "Please select an input file.")
            return

        try:
            if file_path.lower().endswith(".csv"):
                data = pd.read_csv(file_path)
            else:
                data = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")
            return

        # Validate columns
        missing = []
        for col in [settings["front_column"], settings["back_column"]]:
            if col not in data.columns:
                missing.append(col)
        if missing:
            messagebox.showerror("Error", f"Missing columns in data: {', '.join(missing)}")
            return

        try:
            self.generate_flashcard_pdf(data, settings)
            messagebox.showinfo("Success", f"PDF generated as {settings['output_file']}")
        except Exception as e:
            messagebox.showerror("Error", f"PDF generation failed: {e}")

    def generate_flashcard_pdf(self, data, settings):
        # Parse page size and margins
        page_w, page_h = map(float, settings["page_size"].split("x"))
        margins = list(map(float, settings["margins"].split(",")))
        margin_top, margin_bottom, margin_left, margin_right = margins

        cards_per_row = int(settings["cards_per_row"])
        card_width = (page_w - margin_left - margin_right) / cards_per_row
        card_height = 55  # Standard business card height in mm

        page_size_pt = (mm(page_w), mm(page_h))
        card_width_pt = mm(card_width)
        card_height_pt = mm(card_height)
        margin_top_pt, margin_left_pt = mm(margin_top), mm(margin_left)
        margin_bottom_pt, margin_right_pt = mm(margin_bottom), mm(margin_right)

        # Calculate cards per column
        cards_per_col = int((page_h - margin_top - margin_bottom) // card_height)

        output_file = settings.get("output_file", "flashcards.pdf")
        c = canvas.Canvas(output_file, pagesize=page_size_pt)

        # Font settings
        font_name = "Helvetica-Bold"
        font_size = 12
        line_height = font_size * 1.2

        # Helper for color bars
        def draw_color_bar(x, y, width, height, color_hex):
            c.setFillColor(HexColor(color_hex))
            c.rect(x, y, width, height, fill=1, stroke=0)

        # Draw cards front
        idx = 0
        for i, row in data.iterrows():
            card_row = (idx // cards_per_row) % cards_per_col
            card_col = idx % cards_per_row
            if idx > 0 and idx % (cards_per_row * cards_per_col) == 0:
                # Draw cut lines before new page
                draw_cut_lines(c, page_size_pt[0], page_size_pt[1], margin_left_pt, margin_top_pt, card_width_pt, card_height_pt, cards_per_row, cards_per_col)
                c.showPage()
            x = margin_left_pt + card_col * card_width_pt
            y = page_size_pt[1] - margin_top_pt - (card_row + 1) * card_height_pt

            # Color bars
            if settings["color_bar_top"]:
                draw_color_bar(x, y + card_height_pt - mm(5), card_width_pt, mm(5), settings["color_bar_top_color"])
            if settings["color_bar_bottom"]:
                draw_color_bar(x, y, card_width_pt, mm(5), settings["color_bar_bottom_color"])

            # Main text (front)
            text = str(row.get(settings["front_column"], ""))
            max_lines = int((card_height_pt - 2*mm(8)) // line_height)
            lines, overflowed = wrap_text(text, font_name, font_size, card_width_pt - mm(8), max_lines, settings.get("truncate", False))
            if overflowed and not settings.get("truncate", False):
                raise Exception(f"Text does not fit on card at row {i+1}. Enable 'Truncate' or edit your data.")

            c.setFont(font_name, font_size)
            c.setFillColor(HexColor("#000000"))
            for lidx, line in enumerate(lines):
                c.drawCentredString(x + card_width_pt/2, y + card_height_pt - mm(10) - lidx*line_height, line)

            # Index and notes
            if settings["index_column"] and settings["index_column"] in row:
                c.setFont("Helvetica", 8)
                c.drawString(x + mm(2), y + card_height_pt - mm(8), str(row.get(settings["index_column"], "")))
            if settings["notes_column"] and settings["notes_column"] in row:
                c.setFont("Helvetica-Oblique", 8)
                c.drawString(x + mm(2), y + mm(2), str(row.get(settings["notes_column"], "")))

            idx += 1

        # Draw final cut lines for last page
        draw_cut_lines(c, page_size_pt[0], page_size_pt[1], margin_left_pt, margin_top_pt, card_width_pt, card_height_pt, cards_per_row, cards_per_col)
        c.showPage()

        # Draw backs
        idx = 0
        for i, row in data.iterrows():
            card_row = (idx // cards_per_row) % cards_per_col
            card_col = idx % cards_per_row
            if idx > 0 and idx % (cards_per_row * cards_per_col) == 0:
                draw_cut_lines(c, page_size_pt[0], page_size_pt[1], margin_left_pt, margin_top_pt, card_width_pt, card_height_pt, cards_per_row, cards_per_col)
                c.showPage()
            x = margin_left_pt + card_col * card_width_pt
            y = page_size_pt[1] - margin_top_pt - (card_row + 1) * card_height_pt

            # Color bars
            if settings["color_bar_top"]:
                draw_color_bar(x, y + card_height_pt - mm(5), card_width_pt, mm(5), settings["color_bar_top_color"])
            if settings["color_bar_bottom"]:
                draw_color_bar(x, y, card_width_pt, mm(5), settings["color_bar_bottom_color"])

            # Main text (back)
            text = str(row.get(settings["back_column"], ""))
            max_lines = int((card_height_pt - 2*mm(8)) // line_height)
            lines, overflowed = wrap_text(text, font_name, font_size, card_width_pt - mm(8), max_lines, settings.get("truncate", False))
            if overflowed and not settings.get("truncate", False):
                raise Exception(f"Back text does not fit on card at row {i+1}. Enable 'Truncate' or edit your data.")

            c.setFont(font_name, font_size)
            c.setFillColor(HexColor("#000000"))
            for lidx, line in enumerate(lines):
                c.drawCentredString(x + card_width_pt/2, y + card_height_pt - mm(10) - lidx*line_height, line)

            idx += 1

        draw_cut_lines(c, page_size_pt[0], page_size_pt[1], margin_left_pt, margin_top_pt, card_width_pt, card_height_pt, cards_per_row, cards_per_col)
        c.save()

if __name__ == "__main__":
    app = FlashcardApp()
    app.mainloop()
