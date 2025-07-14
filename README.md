# `cardHatch`  – README
# Business and Flash card PDF Generator

Hatch your own Business Cards or Flashcards

## Overview

 This application generates printable PDF flashcards (business card size)
  from spreadsheet data in Excel (`.xlsx`, `.xls`), CSV, or ODS files. 

 Each card supports a front and back, with customizable formatting,
  color-coded bars, and precise alignment for duplex printing. 

 The app features a graphical user interface (GUI) for easy configuration,
  with settings saved for future use and comprehensive logging for troubleshooting.

## Features

- Import data from Excel (`.xlsx`, `.xls`), CSV, or ODS files
- Select columns for front and back of cards, with optional autofill from headers
- Optional `qty` column for variable card quantities, with a global multiplier (up to 4 digits)
- Adjustable cards per row, card size, page size, and margins
- Centered card and text alignment for professional output
- Configurable color bars (top and/or bottom) with color pickers for front and back
- Separate font size, family, style, text color, and background color for front and back
- Support for multi-line text with CR, LF, or CRLF handling
- Text wrapping with optional truncation for overflow
- Visual cut lines for easy card separation
- Duplex printing support with flip mode (long or short edge)
- Create new business card files via a popup dialog
- Settings saved in `cardhatch_settings.json` for reuse
- Comprehensive logging to `cardhatch.log` for debugging

## Installation

1. **Install Python 3** (if not already installed).
2. **Install required packages** by running:
   ```
   pip install pandas reportlab openpyxl odfpy
   ```
   [or]
   ```
   pip install -r requirements.txt
   ```
3. **Save the script** (`CardHatch.py`) to your computer.

## Usage

1. **Run the app:**
   ```
   python CardHatch.py
   ```
2. **In the GUI:**
   - Select your Excel, CSV, or ODS file.
   - Choose the output PDF location.
   - Specify column names for front and back (or enable autofill from headers).
   - Adjust card layout (cards per row, card size, page size, margins).
   - Customize front and back fonts (size, family, style) and colors (text, background, bars).
   - Enable color bars and use color pickers for customization.
   - Set flip mode for duplex printing (long or short edge).
   - Enable "Various amounts from 'qty' col" for per-card quantities and set a global multiplier.
   - Check "Truncate overflow text" to shorten long text with `...`.
   - Use "Create Business Card..." to make a new simple card file.
   - Click **Start** to generate the PDF, or **Exit** to close the app.

## File Format

The input file should have at least two columns for the front and back of the cards. An optional `qty` column can specify card quantities (non-numeric or invalid values default to 1).

| Front      | Back       | qty |
|------------|------------|-----|
| Question 1 | Answer 1   | 2   |
| Question 2 | Answer 2   | 1   |

## Settings

- **Cards per Row:** Number of cards across each page.
- **Card Size:** Width x Height in millimeters (e.g., `85.6x55.0` for business cards).
- **Page Size:** Width x Height in millimeters (e.g., `210x297` for A4).
- **Margins:** Top, Bottom, Left, Right in millimeters (comma-separated).
- **Font Details:** Font size, family (Helvetica, Times-Roman, Courier), style (Normal, Bold, Italic, BoldItalic) for front and back.
- **Colors:** Text, background, and optional top/bottom color bars for front and back.
- **Flip Mode:** Long or short edge for duplex printing alignment.
- **Truncate:** If enabled, text that doesn't fit is shortened with `...`.
- **Quantity:** Use `qty` column for per-card amounts and a global multiplier.

Settings are saved in `cardhatch_settings.json` and prefilled on next use.

## Troubleshooting

- **Missing Columns:** Verify column names match the file headers.
- **Text Overflow:** Enable "Truncate" or shorten content.
- **Output Not Created:** Check file paths and permissions.
- **Invalid Quantities:** Non-numeric or invalid `qty` values default to 1 with warnings.
- **Alignment Issues:** Ensure correct flip mode for your printer.
- **ODS Files:** Verify multi-line text renders correctly due to format variability.
- Check `cardhatch.log` for detailed error messages.

## License

MIT License

Copyright (c) 2025 tiMaxal

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.