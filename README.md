# Flashcard PDF Generator – README

## Overview

This application allows to create printable PDF flashcards (business card size) from two columns in an Excel or CSV file. Each card has a front and back, with options for color-coded bars, index, notes, and more. The app features a graphical user interface (GUI) for easy configuration and saves settings for future use.

## Features

- Import data from Excel (`.xlsx`, `.xls`) or CSV files
- Select which columns to use for the front and back of cards
- Optional columns for index and notes
- Adjustable cards per row, page size, and margins
- Add configurable color bars (top and/or bottom) for color-coding
- Color pickers for bar customization
- Output PDF file location selection
- Text wrapping and optional truncation if content overflows card space
- Visual cut lines for easy card separation
- Settings are saved for the next session

## Installation

1. **Install Python 3** (if not already installed).
2. **Install required packages** by running:
`pip install pandas reportlab openpyxl`
3. **Save the script** (`cardset.py`) to your computer.

## Usage

1. **Run the app:**
`python cardset.py`
2. **In the GUI:**
- Select your Excel/CSV file.
- Choose the output PDF location.
- Enter the column names for the front and back (and optional index/notes).
- Adjust settings for cards per row, page size, margins, and color bars as desired.
- Use the color pickers to set bar colors if enabled.
- Check "Truncate overflow text" if you want long text to be shortened.
- Click **Start** to generate the PDF, or **Exit** to close the app.

## File Format

 Excel/CSV file should have at least two columns for the front and back of the cards.
  Optional columns for index and notes can be included.

| Front      | Back       | Index | Notes      |
|------------|------------|-------|------------|
| Question 1 | Answer 1   | 1     | Example    |
| Question 2 | Answer 2   | 2     | Optional   |

## Settings

- **Cards per Row:** Number of cards across each page.
- **Page Size:** Width x Height in millimeters (e.g., `210x297` for A4).
- **Margins:** Top, Bottom, Left, Right in millimeters (comma-separated).
- **Color Bars:** Add colored bars to top and/or bottom of cards for organization.
- **Truncate:** If enabled, text that doesn't fit will be shortened with `...`.

Settings are saved in `flashcard_gui_settings.json` and prefilled on next use.

## Troubleshooting

- **Missing Columns:** Ensure your file has the correct column names.
- **Text Overflow:** Enable "Truncate" or shorten your card content.
- **Output Not Created:** Check file paths and permissions.

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

