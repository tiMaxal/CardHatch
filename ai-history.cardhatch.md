
MIT License

Copyright (c) 2025 tiMaxal
# `CardHatch`
AI 'voding' [vibe-coding] history


## 20250709

### perplexity-ai
- prompt 1:
    - create py app, that is able to format a printable pdf from 2 spreadsheet columns as
        lists, to business cards [flashcards] with front and back [the 2 lists] aligned

- prompt 2:
    - provide the code as a gui that accepts values for excel/csv file location,
        column names [incl. extra for 'index' + 'notes'
   - if possible, being able to add configurable bars top + bottom for color-coding is
        desired], cards_per_row, page-size and margins [with defaults prefilled first run,
        and settings file to save current values for next operation];

- prompt 3:
    - apply grid columns like the following example for equi-spacing 'start' and 'exit'
        buttons at the bottom [and similarly, other entry fields in the main app window]
    **Configure columns for centering**
```
frame_buttons.columnconfigure(0, weight=1) # left spacer
frame_buttons.columnconfigure(1, weight=0) # left button [start]
frame_buttons.columnconfigure(2, weight=1) # center spacer
frame_buttons.columnconfigure(3, weight=0) # right button [exit]
frame_buttons.columnconfigure(4, weight=1) # right spacer
```

- prompt 4:
    - add color pickers for the top\bottom bar options

- prompt 5:
    - add a field to assign output field [default as origin of excel\csv input]
    - also wrap text content of input columns to fit output of card width and available lines
         [alert and stop if text overflow, ie out of room - allow 'truncate' checkbox]
    - format output with lines demarking cards, for separation [cutting]

## 20250710:

- prompt 6:
    - create output pdf as front/back pages alternating, for print both sides;
    - ensure 'back' cells will align with associated 'front' cells when page is flipped
      - requires:
        - radio button selection for 'flip long edge' or 'flip short edge'
        - if long, change order of each row
        - if short, invert whole page, but ensure cells will align
            vertically                                                                                                         

- prompt 7:
    - logic must assess how many cards fit on each page,
        then create a front and a back page for that many cards,
        then continue to the next amount that fit to the next page,
        and create those as front and back,
        forming a final output that alternates front and rear pages,
        with appropriate alignment

### grok-ai:

- prompt 8:
    - centre cards on page [both vertically n horizontally],
        to ensure alignment when flipped during print

- prompt 9:
    - centre the text in each card [both vertically and horizontally];
    - also provide options to choose Font size, family [and style, bold\italic\etc]
         by type\dropdown, and colour

- prompt 10:
    - provide the complete code, incl full docstrings,
        with good practice logging added to all stages
        [and include logging to file, in cwd]

- prompt 11:
    - remove all controls and references to 'index' and 'notes' columns
    - provide complete code, with full logging to file and docstrings

## 20250711

- prompt 12:
    - also function with .ods files
    - csv, excel and ods types all be shown at the same time in the picker, 
        instead of needing to use a drop-down
    - [and just show 'all' in the drop-down as well]

- prompt 13:
    - include a `Various amounts from 'qty' col` checkbox
        to apply diff quantity to certain cards [via column 'qty'],
        and a `Multiple` value input box to multiply *all* cards by a same amount

- prompt 14:
    - default empty 'qty' values to 1 [before multiplier applied],
        and notify user if any [and list which] 'qty' values are empty,
         not numeric or not positive integers

- prompt 15:
    - have CR [and LF] recognised, for formatted multi-line cells

- prompt 16:
    - create a checkbox to have front\back column names autofill, as read from first 2 column headers of file;
    - can add Arial choice?
    - limit 'multiple' input to a size of approx 4 digits
    - separate gui into sections;
      - file paths + column names
      - card layout
      - font details
      - colors
      - [manipulators]

- prompt 17:
    - swap color buttons with color indicators [put indicator swatch between description and button]
    - split 'Colors' section to two equal parts,
        Front on left and Back on right,
        duplicate the current 'overall' options for each,
        and add an option to each for background color;
    - similarly for 'Fonts'


