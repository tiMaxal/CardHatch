- popup to create business card file
    - option to load previous file
    - 2 text input fields, for front and back 
    - field for filename 
    - save in csv format
    - [later, other formats?]
    - 'cancel' and 'save' buttons 

- option to have text orientated left, central or right

[ai-perplexity 20250715
Potential Improvements]
- Dependency Management:
    - Consider adding a requirements.txt file for easier installation.
- Font Handling:
    - If more fonts are desired in the future, consider dynamic font discovery or allowing user-supplied fonts.
- Internationalization [i18n]:
    - Add support for Unicode and right-to-left languages if needed.
- Accessibility:
    -Add keyboard shortcuts and improve screen reader support for better accessibility.
- Unit Testing:
    -Add unit tests for utility functions (e.g., wrap_text, reorder_for_back) to ensure reliability.
- Error Handling:
    - Consider more granular error messages, especially for file parsing and PDF generation.
- Performance:
    - For very large input files, consider streaming or chunked processing to avoid memory issues.
- Documentation:
    - [modify] README file with usage instructions, screenshots, and troubleshooting tips.

Summary Table: Docstring and Code Review
Area	|Recommendation
Main module docstring	|Update as above to reflect all current features and requirements
Class/method docstrings	|Expand/clarify for FlashcardApp, generate_flashcard_pdf, utility methods
Logging section	|Add a docstring noting dual logging (console and file)
Code layout	|Well-structured; logical sectioning and modular GUI
Potential improvements	|requirements.txt, more fonts, i18n, accessibility, tests, better error handling
- By updating the docstrings as suggested and considering the potential improvements, the codebase will be more maintainable, user-friendly, and robust for future development and user support.
