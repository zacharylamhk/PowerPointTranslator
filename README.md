![PowerPoint Translator UI](screenshots/ui_screenshot.png)

# PowerPoint Translator

A Python application with a graphical user interface (GUI) to translate PowerPoint (.pptx) files from one language to another using the `deep-translator` library with Google Translate as the default engine.

## Features
- Translate text in slides, tables, and charts
- Select source and target languages from a comprehensive list
- Browse for input and output .pptx files
- Progress bar to track translation progress
- Support for auto-detection of source language
- User-friendly GUI built with Tkinter

## Prerequisites
- Python 3.6 or higher
- Required libraries:
  ```bash
  pip install python-pptx
  pip install deep-translator

  #Installation
  
  #Clone the repository:
  git clone https://github.com/zacharylamhk/PowerPointTranslator.git
  cd powerpoint-translator
  
  #Install dependencies:
  pip install -r requirements.txt


  #Run the Program
  python PPTTranslatorApp.py

## Supported Languages
The application uses deep-translator with Google Translate, supporting numerous languages. Refer to the "Language Reference" section in the GUI for a complete list of language codes and names.
Limitations
Requires an internet connection for translation

Subject to Google Translate's rate limits

Formatting may not be perfectly preserved in complex slides

Progress bar updates per slide, not per text element

## Contributing
Fork the repository

Create a feature branch (git checkout -b feature/new-feature)

Commit your changes (git commit -m "Add new feature")

Push to the branch (git push origin feature/new-feature)

Create a Pull Request

## License
This project is licensed under the MIT License - see the LICENSE file for details.
Acknowledgments
Built with python-pptx

Translation powered by deep-translator

GUI created using Tkinter

