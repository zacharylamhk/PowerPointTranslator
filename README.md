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

Installation
Clone the repository:
bash

git clone https://github.com/yourusername/powerpoint-translator.git
cd powerpoint-translator

Install dependencies:
bash

pip install -r requirements.txt

(Create a requirements.txt with python-pptx and deep-translator if desired)

Run the application:
bash

python PPTTranslatorApp.py

Usage
Launch the application:
bash

python PPTTranslatorApp.py

Use the GUI to:
Click "Browse" to select your input PowerPoint file

Specify the output file location (automatically suggested)

Choose source language (default: auto-detect)

Choose target language (default: English)

Click "Translate" to start the process

Watch the progress bar as the translation completes

Find your translated file at the specified output location

UI Screenshot
PowerPoint Translator UI
Supported Languages
The application uses deep-translator with Google Translate, supporting numerous languages. Refer to the "Language Reference" section in the GUI for a complete list of language codes and names.
Limitations
Requires an internet connection for translation

Subject to Google Translate's rate limits

Formatting may not be perfectly preserved in complex slides

Progress bar updates per slide, not per text element

Contributing
Fork the repository

Create a feature branch (git checkout -b feature/new-feature)

Commit your changes (git commit -m "Add new feature")

Push to the branch (git push origin feature/new-feature)

Create a Pull Request

License
This project is licensed under the MIT License - see the LICENSE file for details.
Acknowledgments
Built with python-pptx

Translation powered by deep-translator

GUI created using Tkinter

