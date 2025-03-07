import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pptx import Presentation
from deep_translator import GoogleTranslator
import os

def translate_text(text, source_language, target_language):
    """Translate text from source language to target language using deep-translator"""
    try:
        if text.strip():
            translator = GoogleTranslator(source=source_language, target=target_language)
            translated = translator.translate(text)
            return translated
        return text
    except Exception as e:
        print(f"Translation error: {e}")
        return text

def translate_ppt(input_path, output_path, source_lang, target_lang):
    """Translate PowerPoint file"""
    try:
        prs = Presentation(input_path)
        
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    shape.text = translate_text(shape.text, source_lang, target_lang)
                
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text:
                                cell.text = translate_text(cell.text, source_lang, target_lang)
                
                if shape.has_chart:
                    chart = shape.chart
                    for series in chart.series:
                        if series.name:
                            series.name = translate_text(series.name, source_lang, target_lang)
                    if chart.has_title:
                        chart.chart_title.text_frame.text = translate_text(
                            chart.chart_title.text_frame.text, source_lang, target_lang
                        )

        prs.save(output_path)
        messagebox.showinfo("Success", f"Translation completed!\nSaved as: {output_path}")
        
    except Exception as e:
        messagebox.showerror("Error", f"Error processing presentation: {e}")

class PPTTranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Translator")
        self.root.geometry("500x350")

        # Variables
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.source_lang = tk.StringVar(value='auto')  # Default to auto-detect
        self.target_lang = tk.StringVar(value='en')    # Default to English

        # Get supported languages from deep-translator
        self.languages = GoogleTranslator().get_supported_languages(as_dict=True)

        # GUI Elements
        # Input file
        tk.Label(root, text="Source PPTX:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(root, textvariable=self.input_path, width=40).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(root, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=5, pady=5)

        # Output file
        tk.Label(root, text="Output PPTX:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(root, textvariable=self.output_path, width=40).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(root, text="Browse", command=self.browse_output).grid(row=1, column=2, padx=5, pady=5)

        # Source language
        tk.Label(root, text="Source Language:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        source_combo = ttk.Combobox(root, textvariable=self.source_lang, width=37)
        source_combo['values'] = ['auto'] + list(self.languages.keys())
        source_combo.grid(row=2, column=1, padx=5, pady=5)

        # Target language
        tk.Label(root, text="Target Language:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        target_combo = ttk.Combobox(root, textvariable=self.target_lang, width=37)
        target_combo['values'] = list(self.languages.keys())
        target_combo.grid(row=3, column=1, padx=5, pady=5)

        # Language names display
        tk.Label(root, text="Language Reference:").grid(row=4, column=0, padx=5, pady=5, sticky="ne")
        lang_display = tk.Text(root, height=8, width=40)
        lang_display.grid(row=4, column=1, padx=5, pady=5, columnspan=2)
        for code, lang in sorted(self.languages.items(), key=lambda x: x[1]):
            lang_display.insert(tk.END, f"{code}: {lang}\n")
        lang_display.config(state='disabled')

        # Translate button
        tk.Button(root, text="Translate", command=self.translate).grid(row=5, column=1, pady=20)

    def browse_input(self):
        """Open file dialog for input file"""
        filename = filedialog.askopenfilename(
            filetypes=[("PowerPoint files", "*.pptx *.ppt"), ("All files", "*.*")]
        )
        if filename:
            self.input_path.set(filename)
            output_suggestion = f"translated_{os.path.basename(filename)}"
            self.output_path.set(os.path.join(os.path.dirname(filename), output_suggestion))

    def browse_output(self):
        """Open file dialog for output file"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")]
        )
        if filename:
            self.output_path.set(filename)

    def translate(self):
        """Handle translation process"""
        if not self.input_path.get():
            messagebox.showerror("Error", "Please select an input file!")
            return
        if not self.output_path.get():
            messagebox.showerror("Error", "Please specify an output file!")
            return
        if not os.path.exists(self.input_path.get()):
            messagebox.showerror("Error", "Input file not found!")
            return
        
        translate_ppt(
            self.input_path.get(),
            self.output_path.get(),
            self.source_lang.get(),
            self.target_lang.get()
        )

def main():
    root = tk.Tk()
    app = PPTTranslatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()