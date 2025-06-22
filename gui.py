import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
from extract import process_pdf  # <- NEW: imported function from extract.py

class QuoteExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Insurance Quote Processor")
        self.root.geometry("640x460")
        self.root.configure(bg="#dbe5f1")

        self.quote_folder = ""
        self.output_folder = ""
        self.quote_count = 0

        self.setup_ui()

    def setup_ui(self):
        title = tk.Label(self.root, text="Insurance Quote Processor", font=("Helvetica", 16, "bold"), bg="#dbe5f1", fg="#1f3b57")
        title.pack(pady=10)

        self.info_label = tk.Label(self.root, text="Quote Folder: None | Output Folder: None | Quotes Found: 0", bg="#dbe5f1", fg="#1f3b57", anchor='w', justify="left")
        self.info_label.pack(padx=10, fill=tk.X)

        self.log_frame = tk.Frame(self.root, bg="#e7edf4", bd=2, relief=tk.GROOVE)
        self.log_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        log_label = tk.Label(self.log_frame, text="Log:", anchor='w', font=("Helvetica", 10, "bold"), bg="#e7edf4")
        log_label.pack(fill=tk.X)
        self.log_text = scrolledtext.ScrolledText(self.log_frame, height=10, state='disabled', wrap=tk.WORD, bg="#ffffff")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        button_frame = tk.Frame(self.root, bg="#dbe5f1")
        button_frame.pack(side=tk.BOTTOM, pady=10)

        tk.Button(button_frame, text="Select Quote Folder", command=self.select_quote_folder, width=20, bg="#4a90e2", fg="white").grid(row=0, column=0, padx=10, pady=5)
        tk.Button(button_frame, text="Select Output Folder", command=self.select_output_folder, width=20, bg="#4a90e2", fg="white").grid(row=0, column=1, padx=10, pady=5)
        tk.Button(button_frame, text="Read Quotes", command=self.read_quotes, width=20, bg="#357ABD", fg="white").grid(row=1, column=0, padx=10, pady=5)
        tk.Button(button_frame, text="Generate Word Doc", command=self.generate_doc, width=20, bg="#357ABD", fg="white").grid(row=1, column=1, padx=10, pady=5)

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def update_info_label(self):
        folder_name = os.path.basename(self.quote_folder) if self.quote_folder else "None"
        output_name = os.path.basename(self.output_folder) if self.output_folder else "None"
        self.info_label.config(text=f"Quote Folder: {folder_name} | Output Folder: {output_name} | Quotes Found: {self.quote_count}")

    def select_quote_folder(self):
        folder = filedialog.askdirectory(title="Select Folder with Quote PDFs")
        if folder:
            self.quote_folder = folder
            self.quote_count = len([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
            self.update_info_label()
            self.log(f"Selected quote folder: {folder} ({self.quote_count} PDFs found)")

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder = folder
            self.update_info_label()
            self.log(f"Selected output folder: {folder}")

    def read_quotes(self):
        if not self.quote_folder:
            messagebox.showerror("Error", "Please select a quote folder first.")
            return
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        self.log("Reading and processing all quotes...")

        pdf_files = [f for f in os.listdir(self.quote_folder) if f.lower().endswith('.pdf')]

        for pdf_file in pdf_files:
            try:
                input_path = os.path.join(self.quote_folder, pdf_file)
                output_path = os.path.join(self.output_folder, f"{os.path.splitext(pdf_file)[0]}_filled.json")

                self.log(f"Processing: {pdf_file}")
                process_pdf(input_path, output_path)
                self.log(f"✅ Finished: {pdf_file}")
            except Exception as e:
                self.log(f"❌ Error processing {pdf_file}: {e}")

    def generate_doc(self):
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder first.")
            return
        self.log("Generating Word document... (not implemented yet)")

if __name__ == '__main__':
    root = tk.Tk()
    app = QuoteExtractorGUI(root)
    root.mainloop()
