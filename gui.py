import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
from extract import process_folder

class QuoteExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Insurance Quote Processor")
        self.root.geometry("640x500")
        self.root.configure(bg="#dbe5f1")  # Light blue/gray background

        self.quote_folder = ""
        self.output_folder = ""
        self.quote_count = 0
        self.broker_fee = 20  # Default to 20%

        self.setup_ui()

    def setup_ui(self):
        title = tk.Label(
            self.root, text="Insurance Quote Processor",
            font=("Helvetica", 16, "bold"),
            bg="#dbe5f1", fg="#1f3b57"
        )
        title.pack(pady=10)

        self.info_label = tk.Label(
            self.root, text="Quote Folder: None | Output Folder: None | Quotes Found: 0",
            bg="#dbe5f1", fg="#1f3b57", anchor='w', justify="left"
        )
        self.info_label.pack(padx=10, fill=tk.X)

        # --- Broker Fee Section (slider + input box) ---
        fee_frame = tk.Frame(self.root, bg="#dbe5f1")
        fee_frame.pack(padx=10, pady=(0, 8), fill=tk.X)
        tk.Label(
            fee_frame, text="Broker Fee (%):",
            bg="#dbe5f1", fg="#1f3b57", font=("Helvetica", 11)
        ).pack(side=tk.LEFT)

        self.broker_fee_var = tk.IntVar(value=self.broker_fee)

        # Entry box for fee
        self.fee_entry = tk.Entry(fee_frame, width=5, font=("Helvetica", 11), justify='center')
        self.fee_entry.pack(side=tk.LEFT, padx=(5, 3))
        self.fee_entry.insert(0, str(self.broker_fee))
        self.fee_entry.bind('<FocusOut>', self.entry_broker_fee_update)
        self.fee_entry.bind('<Return>', self.entry_broker_fee_update)

        # Slider for fee
        self.fee_slider = tk.Scale(
            fee_frame,
            from_=0, to=100,
            orient=tk.HORIZONTAL,
            variable=self.broker_fee_var,
            command=self.slider_broker_fee_update,
            showvalue=0,
            resolution=1,
            length=200,
            bg="#dbe5f1",
            troughcolor="#b0c4de",
            highlightthickness=0
        )
        self.fee_slider.pack(side=tk.LEFT, padx=(5,5))

        self.fee_label = tk.Label(
            fee_frame, text="%",
            bg="#dbe5f1", fg="#357ABD", font=("Helvetica", 11, "bold")
        )
        self.fee_label.pack(side=tk.LEFT)

        # --- Log Frame ---
        self.log_frame = tk.Frame(self.root, bg="#e7edf4", bd=2, relief=tk.GROOVE)
        self.log_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        log_label = tk.Label(
            self.log_frame, text="Log:", anchor='w',
            font=("Helvetica", 10, "bold"), bg="#e7edf4"
        )
        log_label.pack(fill=tk.X)
        self.log_text = scrolledtext.ScrolledText(
            self.log_frame, height=10, state='disabled',
            wrap=tk.WORD, bg="#ffffff"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # --- Button Frame ---
        button_frame = tk.Frame(self.root, bg="#dbe5f1")
        button_frame.pack(side=tk.BOTTOM, pady=10)

        tk.Button(
            button_frame, text="Select Quote Folder",
            command=self.select_quote_folder, width=20,
            bg="#4a90e2", fg="white"
        ).grid(row=0, column=0, padx=10, pady=5)
        tk.Button(
            button_frame, text="Select Output Folder",
            command=self.select_output_folder, width=20,
            bg="#4a90e2", fg="white"
        ).grid(row=0, column=1, padx=10, pady=5)
        tk.Button(
            button_frame, text="Read Quotes",
            command=self.read_quotes, width=20,
            bg="#357ABD", fg="white"
        ).grid(row=1, column=0, padx=10, pady=5)
        tk.Button(
            button_frame, text="Generate Word Doc",
            command=self.generate_doc, width=20,
            bg="#357ABD", fg="white"
        ).grid(row=1, column=1, padx=10, pady=5)

    # --- Broker Fee Sync Logic ---
    def slider_broker_fee_update(self, val=None):
        self.broker_fee = self.broker_fee_var.get()
        self.fee_entry.delete(0, tk.END)
        self.fee_entry.insert(0, str(self.broker_fee))

    def entry_broker_fee_update(self, event=None):
        try:
            fee = int(self.fee_entry.get())
            if fee < 0:
                fee = 0
            elif fee > 100:
                fee = 100
        except ValueError:
            fee = 20  # Reset to default on invalid
        self.broker_fee = fee
        self.broker_fee_var.set(fee)
        self.fee_entry.delete(0, tk.END)
        self.fee_entry.insert(0, str(fee))

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
            messagebox.showerror("Error", "Please select an output folder first.")
            return

        self.log(f"Reading all quotes in folder: {self.quote_folder} ...")
        output_path = os.path.join(self.output_folder, "combined_quotes.json")

        try:
            # Pass broker_fee if you want it in your extract/process code
            process_folder(self.quote_folder, output_path)
            self.log(f"Extraction complete. JSON saved to: {output_path}")
            self.log(f"Broker Fee used: {self.broker_fee}%")
        except Exception as e:
            self.log(f"Error during extraction: {e}")

    def generate_doc(self):
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder first.")
            return
        self.log(f"Generate Word Doc clicked. (Functionality not yet implemented.)")

if __name__ == '__main__':
    root = tk.Tk()
    app = QuoteExtractorGUI(root)
    root.mainloop()
