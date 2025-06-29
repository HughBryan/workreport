import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
from extract import process_folder
from report_generator import load_json, generate_report, resource_path

class QuoteExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Insurance Quote Processor")
        self.root.geometry("720x600")
        self.root.configure(bg="#dbe5f1")

        self.quote_folder = ""
        self.output_folder = ""
        self.quote_count = 0
        self.broker_fee = 20  # Default to 20%
        self.commission = 20  # Default to 10%
        self.template_path = None  # Path to the selected template

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

        self.template_label = tk.Label(
            self.root, text="Template: Default (report_template.docx)",
            bg="#dbe5f1", fg="#1f3b57", anchor='w', justify="left"
        )
        self.template_label.pack(padx=10, fill=tk.X)

        # --- Configuration Section ---
        config_frame = tk.LabelFrame(self.root, text="Pricing Configuration", bg="#dbe5f1", fg="#1f3b57",
                                     font=("Helvetica", 11, "bold"))
        config_frame.pack(padx=10, pady=(10, 5), fill=tk.X)

        # Row 1: Broker Fee and Commission
        row1 = tk.Frame(config_frame, bg="#dbe5f1")
        row1.pack(fill=tk.X, pady=5)

        tk.Label(row1, text="Commission (%):", bg="#dbe5f1", font=("Helvetica", 11)).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.commission_var = tk.IntVar(value=self.commission)
        self.comm_entry = tk.Entry(row1, width=5, font=("Helvetica", 11), justify='center')
        self.comm_entry.grid(row=0, column=1, padx=5)
        self.comm_entry.insert(0, str(self.commission))
        self.comm_entry.bind('<FocusOut>', self.entry_commission_update)
        self.comm_entry.bind('<Return>', self.entry_commission_update)
        self.comm_slider = tk.Scale(row1, from_=0, to=100, orient=tk.HORIZONTAL,
                                    variable=self.commission_var, command=self.slider_commission_update,
                                    showvalue=0, resolution=1, length=100,
                                    bg="#dbe5f1", troughcolor="#b0c4de", highlightthickness=0)
        self.comm_slider.grid(row=0, column=2, padx=5)

        tk.Label(row1, text="Broker Fee:", bg="#dbe5f1", font=("Helvetica", 11)).grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.broker_fee_var = tk.IntVar(value=self.broker_fee)
        self.fee_entry = tk.Entry(row1, width=5, font=("Helvetica", 11), justify='center')
        self.fee_entry.grid(row=1, column=1, padx=5)
        self.fee_entry.insert(0, str(self.broker_fee))
        self.fee_entry.bind('<FocusOut>', self.entry_broker_fee_update)
        self.fee_entry.bind('<Return>', self.entry_broker_fee_update)
        self.fee_slider = tk.Scale(row1, from_=0, to=100, orient=tk.HORIZONTAL,
                                   variable=self.broker_fee_var, command=self.slider_broker_fee_update,
                                   showvalue=0, resolution=1, length=100,
                                   bg="#dbe5f1", troughcolor="#b0c4de", highlightthickness=0)
        self.fee_slider.grid(row=1, column=2, padx=5)

        self.use_fixed_fee_var = tk.BooleanVar(value=False)
        self.fixed_fee_check = tk.Checkbutton(row1, text="Use Fixed $ Fee", variable=self.use_fixed_fee_var,
                                              command=self.toggle_fixed_fee, bg="#dbe5f1", font=("Helvetica", 10))
        self.fixed_fee_check.grid(row=1, column=3, padx=10, sticky="w")

        self.fixed_fee_var = tk.DoubleVar(value=0)
        self.fixed_fee_entry = tk.Entry(row1, width=7, font=("Helvetica", 11), justify='center',
                                        textvariable=self.fixed_fee_var, state='disabled')
        self.fixed_fee_entry.grid(row=1, column=4, padx=5)

        # Row 2: Associate Split
        row2 = tk.Frame(config_frame, bg="#dbe5f1")
        row2.pack(fill=tk.X, pady=5)

        tk.Label(row2, text="Associate Split (%):", bg="#dbe5f1", font=("Helvetica", 11)).pack(side=tk.LEFT)

        self.associate_split_var = tk.DoubleVar(value=60.00)
        self.broker_split_var = tk.DoubleVar(value=40.00)

        self.assoc_entry = tk.Entry(row2, width=7, font=("Helvetica", 11), justify='center')
        self.assoc_entry.pack(side=tk.LEFT, padx=(5, 3))
        self.assoc_entry.insert(0, "60.00")
        self.assoc_entry.bind('<FocusOut>', self.entry_associate_split_update)
        self.assoc_entry.bind('<Return>', self.entry_associate_split_update)

        self.assoc_slider = tk.Scale(row2, from_=0, to=100, orient=tk.HORIZONTAL,
                                    resolution=0.1,
                                    variable=self.associate_split_var, command=self.slider_associate_split_update,
                                    showvalue=0, length=100,
                                    bg="#dbe5f1", troughcolor="#b0c4de", highlightthickness=0)
        self.assoc_slider.pack(side=tk.LEFT, padx=(5, 8))

        self.broker_entry = tk.Entry(row2, width=7, font=("Helvetica", 11), justify='center')
        self.broker_entry.pack(side=tk.LEFT, padx=(3, 5))
        self.broker_entry.insert(0, "40.00")
        self.broker_entry.configure(state='readonly')

        tk.Label(row2, text=": Broker Share (%)", bg="#dbe5f1", font=("Helvetica", 11)).pack(side=tk.LEFT)

        # Row 3: Strata Manager
        row3 = tk.Frame(config_frame, bg="#dbe5f1")
        row3.pack(fill=tk.X, pady=5)
        self.strata_checkbox_var = tk.BooleanVar(value=True)
        self.strata_checkbox = tk.Checkbutton(row3, text="Strata Manager:", variable=self.strata_checkbox_var,
                                              bg="#dbe5f1", font=("Helvetica", 11), command=self.toggle_strata_entry)
        self.strata_checkbox.pack(side=tk.LEFT)
        self.strata_manager_var = tk.StringVar()
        self.strata_entry = tk.Entry(row3, width=30, font=("Helvetica", 11), textvariable=self.strata_manager_var)
        self.strata_entry.pack(side=tk.LEFT, padx=(5, 10))

        # Row 4: Longitude Option (radio buttons)
        row4 = tk.Frame(config_frame, bg="#dbe5f1")
        row4.pack(fill=tk.X, pady=5)
        self.longitude_option_var = tk.StringVar(value="current")
        tk.Label(row4, text="Longitude Quotation Basis:", bg="#dbe5f1", font=("Helvetica", 11)).pack(side=tk.LEFT, padx=(0, 8))
        tk.Radiobutton(
            row4,
            text="Current Option",
            variable=self.longitude_option_var,
            value="current",
            bg="#dbe5f1",
            font=("Helvetica", 11)
        ).pack(side=tk.LEFT)
        tk.Radiobutton(
            row4,
            text="Suggested Option",
            variable=self.longitude_option_var,
            value="suggested",
            bg="#dbe5f1",
            font=("Helvetica", 11)
        ).pack(side=tk.LEFT, padx=(10, 0))

        # --- Log Frame ---
        self.log_frame = tk.Frame(self.root, bg="#e7edf4", bd=2, relief=tk.GROOVE)
        self.log_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        log_label = tk.Label(self.log_frame, text="Log:", anchor='w', font=("Helvetica", 10, "bold"), bg="#e7edf4")
        log_label.pack(fill=tk.X)
        self.log_text = scrolledtext.ScrolledText(self.log_frame, height=10, state='disabled', wrap=tk.WORD, bg="#ffffff")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # --- Button Frame ---
        button_frame = tk.Frame(self.root, bg="#dbe5f1")
        button_frame.pack(side=tk.BOTTOM, pady=10)

        tk.Button(button_frame, text="Select Quote Folder", command=self.select_quote_folder, width=20,
                  bg="#4a90e2", fg="white").grid(row=0, column=0, padx=10, pady=5)
        tk.Button(button_frame, text="Select Output Folder", command=self.select_output_folder, width=20,
                  bg="#4a90e2", fg="white").grid(row=0, column=1, padx=10, pady=5)
        tk.Button(button_frame, text="Select Template File", command=self.select_template_file, width=20, bg="#4a90e2", fg="white").grid(row=0, column=2, padx=10, pady=5)

        # --- Action Buttons: Read Quotes, Generate Report ---
        self.read_quotes_btn = tk.Button(button_frame, text="Read Quotes", command=self.read_quotes, width=20,
                                         bg="#A9A9A9", fg="white", state='disabled')
        self.read_quotes_btn.grid(row=1, column=0, padx=10, pady=5)

        self.generate_doc_btn = tk.Button(button_frame, text="Generate Word Doc", command=self.generate_doc, width=20,
                                          bg="#A9A9A9", fg="white", state='disabled')
        self.generate_doc_btn.grid(row=1, column=1, padx=10, pady=5)

        self.update_action_buttons()  # Set initial state

    def update_action_buttons(self):
        ready = bool(self.quote_folder) and bool(self.output_folder)
        if ready:
            self.read_quotes_btn.config(state='normal', bg='#3CB371')       # MediumSeaGreen
            self.generate_doc_btn.config(state='normal', bg='#3CB371')
        else:
            self.read_quotes_btn.config(state='disabled', bg='#A9A9A9')    # Grey
            self.generate_doc_btn.config(state='disabled', bg='#A9A9A9')

    def select_template_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Word Template File",
            filetypes=[("Word Documents", "*.docx")]
        )
        if file_path:
            self.template_path = file_path
            file_name = os.path.basename(file_path)
            self.template_label.config(text=f"Template: {file_name}")
            self.log(f"Selected template file: {file_path}")

    def generate_doc(self):
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder first.")
            return

        json_path = os.path.join(self.output_folder, "combined_quotes.json")
        if not os.path.exists(json_path):
            messagebox.showerror("Error", f"combined_quotes.json not found in {self.output_folder}.")
            return

        self.log("Generating Word report...")

        try:
            data = load_json(json_path)
            broker_fee_pct = self.broker_fee_var.get()
            commission_pct = self.commission_var.get()
            associate_split = self.associate_split_var.get()
            strata_manager = self.strata_manager_var.get() if self.strata_checkbox_var.get() else "None"
            fixed_broker_fee = self.fixed_fee_var.get() if self.use_fixed_fee_var.get() else 0

            if self.template_path:
                template_path = self.template_path
            else:
                template_path = resource_path("report_template.docx")

            output_path = os.path.join(self.output_folder, "Clearlake Insurance Renewal Report 2025-2026.docx")

            generate_report(
                template_path, output_path, data,
                broker_fee_pct, commission_pct,
                associate_split, strata_manager, fixed_broker_fee
            )
            self.log(f"Report generated: {output_path}")
            messagebox.showinfo("Success", f"Report generated:\n{output_path}")

        except Exception as e:
            self.log(f"Error generating report: {e}")
            messagebox.showerror("Error", f"Failed to generate report:\n{e}")

    def toggle_strata_entry(self):
        if self.strata_checkbox_var.get():
            # Strata Manager is ticked: allow associate split input
            self.strata_entry.configure(state='normal')
            self.assoc_entry.configure(state='normal')
            self.assoc_slider.configure(state='normal')
            current_val = self.associate_split_var.get()
            self.broker_entry.config(state='normal')
            self.broker_entry.delete(0, tk.END)
            self.broker_entry.insert(0, str(100 - current_val))
            self.broker_entry.config(state='readonly')
        else:
            # No Strata Manager: lock associate split at 0%
            self.strata_entry.delete(0, tk.END)
            self.strata_entry.configure(state='disabled')
            self.associate_split_var.set(0)
            self.assoc_entry.delete(0, tk.END)
            self.assoc_entry.insert(0, "0")
            self.assoc_entry.configure(state='disabled')
            self.assoc_slider.configure(state='disabled')
            self.broker_entry.config(state='normal')
            self.broker_entry.delete(0, tk.END)
            self.broker_entry.insert(0, "100")
            self.broker_entry.config(state='readonly')

    def slider_associate_split_update(self, val=None):
        val = round(self.associate_split_var.get(), 2)
        broker_val = round(100 - val, 2)
        self.assoc_entry.delete(0, tk.END)
        self.assoc_entry.insert(0, f"{val:.2f}")
        self.broker_entry.config(state='normal')
        self.broker_entry.delete(0, tk.END)
        self.broker_entry.insert(0, f"{broker_val:.2f}")
        self.broker_entry.config(state='readonly')

    def entry_associate_split_update(self, event=None):
        try:
            val = round(float(self.assoc_entry.get()), 2)
            if val < 0: val = 0.0
            if val > 100: val = 100.0
        except ValueError:
            val = 20.0
        self.associate_split_var.set(val)
        self.slider_associate_split_update()

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

    # --- Commission Sync Logic ---
    def slider_commission_update(self, val=None):
        self.commission = self.commission_var.get()
        self.comm_entry.delete(0, tk.END)
        self.comm_entry.insert(0, str(self.commission))

    def entry_commission_update(self, event=None):
        try:
            comm = int(self.comm_entry.get())
            if comm < 0:
                comm = 0
            elif comm > 100:
                comm = 100
        except ValueError:
            comm = 10  # Reset to default on invalid
        self.commission = comm
        self.commission_var.set(comm)
        self.comm_entry.delete(0, tk.END)
        self.comm_entry.insert(0, str(comm))

    def log(self, message):
        self.log_text.config(state='normal')
        # Always prefix with '>'
        self.log_text.insert(tk.END, f"> {message}\n")
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
        self.update_action_buttons()

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder = folder
            self.update_info_label()
            self.log(f"Selected output folder: {folder}")
        self.update_action_buttons()

    def read_quotes(self):
        self.log(f"Begining quote extraction...")
        if not self.quote_folder:
            messagebox.showerror("Error", "Please select a quote folder first.")
            return

        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder first.")
            return

        self.log(f"Reading all quotes in folder: {self.quote_folder} ...")
        output_path = os.path.join(self.output_folder, "combined_quotes.json")

        try:
           
            # Pass broker_fee and commission to your extract/process code if needed
            process_folder(
                self.quote_folder,
                output_path,
                log_callback=self.log,
                longitude_option=self.longitude_option_var.get()
            )
            
            self.log(f"Extraction complete. JSON saved to: {output_path}")

        except Exception as e:
            self.log(f"Error during extraction: {e}")

    def toggle_fixed_fee(self):
        if self.use_fixed_fee_var.get():  
            self.broker_fee_var.set(0)
            self.fee_entry.config(state='normal')
            self.fee_entry.delete(0, tk.END)
            self.fee_entry.insert(0, "0")
            self.fee_entry.config(state='disabled')
            self.fee_slider.config(state='disabled')
            self.fixed_fee_entry.config(state='normal')
        else:
            self.fee_entry.config(state='normal')
            self.fee_slider.config(state='normal')
            self.fixed_fee_entry.config(state='disabled')


if __name__ == '__main__':
    root = tk.Tk()
    app = QuoteExtractorGUI(root)
    root.mainloop()
