import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
from extract import process_folder, resource_path
from report_generator import load_json, generate_report, format_currency
import ttkbootstrap as ttk
from ttkbootstrap.constants import *


class QuoteExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Insurance Quote Processor")
        self.root.update_idletasks()
        self.root.minsize(self.root.winfo_width(), self.root.winfo_height())

        self.quote_folder = ""
        self.output_folder = ""
        self.quote_count = 0
        self.broker_fee = 20
        self.commission = 20
        self.template_path = None

        self.setup_ui()

    def setup_ui(self):
        header_frame = ttk.Labelframe(self.root, text="", padding=10)
        header_frame.pack(fill='x', padx=10, pady=(15, 10))

        ttk.Separator(header_frame, orient='horizontal').pack(fill='x', pady=(0, 5))

        self.info_label = ttk.Label(header_frame, text="Quote Folder: None | Output Folder: None | Quotes Found: 0")
        self.info_label.pack(pady=(5, 0), fill='x')

        self.template_label = ttk.Label(header_frame, text="Template: Default (report_template.docx)")
        self.template_label.pack(pady=(2, 5), fill='x')

        config_frame = ttk.Labelframe(self.root, text="Pricing Configuration")
        config_frame.pack(padx=10, pady=(10, 5), fill='x')

        config_columns = ttk.Frame(config_frame)
        config_columns.pack(fill='x', padx=10)

        # Left: Default Information
        default_info_frame = ttk.Labelframe(config_columns, text="Default Information", padding=10)
        default_info_frame.grid(row=0, column=0, sticky='n')

        row1 = ttk.Frame(default_info_frame)
        row1.pack(fill='x', pady=5)

        ttk.Label(row1, text="Commission (%):", width=18, anchor='w').grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.commission_var = ttk.DoubleVar(value=self.commission)
        self.comm_entry = ttk.Entry(row1, width=5, textvariable=self.commission_var, justify='center')
        self.comm_entry.grid(row=0, column=1, padx=5)
        self.comm_slider = ttk.Scale(row1, from_=0, to=100, orient='horizontal', variable=self.commission_var, length=100)
        self.comm_slider.config(command=self.slider_commission_update)
        self.comm_slider.grid(row=0, column=2, padx=5)

        ttk.Label(row1, text="Broker Fee (%):", width=18, anchor='w').grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.broker_fee_var = ttk.DoubleVar(value=self.broker_fee)
        self.fee_entry = ttk.Entry(row1, width=5, textvariable=self.broker_fee_var, justify='center')
        self.fee_entry.grid(row=1, column=1, padx=5)
        self.fee_slider = ttk.Scale(row1, from_=0, to=100, orient='horizontal', variable=self.broker_fee_var, length=100)
        self.fee_slider.config(command=self.slider_broker_fee_update)
        self.fee_slider.grid(row=1, column=2, padx=5)

        self.use_fixed_fee_var = ttk.BooleanVar(value=False)
        self.fixed_fee_check = ttk.Checkbutton(row1, text="Use Fixed $ Fee", variable=self.use_fixed_fee_var, command=self.toggle_fixed_fee)
        self.fixed_fee_check.grid(row=1, column=3, padx=10, sticky="w")

        self.fixed_fee_var = ttk.DoubleVar(value=0)
        self.fixed_fee_entry = ttk.Entry(row1, width=7, textvariable=self.fixed_fee_var, justify='center', state='disabled')
        self.fixed_fee_entry.grid(row=1, column=4, padx=5)

        row2 = ttk.Frame(default_info_frame)
        row2.pack(fill='x', pady=5)
        ttk.Label(row2, text="Associate Split (%):", width=18, anchor='w').pack(side='left')
        self.associate_split_var = ttk.DoubleVar(value=60.00)
        self.broker_split_var = ttk.DoubleVar(value=40.00)
        self.assoc_entry = ttk.Entry(row2, width=7, textvariable=self.associate_split_var, justify='center')
        self.assoc_entry.pack(side='left', padx=(5, 3))
        self.assoc_slider = ttk.Scale(row2, from_=0, to=100, orient='horizontal', variable=self.associate_split_var, length=100)
        self.assoc_slider.config(command=self.slider_associate_split_update)
        self.assoc_slider.pack(side='left', padx=(5, 8))
        self.broker_entry = ttk.Entry(row2, width=7, justify='center')
        self.broker_entry.pack(side='left', padx=(3, 5))
        self.broker_entry.insert(0, "40.00")
        self.broker_entry.config(state='readonly')
        ttk.Label(row2, text=": Broker Share (%)").pack(side='left')

        row3 = ttk.Frame(default_info_frame)
        row3.pack(fill='x', pady=5)
        self.strata_checkbox_var = ttk.BooleanVar(value=True)
        self.strata_checkbox = ttk.Checkbutton(row3, text="Strata Manager:", variable=self.strata_checkbox_var)
        self.strata_checkbox.pack(side='left')
        self.strata_manager_var = ttk.StringVar()
        self.strata_entry = ttk.Entry(row3, width=30, textvariable=self.strata_manager_var)
        self.strata_entry.pack(side='left', padx=(5, 10))

        # Move log under default_info_frame
        log_frame = ttk.Labelframe(default_info_frame, text="Log")
        log_frame.pack(fill='x', padx=5, pady=(10, 0))
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', wrap='word')
        self.log_text.pack(fill='both', padx=5, pady=5)

        # Right: Previous Invoice Section
        invoice_frame = ttk.Labelframe(config_columns, text="Previous Invoice", padding=10)
        invoice_frame.grid(row=0, column=1, sticky='n')

        self.include_invoice_var = ttk.BooleanVar(value=False)
        invoice_check = ttk.Checkbutton(invoice_frame, text="Include last year invoice", variable=self.include_invoice_var, command=self.toggle_invoice_fields)
        invoice_check.pack(anchor='w', pady=(0, 5))

        self.invoice_fields = {}
        fields = [
            ("Insurer / Underwriter", "str"),
            ("Base Premium", "float"),
            ("ESL", "float"),
            ("GST", "float"),
            ("Stamp Duty", "float"),
            ("Insurer / underwriter fee", "float"),
            ("Insurer / underwriter GST", "float"),
            ("Broker Fee", "float"),
            ("Broker Fee GST", "float"),
            ("Total Premium", "float"),
        ]

        for i, (label, ftype) in enumerate(fields):
            frame = ttk.Frame(invoice_frame)
            frame.pack(fill='x', pady=2)
            label_widget = ttk.Label(frame, text=label + ":", width=25, anchor='e')
            label_widget.pack(side='left', padx=(0, 5))
            var = tk.StringVar()
            if ftype == "float":
                var.set("0.00")
            entry = ttk.Entry(frame, textvariable=var, justify='right', state='disabled')
            entry.pack(side='left', fill='x', expand=True)
            self.invoice_fields[label] = (entry, var, ftype)

        container = ttk.Frame(self.root)
        container.pack(pady=(5, 15), fill='x')
        button_frame = ttk.Frame(container)
        button_frame.pack()

        ttk.Button(button_frame, text="Select Quote Folder", command=self.select_quote_folder, width=20, bootstyle="primary").grid(row=0, column=0, padx=10, pady=5)
        ttk.Button(button_frame, text="Select Output Folder", command=self.select_output_folder, width=20, bootstyle="primary").grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(button_frame, text="Select Template File", command=self.select_template_file, width=20, bootstyle="primary").grid(row=0, column=2, padx=10, pady=5)

        self.read_quotes_btn = ttk.Button(button_frame, text="Read Quotes", command=self.read_quotes, width=20, state='disabled', bootstyle="secondary")
        self.read_quotes_btn.grid(row=1, column=0, padx=10, pady=5)

        self.generate_doc_btn = ttk.Button(button_frame, text="Generate Word Doc", command=self.generate_doc, width=20, state='disabled', bootstyle="secondary")
        self.generate_doc_btn.grid(row=1, column=1, padx=10, pady=5)

        self.update_action_buttons()

    def toggle_invoice_fields(self):
        state = 'normal' if self.include_invoice_var.get() else 'disabled'
        for entry, _, _ in self.invoice_fields.values():
            entry.config(state=state)

    # Other methods: slider_commission_update, slider_broker_fee_update, slider_associate_split_update, toggle_fixed_fee, generate_doc,
    # update_action_buttons, log, select_quote_folder, select_output_folder, read_quotes, etc. remain unchanged from original file.


    def update_action_buttons(self):
        ready = bool(self.quote_folder) and bool(self.output_folder)
        if ready:
            self.read_quotes_btn.config(state='normal', bootstyle='success')
            self.generate_doc_btn.config(state='normal', bootstyle='success')
        else:
            self.read_quotes_btn.config(state='disabled', bootstyle='secondary')
            self.generate_doc_btn.config(state='disabled', bootstyle='secondary')


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
        try:
            val = round(float(val))
            if val < 0: val = 0
            if val > 100: val = 100
            self.associate_split_var.set(val)
            self.broker_split_var.set(100 - val)
            self.assoc_entry.delete(0, tk.END)
            self.assoc_entry.insert(0, f"{val:.2f}")
            self.broker_entry.config(state='normal')
            self.broker_entry.delete(0, tk.END)
            self.broker_entry.insert(0, f"{(100 - val):.2f}")
            self.broker_entry.config(state='readonly')
        except (ValueError, tk.TclError):
            pass

    def entry_associate_split_update(self, event=None):
        try:
            val = round(float(self.assoc_entry.get()), 2)
            if val < 0: val = 0.0
            if val > 100: val = 100.0
        except ValueError:
            val = 20.0
        self.associate_split_var.set(val)
        self.broker_split_var.set(100 - val)
        self.slider_associate_split_update()

    def slider_broker_fee_update(self, val=None):
        try:
            val = round(float(val))
            self.broker_fee_var.set(val)
            self.fee_entry.delete(0, tk.END)
            self.fee_entry.insert(0, f"{val:.2f}")
        except (ValueError, tk.TclError):
            pass

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
        try:
            val = round(float(val))
            self.commission_var.set(val)
            self.comm_entry.delete(0, tk.END)
            self.comm_entry.insert(0, f"{val:.2f}")
        except (ValueError, tk.TclError):
            pass

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
    root = ttk.Window(themename="yeti")
    app = QuoteExtractorGUI(root)
    root.mainloop()
