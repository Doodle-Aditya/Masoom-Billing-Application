# gui/app_gui.py
import os
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from num2words import num2words

from excel_handler.excel_manager import (
    save_to_excel, next_ref_no, load_donors_from_pdf
)
from pdf_generator.pdf_overlay import create_overlay   # UPDATED
from pdf_generator.pdf_merge import merge_pdfs

TEMPLATE_PDF = os.path.join("pdf_generator", "template.pdf")

class BillingApp:
    def __init__(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.app = ctk.CTk()
        self.app.title("Billing Application")
        self.app.geometry("820x820")

        self.build_ui()

        # Load donors
        donors = load_donors_from_pdf()
        if donors:
            self.donor_dropdown.configure(values=donors)
        else:
            self.donor_dropdown.configure(values=["NTT", "Masoom", "Other"])

    def build_ui(self):
        self.scroll = ctk.CTkScrollableFrame(self.app, width=800, height=780)
        self.scroll.pack(padx=10, pady=10, fill="both", expand=True)

        frm = self.scroll
        row = 0

        def add_label_entry(label_text, var_attr):
            nonlocal row
            lbl = ctk.CTkLabel(frm, text=label_text)
            lbl.grid(row=row, column=0, sticky="w", padx=8, pady=6)
            ent = ctk.CTkEntry(frm, width=420)
            ent.grid(row=row, column=1, padx=8, pady=6)
            setattr(self, var_attr, ent)
            row += 1
            return ent

        # Manual Entries
        add_label_entry("College Name", "college_name_entry")
        add_label_entry("Class", "class_entry")

        # Reference No
        lbl = ctk.CTkLabel(frm, text="Reference No (Auto)")
        lbl.grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.ref_no_var = ctk.StringVar(value=next_ref_no())
        ref_entry = ctk.CTkEntry(frm, textvariable=self.ref_no_var, width=420)
        ref_entry.grid(row=row, column=1, padx=8, pady=6)
        row += 1

        # Date Auto
        lbl = ctk.CTkLabel(frm, text="Date (Auto)")
        lbl.grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.date_var = ctk.StringVar(value=datetime.today().strftime("%d/%m/%Y"))
        date_entry = ctk.CTkEntry(frm, textvariable=self.date_var, width=420)
        date_entry.grid(row=row, column=1, padx=8, pady=6)
        row += 1

        add_label_entry("Student Name", "student_name_entry")
        add_label_entry("College Fees (numeric)", "college_fees_entry")

        # Total Fees Auto
        lbl = ctk.CTkLabel(frm, text="Total Fees (Auto)")
        lbl.grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.total_fees_var = ctk.StringVar(value="")
        total_entry = ctk.CTkEntry(frm, textvariable=self.total_fees_var,
                                   width=420, state="readonly")
        total_entry.grid(row=row, column=1, padx=8, pady=6)
        row += 1

        # Masoom 75% Auto
        lbl = ctk.CTkLabel(frm, text="Masoom Contribution 75% (Auto)")
        lbl.grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.masoom_var = ctk.StringVar(value="")
        masoom_entry = ctk.CTkEntry(frm, textvariable=self.masoom_var,
                                    width=420, state="readonly")
        masoom_entry.grid(row=row, column=1, padx=8, pady=6)
        row += 1

        # Student 25%
        lbl = ctk.CTkLabel(frm, text="Student Contribution 25% (Auto)")
        lbl.grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.student_var = ctk.StringVar(value="")
        student_entry = ctk.CTkEntry(frm, textvariable=self.student_var,
                                     width=420, state="readonly")
        student_entry.grid(row=row, column=1, padx=8, pady=6)
        row += 1

        # Payable
        lbl = ctk.CTkLabel(frm, text="Payable (Auto)")
        lbl.grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.payable_var = ctk.StringVar(value="")
        payable_entry = ctk.CTkEntry(frm, textvariable=self.payable_var,
                                     width=420, state="readonly")
        payable_entry.grid(row=row, column=1, padx=8, pady=6)
        row += 1

        # Words
        lbl = ctk.CTkLabel(frm, text="Rs In Words (Auto)")
        lbl.grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.words_var = ctk.StringVar(value="")
        words_entry = ctk.CTkEntry(frm, textvariable=self.words_var,
                                   width=420, state="readonly")
        words_entry.grid(row=row, column=1, padx=8, pady=6)
        row += 1

        # Donor dropdown
        lbl = ctk.CTkLabel(frm, text="Donor Name")
        lbl.grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.donor_dropdown = ctk.CTkComboBox(frm, values=["Loading..."],
                                              width=420)
        self.donor_dropdown.grid(row=row, column=1, padx=8, pady=6)
        row += 1

        # Bank/Cheque fields
        add_label_entry("Cheque Issue On Name", "cheque_issue_entry")
        add_label_entry("Account Holder Name", "account_holder_entry")
        add_label_entry("Bank Name", "bank_name_entry")
        add_label_entry("Account Number", "account_number_entry")
        add_label_entry("IFSC Code", "ifsc_entry")
        add_label_entry("Cheque Number", "cheque_number_entry")

        add_label_entry("Prepared By (for sign)", "prepared_by_entry")

        # Buttons
        btn_frame = ctk.CTkFrame(frm)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=12)

        calc_btn = ctk.CTkButton(btn_frame, text="Calculate",
                                 command=self.calculate_all)
        calc_btn.grid(row=0, column=0, padx=8)

        gen_btn = ctk.CTkButton(btn_frame, text="Generate Bill",
                                command=self.generate_bill)
        gen_btn.grid(row=0, column=1, padx=8)

        clear_btn = ctk.CTkButton(btn_frame, text="Clear",
                                  command=self.clear_form)
        clear_btn.grid(row=0, column=2, padx=8)

    def calculate_all(self):
        try:
            fees_text = self.college_fees_entry.get().strip().replace(",", "")
            fees = float(fees_text) if fees_text else 0.0
        except ValueError:
            messagebox.showerror("Invalid input",
                                 "College Fees must be a number")
            return

        total = fees
        masoom = round(total * 0.75)
        student = round(total * 0.25)
        payable = masoom

        self.total_fees_var.set(str(int(total)
                                    if total.is_integer() else total))
        self.masoom_var.set(str(masoom))
        self.student_var.set(str(student))
        self.payable_var.set(str(payable))

        try:
            words = num2words(int(payable), to="cardinal",
                              lang="en_IN").title() + " Only"
        except Exception:
            words = num2words(int(payable), to="cardinal").title() + " Only"

        self.words_var.set(words)

    def collect_data(self):
        return {
            "college_name": self.college_name_entry.get().strip(),
            "class": self.class_entry.get().strip(),
            "ref_no": self.ref_no_var.get().strip(),
            "date": self.date_var.get().strip(),
            "student_name": self.student_name_entry.get().strip(),
            "college_fees": self.college_fees_entry.get().strip(),
            "total_fees": self.total_fees_var.get().strip(),
            "masoom_contribution": self.masoom_var.get().strip(),
            "student_contribution": self.student_var.get().strip(),
            "payable": self.payable_var.get().strip(),
            "rs_in_words": self.words_var.get().strip(),
            "donor_name": self.donor_dropdown.get().strip(),
            "cheque_issue_name": self.cheque_issue_entry.get().strip(),
            "account_holder": self.account_holder_entry.get().strip(),
            "bank_name": self.bank_name_entry.get().strip(),
            "account_number": self.account_number_entry.get().strip(),
            "ifsc": self.ifsc_entry.get().strip(),
            "cheque_number": self.cheque_number_entry.get().strip(),
            "prepared_by": self.prepared_by_entry.get().strip()
        }

    def generate_bill(self):
        self.calculate_all()

        # Refresh reference number
        self.ref_no_var.set(next_ref_no())
        data = self.collect_data()

        if not data["student_name"]:
            messagebox.showerror("Missing field",
                                 "Please enter Student Name")
            return

        # Save to Excel
        try:
            save_to_excel(data)
        except Exception as e:
            messagebox.showerror("Excel error",
                                 f"Failed to save to Excel:\n{e}")
            return

        # Generate PDF paths
        overlay_path = os.path.join("pdf_generator", "overlay.pdf")
        output_fname = f"Bill_{data['student_name'].replace(' ', '_')}_{data['ref_no'].replace('/', '_')}.pdf"
        output_path = os.path.join("output", "bills", output_fname)

        # FIXED: Updated call to create_overlay
        try:
            create_overlay(data, overlay_path, TEMPLATE_PDF)
            merge_pdfs(TEMPLATE_PDF, overlay_path, output_path)
        except Exception as e:
            messagebox.showerror("PDF error",
                                 f"Failed to generate PDF:\n{e}")
            return

        messagebox.showinfo("Success",
                            f"Bill generated:\n{output_path}\n\nSaved to Excel.")

        # Refresh next ref no.
        self.ref_no_var.set(next_ref_no())

    def clear_form(self):
        self.college_name_entry.delete(0, "end")
        self.class_entry.delete(0, "end")
        self.student_name_entry.delete(0, "end")
        self.college_fees_entry.delete(0, "end")
        self.total_fees_var.set("")
        self.masoom_var.set("")
        self.student_var.set("")
        self.payable_var.set("")
        self.words_var.set("")
        self.cheque_issue_entry.delete(0, "end")
        self.account_holder_entry.delete(0, "end")
        self.bank_name_entry.delete(0, "end")
        self.account_number_entry.delete(0, "end")
        self.ifsc_entry.delete(0, "end")
        self.cheque_number_entry.delete(0, "end")
        self.prepared_by_entry.delete(0, "end")

    def run(self):
        self.app.mainloop()
