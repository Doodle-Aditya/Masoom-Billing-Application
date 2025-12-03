import os
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from num2words import num2words

from excel_handler.excel_manager import (
    save_to_excel, next_ref_no, load_donors_from_pdf
)

from excel_handler.fill_template import fill_excel_template, excel_to_pdf


class BillingApp:
    def __init__(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.app = ctk.CTk()
        self.app.title("Billing Application")
        self.app.geometry("850x820")

        self.build_ui()

        # Load donors
        donors = load_donors_from_pdf()
        if donors:
            self.donor_dropdown.configure(values=donors)
        else:
            self.donor_dropdown.configure(values=["NTT", "Masoom", "Other"])

    def build_ui(self):
        self.scroll = ctk.CTkScrollableFrame(self.app, width=820, height=780)
        self.scroll.pack(padx=10, pady=10, fill="both", expand=True)

        frm = self.scroll
        row = 0

        def add_label_entry(label_text, attr):
            nonlocal row
            ctk.CTkLabel(frm, text=label_text).grid(
                row=row, column=0, sticky="w", padx=8, pady=6
            )
            entry = ctk.CTkEntry(frm, width=420)
            entry.grid(row=row, column=1, padx=8, pady=6)
            setattr(self, attr, entry)
            row += 1
            return entry

        # Basic Inputs
        add_label_entry("College Name", "college_name_entry")
        add_label_entry("Class", "class_entry")

        # Reference No (Auto)
        ctk.CTkLabel(frm, text="Reference No (Auto)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.ref_no_var = ctk.StringVar(value=next_ref_no())
        ctk.CTkEntry(frm, textvariable=self.ref_no_var, width=420).grid(row=row, column=1, padx=8, pady=6)
        row += 1

        # Date Auto
        ctk.CTkLabel(frm, text="Date (Auto)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.date_var = ctk.StringVar(value=datetime.today().strftime("%d/%m/%Y"))
        ctk.CTkEntry(frm, textvariable=self.date_var, width=420).grid(row=row, column=1, padx=8, pady=6)
        row += 1

        add_label_entry("Student Name", "student_name_entry")

        # Receipt 1 & Receipt 2
        add_label_entry("College Receipt 1 Amount", "receipt1_entry")
        add_label_entry("College Receipt 2 Amount (optional)", "receipt2_entry")

        # Auto fields (editable)
        def add_auto(label, var_name):
            nonlocal row
            ctk.CTkLabel(frm, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=6)
            var = ctk.StringVar(value="")
            ent = ctk.CTkEntry(frm, textvariable=var, width=420)  # editable
            ent.grid(row=row, column=1, padx=8, pady=6)
            setattr(self, var_name, var)
            row += 1

        add_auto("Total Fees (Auto)", "total_fees_var")
        add_auto("Masoom Contribution 75% (Auto)", "masoom_var")
        add_auto("Student Contribution 25% (Auto)", "student_var")
        add_auto("Payable (Auto)", "payable_var")
        add_auto("Rs In Words (Auto)", "words_var")

        # Donor Dropdown
        ctk.CTkLabel(frm, text="Donor Name").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.donor_dropdown = ctk.CTkComboBox(frm, width=420, values=["Loading..."])
        self.donor_dropdown.grid(row=row, column=1, padx=8, pady=6)
        row += 1

        # Bank Details
        add_label_entry("Cheque Issue On Name", "cheque_issue_entry")
        add_label_entry("Account Holder Name", "account_holder_entry")
        add_label_entry("Bank Name", "bank_name_entry")
        add_label_entry("Account Number", "account_number_entry")
        add_label_entry("IFSC Code", "ifsc_entry")
        add_label_entry("Cheque Number", "cheque_number_entry")

        # Prepared By
        add_label_entry("Prepared By (Sign)", "prepared_by_entry")

        # Buttons
        btn_frame = ctk.CTkFrame(frm)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=12)

        ctk.CTkButton(btn_frame, text="Calculate", command=self.calculate_all).grid(row=0, column=0, padx=8)
        ctk.CTkButton(btn_frame, text="Generate Bill", command=self.generate_bill).grid(row=0, column=1, padx=8)
        ctk.CTkButton(btn_frame, text="Clear", command=self.clear_form).grid(row=0, column=2, padx=8)

    def calculate_all(self):
        try:
            r1 = float(self.receipt1_entry.get().replace(",", "") or 0)
            r2 = float(self.receipt2_entry.get().replace(",", "") or 0)
        except:
            messagebox.showerror("Invalid Input", "Receipt amounts must be numbers")
            return

        total = r1 + r2
        masoom = round(total * 0.75)
        student = round(total * 0.25)

        self.total_fees_var.set(str(int(total)))
        self.masoom_var.set(str(int(masoom)))
        self.student_var.set(str(int(student)))
        self.payable_var.set(str(int(masoom)))

        try:
            words = num2words(masoom, to="cardinal", lang="en_IN").title() + " Only"
        except:
            words = num2words(masoom).title() + " Only"

        self.words_var.set(words)

    def collect_data(self):
        return {
            "college_name": self.college_name_entry.get(),
            "class": self.class_entry.get(),
            "ref_no": self.ref_no_var.get(),
            "date": self.date_var.get(),
            "student_name": self.student_name_entry.get(),
            "receipt1": self.receipt1_entry.get(),
            "receipt2": self.receipt2_entry.get(),
            "total_fees": self.total_fees_var.get(),
            "masoom_contribution": self.masoom_var.get(),
            "student_contribution": self.student_var.get(),
            "payable": self.payable_var.get(),
            "rs_in_words": self.words_var.get(),
            "donor_name": self.donor_dropdown.get(),
            "cheque_issue_name": self.cheque_issue_entry.get(),
            "account_holder": self.account_holder_entry.get(),
            "bank_name": self.bank_name_entry.get(),
            "account_number": self.account_number_entry.get(),
            "ifsc": self.ifsc_entry.get(),
            "cheque_number": self.cheque_number_entry.get(),
            "prepared_by": self.prepared_by_entry.get(),
        }

    def generate_bill(self):
        self.calculate_all()

        # Refresh reference number
        self.ref_no_var.set(next_ref_no())
        data = self.collect_data()

        if not data["student_name"]:
            messagebox.showerror("Missing Field", "Enter student name")
            return

        save_to_excel(data)

        # File paths
        excel_bill_path = os.path.join(
            "output",
            "bills",
            f"Bill_{data['student_name'].replace(' ', '_')}_{data['ref_no'].replace('/', '_')}.xlsx"
        )

        pdf_path = excel_bill_path.replace(".xlsx", ".pdf")

        # Fill template
        fill_excel_template(data, excel_bill_path)

        # Convert to PDF
        excel_to_pdf(excel_bill_path, pdf_path)

        messagebox.showinfo("Success", f"Bill Generated:\n{pdf_path}")

    def clear_form(self):
        for field in [
            "college_name_entry", "class_entry", "student_name_entry",
            "receipt1_entry", "receipt2_entry", "cheque_issue_entry",
            "account_holder_entry", "bank_name_entry", "account_number_entry",
            "ifsc_entry", "cheque_number_entry", "prepared_by_entry",
        ]:
            getattr(self, field).delete(0, "end")

        self.total_fees_var.set("")
        self.masoom_var.set("")
        self.student_var.set("")
        self.payable_var.set("")
        self.words_var.set("")

    def run(self):
        self.app.mainloop()
