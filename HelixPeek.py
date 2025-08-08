import sys
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter.constants import *
import csv
import os
from typing import Any
from collections import defaultdict
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
import sys

def resource_path(filename):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)

CSV_FILENAME = resource_path("Spreadsheetdata.csv")

class Toplevel1:
        
    def __init__(self, root=None):
        '''This class configures and populates the toplevel window.'''
        self.root = root
        self.root.geometry("1024x768+1056+400")
        self.root.minsize(120, 770)
        self.root.maxsize(4484, 1421)
        self.root.resizable(0, 0)
        self.root.title("HelixPeek | version 0.70β")
        self.root.configure(background="#b0c4c4")
        self.root.configure(highlightbackground="#b0c4c4")
        self.root.configure(highlightcolor="#000000")

        # Load spreadsheet data
        self.spreadsheet_data = self.load_spreadsheet_data(CSV_FILENAME)

        # Initialize category checkboxes variables
        self.selected_categories = {
            "HEALTH": tk.BooleanVar(value=True),
            "FITNESS": tk.BooleanVar(value=True),
            "APPEARANCE": tk.BooleanVar(value=True),
            "HEALTH RISKS": tk.BooleanVar(value=True),
            "DIET": tk.BooleanVar(value=True),
            "COGNITIVE & EMOTIONAL TRAITS": tk.BooleanVar(value=True),
            "MEDICATION": tk.BooleanVar(value=True),
            "SLEEP": tk.BooleanVar(value=True),
        }

        # Store the last processed DNA lines
        self.last_dna_lines = []

        # Menu bar
        self.menubar = tk.Menu(self.root, font="TkMenuFont", bg="#d9d9d9", fg="#000000")
        self.root.configure(menu=self.menubar)

        # Style configuration
        self._style_code()

        # Checkbuttons for categories
        self.TCheckbutton1_1_1_2 = ttk.Checkbutton(self.root, style='Custom.TCheckbutton')
        self.TCheckbutton1_1_1_2.place(relx=0.022, rely=0.153, relwidth=0.059, relheight=0.0, height=21)
        self.TCheckbutton1_1_1_2.configure(variable=self.selected_categories["HEALTH"])
        self.TCheckbutton1_1_1_2.configure(text='''Health''')
        self.TCheckbutton1_1_1_2.configure(compound='left')

        self.TCheckbutton1_1_1 = ttk.Checkbutton(self.root, style='Custom.TCheckbutton')
        self.TCheckbutton1_1_1.place(relx=0.1, rely=0.153, relwidth=0.061, relheight=0.0, height=21)
        self.TCheckbutton1_1_1.configure(variable=self.selected_categories["FITNESS"])
        self.TCheckbutton1_1_1.configure(text='''Fitness''')
        self.TCheckbutton1_1_1.configure(compound='left')

        self.TCheckbutton1_2 = ttk.Checkbutton(self.root, style='Custom.TCheckbutton')
        self.TCheckbutton1_2.place(relx=0.183, rely=0.153, relwidth=0.094, relheight=0.0, height=21)
        self.TCheckbutton1_2.configure(variable=self.selected_categories["APPEARANCE"])
        self.TCheckbutton1_2.configure(text='''Appearance''')
        self.TCheckbutton1_2.configure(compound='left')

        self.TCheckbutton1_3 = ttk.Checkbutton(self.root, style='Custom.TCheckbutton')
        self.TCheckbutton1_3.place(relx=0.288, rely=0.153, relwidth=0.094, relheight=0.0, height=21)
        self.TCheckbutton1_3.configure(variable=self.selected_categories["HEALTH RISKS"])
        self.TCheckbutton1_3.configure(text='''Health Risks''')
        self.TCheckbutton1_3.configure(compound='left')

        self.TCheckbutton1_1_1_1 = ttk.Checkbutton(self.root, style='Custom.TCheckbutton')
        self.TCheckbutton1_1_1_1.place(relx=0.476, rely=0.153, relwidth=0.05, relheight=0.0, height=21)
        self.TCheckbutton1_1_1_1.configure(variable=self.selected_categories["DIET"])
        self.TCheckbutton1_1_1_1.configure(text='''Diet''')
        self.TCheckbutton1_1_1_1.configure(compound='left')

        self.TCheckbutton1 = ttk.Checkbutton(self.root, style='Custom.TCheckbutton')
        self.TCheckbutton1.place(relx=0.542, rely=0.153, relwidth=0.189, relheight=0.0, height=21)
        self.TCheckbutton1.configure(variable=self.selected_categories["COGNITIVE & EMOTIONAL TRAITS"])
        self.TCheckbutton1.configure(text='''Cognitive & Emotional Traits''')
        self.TCheckbutton1.configure(compound='left')

        self.TCheckbutton1_1 = ttk.Checkbutton(self.root, style='Custom.TCheckbutton')
        self.TCheckbutton1_1.place(relx=0.732, rely=0.153, relwidth=0.096, relheight=0.0, height=21)
        self.TCheckbutton1_1.configure(variable=self.selected_categories["MEDICATION"])
        self.TCheckbutton1_1.configure(text='''Medication''')
        self.TCheckbutton1_1.configure(compound='left')

        self.TCheckbutton1_1_1_2_1 = ttk.Checkbutton(self.root, style='Custom.TCheckbutton')
        self.TCheckbutton1_1_1_2_1.place(relx=0.398, rely=0.153, relwidth=0.058, relheight=0.0, height=21)
        self.TCheckbutton1_1_1_2_1.configure(variable=self.selected_categories["SLEEP"])
        self.TCheckbutton1_1_1_2_1.configure(text='''Sleep''')
        self.TCheckbutton1_1_1_2_1.configure(compound='left')

        # Buttons
        self.Button1_1_1_1 = tk.Button(self.root)
        self.Button1_1_1_1.place(relx=0.022, rely=0.038, height=26, width=177)
        self.Button1_1_1_1.configure(activebackground="#b0c4c4", activeforeground="black", background="#b3cfcf")
        self.Button1_1_1_1.configure(disabledforeground="#a3a3a3", font="-family {Segoe UI} -size 9")
        self.Button1_1_1_1.configure(foreground="#2F5D5E", highlightbackground="#b0c4c4", highlightcolor="#000000")
        self.Button1_1_1_1.configure(text='''Include rsID & genotype''')
        self.Button1_1_1_1.configure(command=self.toggle_rsID_genotype)
        self.include_rsid_genotype = False

        self.Button1_1 = tk.Button(self.root)
        self.Button1_1.place(relx=0.022, rely=0.088, height=26, width=177)
        self.Button1_1.configure(activebackground="#d9d9d9", activeforeground="black", background="#b3cfcf")
        self.Button1_1.configure(disabledforeground="#a3a3a3", font="-family {Segoe UI} -size 9")
        self.Button1_1.configure(foreground="#2F5D5E", highlightbackground="#d9d9d9", highlightcolor="#000000")
        self.Button1_1.configure(text='''Save Report as PDF''')
        self.Button1_1.configure(command=self.save_report_as_pdf)

        self.Button1_1_1 = tk.Button(self.root)
        self.Button1_1_1.place(relx=0.212, rely=0.038, height=66, width=317)
        self.Button1_1_1.configure(activebackground="#d9d9d9", activeforeground="black", background="#b3cfcf")
        self.Button1_1_1.configure(disabledforeground="#a3a3a3", font="-family {Segoe UI} -size 14")
        self.Button1_1_1.configure(foreground="#2F5D5E", highlightbackground="#d9d9d9", highlightcolor="#000000")
        self.Button1_1_1.configure(text='''Open raw DNA data file''')
        self.Button1_1_1.configure(command=self.upload_file)

        # Logo Label
        self.Label2 = tk.Label(self.root, text="To hide certain categories, uncheck them\nbefore opening the raw DNA file. Click the\ntop-left button to show rsIDs and genotypes.\nYou can copy the text or export it as a PDF.", font=("Segoe UI", 10), bg="#b0c4c4", fg="#4d7373")
        self.Label2.place(relx=0.56, rely=0.008, height=111, width=300)
        try:
            photo_location = resource_path("dnalogo_small.png")
            self._img0 = tk.PhotoImage(file=photo_location)
            self.Label1 = tk.Label(self.root)
            self.Label1.place(relx=0.865, rely=0.014, height=141, width=111)
            self.Label1.configure(background="#b0c4c4", image=self._img0)
        except Exception:
            self.Label1 = tk.Label(self.root)
            self.Label1.place(relx=0.850, rely=0.014, height=141, width=111)
            self.Label1.configure(background="#b0c4c4", text='''Logo not found''')

        # Scrolled Text
        self.Scrolledtext1 = scrolledtext.ScrolledText(self.root)
        self.Scrolledtext1.place(relx=0.022, rely=0.199, relheight=0.787, relwidth=0.955)
        self.Scrolledtext1.configure(background="#0078d7", foreground="#FFFFFF", font=("Courier", 12))
        self.Scrolledtext1.configure(highlightbackground="#d9d9d9", highlightcolor="#000000")
        self.Scrolledtext1.configure(insertbackground="#000000", insertborderwidth="3")
        self.Scrolledtext1.configure(selectbackground="#d9d9d9", selectforeground="black", wrap="word")
        self.Scrolledtext1.insert(tk.END, "Load a DNA file to begin analysis.")

    def _style_code(self):
        style = ttk.Style()
        style.theme_use('default')
        style.configure('.', font="TkDefaultFont")
        style.configure('Custom.TCheckbutton', background='#b0c4c4', foreground='#000000')
        if sys.platform == "win32":
            try:
                style.tk.call('source', os.path.join(os.path.dirname(__file__), 'themes', 'default.tcl'))
            except:
                pass

    def load_spreadsheet_data(self, filename: str) -> list[dict[str, Any]]:
        if not os.path.exists(filename):
            messagebox.showerror("Error", f"CSV file not found:\n{filename}")
            return []
        try:
            with open(filename, "r", encoding="utf-8") as f:
                dialect = csv.Sniffer().sniff(f.read(1024), delimiters=";,")
                f.seek(0)
                reader = csv.DictReader(f, dialect=dialect)
                if reader.fieldnames:
                    reader.fieldnames = [name.strip('" ') for name in reader.fieldnames]
                return list(reader)
        except Exception as e:
            messagebox.showerror("Error", f"Error reading CSV file:\n{e}")
            return []

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                self.last_dna_lines = file.readlines()
            self.Scrolledtext1.delete("1.0", tk.END)
            self.Scrolledtext1.insert(tk.END, "Processing DNA file, please wait...")
            self.root.update()
            report = self.analyze_dna(self.last_dna_lines)
            self.Scrolledtext1.delete("1.0", tk.END)
            self.Scrolledtext1.insert(tk.END, report)
        except Exception as e:
            messagebox.showerror("Error", f"Error processing DNA file:\n{e}")

    def analyze_dna(self, lines: list[str]) -> str:
        y_markers = {"rs2032597", "rs768983", "rs390240", "rs390307", "rs1229982"}
        found_y_markers = set()
        genotypes = {}
        categorized = defaultdict(list)
        hair_color_results = []

        for line in lines:
            if line.startswith("#") or not line.strip():
                continue
            parts = line.strip().split("\t")
            if len(parts) < 5:
                continue
            rsid = parts[0]
            allele1 = parts[3].upper()
            allele2 = parts[4].upper()
            genotype = "".join(sorted([allele1, allele2]))
            genotypes[rsid] = genotype
            if rsid in y_markers and allele1 == allele2 and all(a in "ACGT" for a in (allele1, allele2)):
                found_y_markers.add(rsid)

        genetic_sex = self.detect_sex(genotypes)
        is_male = "Male" in genetic_sex

        for entry in self.spreadsheet_data:
            rule = str(entry.get("genotype_rule", "")).strip('" ')
            trait = str(entry.get("trait", "")).strip('" ')
            priority = int(entry.get("priority", 1))
            result = str(entry.get("result", "")).strip('" ').replace("\\n", "\n")
            category = str(entry.get("category", "Other")).strip('" ')

            category_key = category.upper()
            if not self.selected_categories.get(category_key, tk.BooleanVar(value=True)).get():
                continue

            if is_male and trait.lower() == "baldness risk" and "female pattern" in result.lower():
                continue
            if not is_male and trait.lower() == "baldness risk" and "male pattern" in result.lower():
                continue
            if not is_male and trait.lower() == "prostate cancer":
                continue

            if self.match_rule(genotypes, rule):
                result_text = f"\n{trait}:\n{result}"
                
                
                if self.include_rsid_genotype:
                   # result_text += f"\n(Rule: {rule}, Trait: {trait})"
                    rsids_in_rule = [p.strip().split("=")[0] for p in rule.split("AND") if "=" in p]
                    for rid in rsids_in_rule:
                        genotype = genotypes.get(rid.strip(), "N/A")
                        result_text += f"\n(rsID: {rid}, Genotype: {genotype})"
                   
                if trait.lower() == "hair color":
                    hair_color_results.append((priority, category, result_text))
                else:
                    categorized[category_key].append((category, result_text))

        if hair_color_results:
            hair_color_results.sort(key=lambda x: x[0])
            _, category, result = hair_color_results[0]
            categorized[category.upper()].append((category, result))
        report = f"""DNA-based Trait Report for Genetic {'Male (XY)' if is_male else 'Female (XX)'}

This report presents results from analyzing selected genetic markers in your raw DNA data, matched to traits in scientific literature. It covers categories like physical appearance (e.g., eye color, hair type), health predispositions (e.g., disease risks, nutrient metabolism), nutrition preferences (e.g., taste sensitivity, dietary tolerances), sleep patterns (e.g., chronotype, sleep duration), and unique traits (e.g., photic sneeze reflex). Each trait is linked to specific single nucleotide polymorphisms (SNPs) and their genotypes, based on publicly available research.

Important: This report is for entertainment and educational purposes only and is not a medical diagnostic tool. Genetic traits are influenced by multiple genes and environmental factors, so results are general interpretations, not definitive. Do not use this information for medical decisions. Consult a healthcare professional for personalized advice.

Generated with enthusiasm by a team passionate about genetics, powered by SNPs, code, and a lot of coffee!

"""
        if not categorized:
            report += "No matching genetic variants found."
        else:
            for category_key in sorted(categorized):
                original_category, _ = categorized[category_key][0]
                lower_category = original_category.lower()
                report += f"[{original_category}]\n" + "--"+("-" * len(original_category)) + \
                          f"\n\nWe found the following information in your raw DNA data file about your {lower_category}.\n"
                for _, line in categorized[category_key]:
                    report += line + "\n"
                report += "\n"

        report += "Thank you for exploring your DNA with us!"
        return report

    def save_report_as_pdf(self):
        report_text = self.Scrolledtext1.get("1.0", tk.END).strip()
        if not report_text or "Load a DNA file" in report_text:
            messagebox.showwarning("Warning", "No analysis report to save.")
            return
        pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save Report as PDF")
        if not pdf_path:
            return
        if not os.path.isdir(os.path.dirname(pdf_path)):
            messagebox.showerror("Error", "Invalid directory for saving PDF.")
            return
        try:
            c = canvas.Canvas(pdf_path, pagesize=A4)
            _width, height = A4
            margin = 40
            y = height - margin
            lines = report_text.split("\n")
            for i, line in enumerate(lines):
                if y < margin:
                    c.showPage()
                    y = height - margin
                is_heading = (
                    i + 1 < len(lines) and lines[i + 1].strip() == '-' * len(line.strip()) and line.isupper()
                )
                if is_heading:
                    y -= 10
                    c.setFont("Helvetica-Bold", 12)
                    c.drawString(margin, y, line)
                    y -= 2
                    c.line(margin, y, margin + c.stringWidth(line, "Helvetica-Bold", 12), y)
                    y -= 14
                    y -= 10
                elif line.strip() == '' or line.strip() == '-' * len(line.strip()):
                    y -= 10
                else:
                    c.setFont("Helvetica", 10)
                    wrapped = simpleSplit(line, "Helvetica", 10, _width - 2 * margin)
                    for wrap_line in wrapped:
                        c.drawString(margin, y, wrap_line)
                        y -= 14
            c.save()
            messagebox.showinfo("Success", f"Report saved as:\n{pdf_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save PDF:\n{e}")

    def match_rule(self, genotypes: dict[str, str], rule: str) -> bool:
        parts = rule.split(" AND ")
        for part in parts:
            if "=" not in part:
                return False
            rsid, expected = part.strip().split("=")
            actual = genotypes.get(rsid.strip(), "").upper()
            expected = "".join(sorted(expected.strip().upper()))
            actual = "".join(sorted(actual))
            if not actual or actual != expected:
                return False
        return True
    
    def detect_sex(self, genotypes: dict[str, str]) -> str:
        # Betrouwbare marker voor vrouwen: rs590787 (X-linked, wordt heterozygoot bij vrouwen)
        sex_marker = genotypes.get("rs590787", "")
        if sex_marker in {"CT", "TC"}:
            return "Female (XX)"
        
        # Tweede lijn: kijk naar Y-markers (alleen als homozygoot en geldig)
        y_markers = {"rs2032597", "rs768983", "rs390240", "rs390307", "rs1229982"}
        y_count = sum(1 for rsid in y_markers if genotypes.get(rsid, "") in {"AA", "CC", "GG", "TT"})
        if y_count >= 2:
            return "Male (XY)"
        
        # Default: onbekend of geen betrouwbare markers → vrouw aannemen om veilig te zijn
        return "Female (XX)"
    
    def toggle_rsID_genotype(self):
        self.include_rsid_genotype = not self.include_rsid_genotype
        if self.last_dna_lines:
            self.Scrolledtext1.delete("1.0", tk.END)
            self.Scrolledtext1.insert(tk.END, "Updating report, please wait...")
            self.root.update()
            report = self.analyze_dna(self.last_dna_lines)
            self.Scrolledtext1.delete("1.0", tk.END)
            self.Scrolledtext1.insert(tk.END, report)

def start_up():
    root = tk.Tk()
    app = Toplevel1(root)
    root.mainloop()

if __name__ == '__main__':
    start_up()