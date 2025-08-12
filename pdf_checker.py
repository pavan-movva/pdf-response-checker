import re
import pdfplumber
import fitz   # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import os

# ---------------- PDF Parsing Functions ---------------- #

def extract_responses(pdf_path):
    """Return dict: qid -> set(['1','2',...])"""
    responses = {}
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join((page.extract_text() or "") for page in pdf.pages)
    matches = re.findall(r"Question ID\s*:\s*(\d+)[\s\S]*?Chosen Option\s*:\s*([0-9,\|\s]+)", text, flags=re.IGNORECASE)
    for qid, opttext in matches:
        opts = set(re.findall(r"\d+", opttext))
        responses[qid] = opts
    return responses

def extract_answerkey_with_colors(pdf_path):
    """Extract correct answers (green text) and ambiguous QIDs from answer key."""
    doc = fitz.open(pdf_path)
    answerkey = {}
    ambiguous = set()
    current_qid = None

    for page in doc:
        blocks = page.get_text("dict").get("blocks", [])
        for b in blocks:
            for line in b.get("lines", []):
                line_text = "".join(s.get("text", "") for s in line.get("spans", [])).strip()

                # Match Question ID
                m_q = re.search(r"Question\s+Id\s*:\s*(\d+)", line_text, flags=re.IGNORECASE)
                if m_q:
                    current_qid = m_q.group(1)
                    answerkey.setdefault(current_qid, set())

                # Detect ambiguous note
                if current_qid and ("ambigu" in line_text.lower() or "candidate will get full marks" in line_text.lower()):
                    ambiguous.add(current_qid)

                # Detect correct options by color
                for s in line.get("spans", []):
                    span_text = s.get("text", "").strip()
                    m_opt = re.match(r"^(\d+)\.\s*", span_text)
                    if m_opt and current_qid:
                        opt_num = m_opt.group(1)
                        color = s.get("color")
                        r = g = b = 0
                        try:
                            if isinstance(color, (int, float)):
                                color_int = int(color)
                                r = (color_int >> 16) & 255
                                g = (color_int >> 8) & 255
                                b = color_int & 255
                            elif isinstance(color, (list, tuple)) and len(color) >= 3:
                                r = int(color[0] * 255)
                                g = int(color[1] * 255)
                                b = int(color[2] * 255)
                        except:
                            r = g = b = 0
                        if (g > 120) and (g > r + 30) and (g > b + 30):
                            answerkey[current_qid].add(opt_num)

    doc.close()
    return answerkey, ambiguous

def calculate_score(responses, answerkey, ambiguous):
    """Return counts and details list."""
    correct = 0
    wrong = 0
    details = []  # (QID, chosen_set, correct_set, result)

    for qid, chosen in responses.items():
        if qid not in answerkey:
            details.append((qid, chosen, set(), "no-key"))
            wrong += 1
            continue

        correct_set = answerkey[qid]
        chosen_set = set(chosen)

        if len(correct_set) == 1:
            if chosen_set == correct_set:
                correct += 1
                details.append((qid, chosen_set, correct_set, "correct"))
            else:
                wrong += 1
                details.append((qid, chosen_set, correct_set, "wrong"))
        else:
            if qid in ambiguous:
                if chosen_set & correct_set:
                    correct += 1
                    details.append((qid, chosen_set, correct_set, "correct (ambiguous)"))
                else:
                    wrong += 1
                    details.append((qid, chosen_set, correct_set, "wrong (ambiguous)"))
            else:
                if chosen_set == correct_set:
                    correct += 1
                    details.append((qid, chosen_set, correct_set, "correct (multi)"))
                else:
                    wrong += 1
                    details.append((qid, chosen_set, correct_set, "wrong (multi)"))

    return correct, wrong, details

# ---------------- Excel Export ---------------- #

def save_to_excel(details, correct, wrong, final_score, save_dir):
    # Ensure folder exists
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    # Create unique filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(save_dir, f"results_{timestamp}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"

    # Header
    headers = ["Question ID", "Chosen Options", "Correct Options", "Result"]
    ws.append(headers)
    for col in ws.iter_cols(min_row=1, max_row=1, max_col=len(headers)):
        for cell in col:
            cell.font = Font(bold=True)

    # Rows
    for qid, chosen, correct_set, result in details:
        ws.append([
            qid,
            ", ".join(sorted(chosen)) if chosen else "",
            ", ".join(sorted(correct_set)) if correct_set else "",
            result
        ])

    # Summary row
    ws.append([])
    ws.append(["Summary"])
    ws.append(["Correct Answers", correct])
    ws.append(["Wrong Answers", wrong])
    ws.append(["Final Score (Correct/2)", final_score])

    wb.save(filename)
    return filename

# ---------------- GUI ---------------- #

class PDFCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Response Sheet & Answer Key Checker")

        self.response_path = None
        self.answerkey_path = None

        tk.Button(root, text="Select Response Sheet PDF", command=self.load_response_pdf).pack(pady=5)
        tk.Button(root, text="Select Answer Key PDF", command=self.load_answerkey_pdf).pack(pady=5)
        tk.Button(root, text="Process & Save to Excel", command=self.process_files).pack(pady=10)

        self.result_area = scrolledtext.ScrolledText(root, width=100, height=30)
        self.result_area.pack(padx=10, pady=10)

    def load_response_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.response_path = path
            messagebox.showinfo("File Selected", f"Response Sheet loaded:\n{path}")

    def load_answerkey_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.answerkey_path = path
            messagebox.showinfo("File Selected", f"Answer Key loaded:\n{path}")

    def process_files(self):
        if not self.response_path or not self.answerkey_path:
            messagebox.showerror("Error", "Please select both PDF files.")
            return

        self.result_area.delete(1.0, tk.END)
        try:
            responses = extract_responses(self.response_path)
            answerkey, ambiguous = extract_answerkey_with_colors(self.answerkey_path)
            correct, wrong, details = calculate_score(responses, answerkey, ambiguous)

            final_score = correct / 2  # Final score calculation

            # Save in same folder as response sheet
            save_dir = os.path.dirname(self.response_path)
            excel_file = save_to_excel(details, correct, wrong, final_score, save_dir)

            # Display in GUI
            self.result_area.insert(tk.END, f"Correct answers: {correct}\n")
            self.result_area.insert(tk.END, f"Wrong answers: {wrong}\n")
            self.result_area.insert(tk.END, f"Final Score (Correct/2): {final_score}\n")
            self.result_area.insert(tk.END, f"Results saved to: {excel_file}\n\n")
            self.result_area.insert(tk.END, "=== All Results ===\n")
            for qid, chosen, corr, result in details:
                self.result_area.insert(tk.END, f"QID: {qid} | Chosen: {sorted(chosen)} | Correct: {sorted(corr)} | Result: {result}\n")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

# ---------------- Main ---------------- #

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCheckerApp(root)
    root.mainloop()
