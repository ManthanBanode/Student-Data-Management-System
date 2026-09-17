import os
import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from openpyxl import Workbook

EXCEL_FILE = "student_results.xlsx"

def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Results"
        headers = [
            "Name", "Roll No.", "Class", 
            "Subject 1", "Subject 2", "Subject 3", "Subject 4", "Subject 5", 
            "Total Marks", "Percentage", "Result"
        ]
        ws.append(headers)
        wb.save(EXCEL_FILE)

class StudentResultApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Result Management System")
        self.root.geometry("750x550")
        self.root.resizable(False, False)

        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        title_label = tk.Label(
            root, 
            text="STUDENT RESULT MANAGEMENT SYSTEM", 
            font=("Helvetica", 16, "bold"), 
            bg="#2c3e50", 
            fg="white", 
            pady=12
        )
        title_label.pack(fill=tk.X)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_add = ttk.Frame(self.notebook)
        self.tab_get = ttk.Frame(self.notebook)
        self.tab_show_all = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_add, text=" 💾 Add Student ")
        self.notebook.add(self.tab_get, text=" 🔍 Get Result ")
        self.notebook.add(self.tab_show_all, text=" 📋 Show All Results ")

        self.setup_add_student_tab()
        self.setup_get_result_tab()
        self.setup_show_all_tab()

    def setup_add_student_tab(self):
        frame = ttk.LabelFrame(self.tab_add, text=" Enter Student & Marks Details ")
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        self.var_name = tk.StringVar()
        self.var_roll = tk.StringVar()
        self.var_class = tk.StringVar()
        self.var_sub1 = tk.StringVar()
        self.var_sub2 = tk.StringVar()
        self.var_sub3 = tk.StringVar()
        self.var_sub4 = tk.StringVar()
        self.var_sub5 = tk.StringVar()

        labels_vars = [
            ("Name:", self.var_name),
            ("Roll No.:", self.var_roll),
            ("Class:", self.var_class),
            ("Subject 1 Marks:", self.var_sub1),
            ("Subject 2 Marks:", self.var_sub2),
            ("Subject 3 Marks:", self.var_sub3),
            ("Subject 4 Marks:", self.var_sub4),
            ("Subject 5 Marks:", self.var_sub5),
        ]

        for i, (label_text, var) in enumerate(labels_vars):
            row = i % 4
            col = (i // 4) * 2
            
            lbl = ttk.Label(frame, text=label_text, font=("Arial", 10))
            lbl.grid(row=row, column=col, sticky=tk.W, padx=15, pady=12)
            
            entry = ttk.Entry(frame, textvariable=var, font=("Arial", 10), width=18)
            entry.grid(row=row, column=col+1, padx=10, pady=12)

        save_btn = tk.Button(
            frame, text="💾 Save", font=("Arial", 11, "bold"), 
            bg="#27ae60", fg="white", width=15, command=self.save_student
        )
        save_btn.grid(row=4, column=0, columnspan=4, pady=20)

    def save_student(self):
        name = self.var_name.get().strip()
        roll = self.var_roll.get().strip()
        cls_name = self.var_class.get().strip()
        subs = [
            self.var_sub1.get().strip(),
            self.var_sub2.get().strip(),
            self.var_sub3.get().strip(),
            self.var_sub4.get().strip(),
            self.var_sub5.get().strip()
        ]

        if not name or not roll or not cls_name or any(s == "" for s in subs):
            messagebox.showerror("Error", "All fields are required!")
            return

        try:
            marks = [float(s) for s in subs]
            for m in marks:
                if m < 0 or m > 100:
                    messagebox.showerror("Error", "Marks should be between 0 and 100!")
                    return
        except ValueError:
            messagebox.showerror("Error", "Marks and Roll No. must be numeric!")
            return

        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if str(row[1]) == str(roll):
                messagebox.showerror("Error", f"Roll No. {roll} already exists!")
                wb.close()
                return

        total = sum(marks)
        percentage = round(total / 5.0, 2)
        result = "Pass" if all(m >= 33 for m in marks) else "Fail"

        record = [
            name, roll, cls_name, 
            marks[0], marks[1], marks[2], marks[3], marks[4],
            total, f"{percentage}%", result
        ]

        ws.append(record)
        wb.save(EXCEL_FILE)
        wb.close()

        messagebox.showinfo("Success", f"Student Record for {name} saved successfully!")
        self.clear_add_form()

    def clear_add_form(self):
        self.var_name.set("")
        self.var_roll.set("")
        self.var_class.set("")
        self.var_sub1.set("")
        self.var_sub2.set("")
        self.var_sub3.set("")
        self.var_sub4.set("")
        self.var_sub5.set("")

    def setup_get_result_tab(self):
        frame = ttk.LabelFrame(self.tab_get, text=" Search Student Result ")
        frame.pack(fill=tk.X, padx=15, pady=15)

        ttk.Label(frame, text="Enter Roll No.:", font=("Arial", 11)).pack(side=tk.LEFT, padx=10, pady=15)
        
        self.var_search_roll = tk.StringVar()
        search_entry = ttk.Entry(frame, textvariable=self.var_search_roll, font=("Arial", 11), width=15)
        search_entry.pack(side=tk.LEFT, padx=10, pady=15)

        search_btn = tk.Button(
            frame, text="🔍 Get Result", font=("Arial", 10, "bold"), 
            bg="#2980b9", fg="white", command=self.get_result
        )
        search_btn.pack(side=tk.LEFT, padx=10, pady=15)

        cols = ("Name", "Roll No.", "Class", "Total Marks", "Percentage", "Result")
        self.result_tree = ttk.Treeview(self.tab_get, columns=cols, show="headings", height=3)
        self.result_tree.pack(fill=tk.X, padx=15, pady=10)

        for col in cols:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, anchor=tk.CENTER, width=110)

    def get_result(self):
        roll_query = self.var_search_roll.get().strip()
        if not roll_query:
            messagebox.showwarning("Warning", "Please enter a Roll No. to search!")
            return

        for row in self.result_tree.get_children():
            self.result_tree.delete(row)

        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active

        found = False
        for row in ws.iter_rows(min_row=2, values_only=True):
            if str(row[1]) == roll_query:
                data_to_show = (row[0], row[1], row[2], row[8], row[9], row[10])
                self.result_tree.insert("", tk.END, values=data_to_show)
                found = True
                break

        wb.close()

        if not found:
            messagebox.showerror("Not Found", "❌ Student record not found.")

    def setup_show_all_tab(self):
        top_frame = ttk.Frame(self.tab_show_all)
        top_frame.pack(fill=tk.X, padx=15, pady=10)

        show_btn = tk.Button(
            top_frame, text="📋 Show All Results", font=("Arial", 11, "bold"), 
            bg="#8e44ad", fg="white", command=self.load_all_results
        )
        show_btn.pack(side=tk.LEFT)

        tree_frame = ttk.Frame(self.tab_show_all)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        cols = ("Name", "Roll No.", "Class", "Total Marks", "Percentage", "Result")
        self.all_tree = ttk.Treeview(tree_frame, columns=cols, show="headings")

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.all_tree.yview)
        self.all_tree.configure(yscroll=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.all_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for col in cols:
            self.all_tree.heading(col, text=col)
            self.all_tree.column(col, anchor=tk.CENTER, width=110)

    def load_all_results(self):
        for row in self.all_tree.get_children():
            self.all_tree.delete(row)

        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active

        for row in ws.iter_rows(min_row=2, values_only=True):
            data_to_show = (row[0], row[1], row[2], row[8], row[9], row[10])
            self.all_tree.insert("", tk.END, values=data_to_show)

        wb.close()

if __name__ == "__main__":
    initialize_excel()
    root = tk.Tk()
    app = StudentResultApp(root)
    root.mainloop()