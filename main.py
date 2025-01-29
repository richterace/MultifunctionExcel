import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Processor App")
        self.root.geometry("600x400")
        
        self.filename = None
        self.df = None
        
        # UI Elements
        self.upload_btn = tk.Button(root, text="Upload Excel File", command=self.load_file)
        self.upload_btn.pack(pady=10)
        
        self.process_btn = tk.Button(root, text="Process Data", command=self.process_data, state=tk.DISABLED)
        self.process_btn.pack(pady=10)
        
        self.save_btn = tk.Button(root, text="Save Processed File", command=self.save_file, state=tk.DISABLED)
        self.save_btn.pack(pady=10)
        
        self.status_label = tk.Label(root, text="No file loaded", fg="red")
        self.status_label.pack(pady=10)
        
        self.tree = ttk.Treeview(root)
        self.tree.pack(expand=True, fill=tk.BOTH)
    
    def load_file(self):
        self.filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.filename:
            self.df = pd.read_excel(self.filename)
            self.display_data()
            self.process_btn.config(state=tk.NORMAL)
            self.status_label.config(text=f"Loaded: {self.filename}", fg="green")
    
    def display_data(self):
        if self.df is not None:
            self.tree.delete(*self.tree.get_children())
            self.tree["columns"] = list(self.df.columns)
            self.tree["show"] = "headings"
            
            for col in self.df.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100)
            
            for _, row in self.df.iterrows():
                self.tree.insert("", tk.END, values=list(row))
    
    def process_data(self):
        if self.df is not None:
            # Processing functions
            self.df.fillna(self.df.mean(numeric_only=True), inplace=True)  # Fill missing values with mean
            self.df.drop_duplicates(inplace=True)  # Remove duplicates
            
            for col in self.df.select_dtypes(include=["object"]).columns:
                self.df[col] = self.df[col].str.upper()  # Convert text to uppercase
            
            self.display_data()
            self.save_btn.config(state=tk.NORMAL)
            messagebox.showinfo("Success", "Data processed successfully!")
    
    def save_file(self):
        if self.df is not None:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if save_path:
                self.df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", "File saved successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
