import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Processor App with Pivot Table")
        
        self.filename = None
        self.df = None
        
        self.upload_btn = tk.Button(root, text="Upload Excel File", command=self.load_file)
        self.upload_btn.pack(pady=10)
        
        self.process_btn = tk.Button(root, text="Process Data", command=self.process_data, state=tk.DISABLED)
        self.process_btn.pack(pady=10)
        
        self.save_btn = tk.Button(root, text="Save Processed File", command=self.save_file, state=tk.DISABLED)
        self.save_btn.pack(pady=10)
        
        self.pivot_btn = tk.Button(root, text="Generate Pivot Table", command=self.generate_pivot, state=tk.DISABLED)
        self.pivot_btn.pack(pady=10)
        
        self.stats_btn = tk.Button(root, text="Calculate Stats", command=self.calculate_stats, state=tk.DISABLED)
        self.stats_btn.pack(pady=10)
        
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
            self.pivot_btn.config(state=tk.NORMAL)
            self.stats_btn.config(state=tk.NORMAL)
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
            self.df.fillna(self.df.mean(numeric_only=True), inplace=True)
            self.df.drop_duplicates(inplace=True)
            
            for col in self.df.select_dtypes(include=["object"]).columns:
                self.df[col] = self.df[col].str.upper()
            
            self.display_data()
            self.save_btn.config(state=tk.NORMAL)
            messagebox.showinfo("Success", "Data processed successfully!")
    
    def save_file(self):
        if self.df is not None:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if save_path:
                self.df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", "File saved successfully!")
    
    def generate_pivot(self):
        if self.df is not None:
            try:
                pivot_table = self.df.pivot_table(index=self.df.columns[0], aggfunc='sum')
                self.df = pivot_table.reset_index()
                self.display_data()
                messagebox.showinfo("Success", "Pivot table generated successfully!")
                self.save_btn.config(state=tk.NORMAL)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to generate pivot table: {e}")
    
    def calculate_stats(self):
        if self.df is not None:
            try:
                stats = {
                    'Mean': self.df.mean(numeric_only=True),
                    'Median': self.df.median(numeric_only=True),
                    'Mode': self.df.mode().iloc[0]
                }
                stats_df = pd.DataFrame(stats)
                messagebox.showinfo("Statistics", stats_df.to_string())
            except Exception as e:
                messagebox.showerror("Error", f"Failed to calculate statistics: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
