import pandas as pd

class ExcelPivotProcessor():
    def __init__(self, root):
        super().__init__(root)
        self.pivot_btn = tk.Button(root, text="Generate Pivot Table", command=self.generate_pivot, state=tk.DISABLED)
        self.pivot_btn.pack(pady=10)
    
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

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelPivotProcessor(root)
    root.mainloop()
