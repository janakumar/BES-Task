import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

class FileProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to Excel Processor")
        
        self.data1_path = ""
        self.desc_path = ""
        self.boq_path = ""
        
        tk.Label(root, text="Select Data1 CSV:").grid(row=0, column=0)
        self.data1_btn = tk.Button(root, text="Browse", command=self.load_data1)
        self.data1_btn.grid(row=0, column=1)
        
        tk.Label(root, text="Select Description CSV:").grid(row=1, column=0)
        self.desc_btn = tk.Button(root, text="Browse", command=self.load_desc)
        self.desc_btn.grid(row=1, column=1)
        
        tk.Label(root, text="Select BOQ Excel File:").grid(row=2, column=0)
        self.boq_btn = tk.Button(root, text="Browse", command=self.load_boq)
        self.boq_btn.grid(row=2, column=1)
        
        self.process_btn = tk.Button(root, text="Process and Export", command=self.process_files)
        self.process_btn.grid(row=3, column=0, columnspan=2)
    
    def load_data1(self):
        self.data1_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    
    def load_desc(self):
        self.desc_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    
    def load_boq(self):
        self.boq_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    
    def process_files(self):
        if not self.data1_path or not self.desc_path or not self.boq_path:
            messagebox.showerror("Error", "Please select all files")
            return
        
        try:
            # Load data
            boq_df = pd.read_excel(self.boq_path)
            data1_df = pd.read_csv(self.data1_path)
            desc_df = pd.read_csv(self.desc_path)

            # Standardize column names (strip spaces)
            boq_df.columns = boq_df.columns.str.strip()
            data1_df.columns = data1_df.columns.str.strip()
            desc_df.columns = desc_df.columns.str.strip()

            # Rename columns in Data1.csv to match BOQ for merging
            data1_df.rename(columns={"RM (m)": "RMT"}, inplace=True)

            # Rename columns in Description.csv for merging
            desc_df.rename(columns={"Description": "Brief Description", "Cost": "Rate (INR)"}, inplace=True)

            # Merge BOQ and Data1.csv on "Facade Type" & "Sub Type"
            merged_df = pd.merge(boq_df, data1_df, on=["Facade Type", "Sub Type"], how="outer", sort=False)

            # Merge the result with Description.csv on "Facade Type" & "Sub Type"
            merged_df = pd.merge(merged_df, desc_df, on=["Facade Type", "Sub Type"], how="outer", sort=False)

            # Find common values of 'Facade Type' and 'Sub Type' in all three DataFrames
            common_facade_types = set(boq_df["Facade Type"]) & set(data1_df["Facade Type"]) & set(desc_df["Facade Type"])
            common_sub_types = set(boq_df["Sub Type"]) & set(data1_df["Sub Type"]) & set(desc_df["Sub Type"])

            # Filter merged_df based on these common values
            filtered_df = merged_df[(merged_df["Facade Type"].isin(common_facade_types)) & (merged_df["Sub Type"].isin(common_sub_types))]
            filtered_df = filtered_df.dropna(subset=["Facade Type", "Sub Type"])

            # Remove empty columns
            filtered_df = filtered_df.dropna(axis=1, how='all')

            # Save to Excel
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            filtered_df.to_excel(output_file, index=False)
            
            messagebox.showinfo("Success", "File exported successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = FileProcessor(root)
    root.mainloop()