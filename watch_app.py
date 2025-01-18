import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os
import openpyxl
from openpyxl.styles import Font

class WatchPricingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Watch Battery Price Manager")
        self.root.geometry("800x600")
        
        # Set app style
        self.setup_styles()
        
        # Initialize data storage
        self.excel_file = os.path.join(os.path.expanduser("~"), "Documents", "watch_pricing.xlsx")
        self.setup_excel_file()
        
        # Create main container
        self.create_main_interface()
        
        # Load existing data
        self.load_existing_data()

    def setup_styles(self):
        style = ttk.Style()
        style.configure("Header.TLabel", font=('Arial', 16, 'bold'))
        style.configure("Status.TLabel", font=('Arial', 10))
        
    def setup_excel_file(self):
        if not os.path.exists(self.excel_file):
            wb = openpyxl.Workbook()
            for i in range(1, 6):
                sheet = wb.create_sheet(f"Category {i}")
                headers = ["Brand", "Price", "Date Added"]
                sheet.append(headers)
                # Make headers bold
                for col in range(1, len(headers) + 1):
                    sheet.cell(row=1, column=col).font = Font(bold=True)
            # Remove default sheet
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])
            wb.save(self.excel_file)

    def create_main_interface(self):
        # Header
        header = ttk.Label(self.root, text="Watch Battery Price Manager", style="Header.TLabel")
        header.pack(pady=20)

        # Input Frame
        input_frame = ttk.LabelFrame(self.root, text="Add New Watch", padding=10)
        input_frame.pack(fill=tk.X, padx=20, pady=10)

        # Brand Entry
        brand_frame = ttk.Frame(input_frame)
        brand_frame.pack(fill=tk.X, pady=5)
        ttk.Label(brand_frame, text="Watch Brand:").pack(side=tk.LEFT)
        self.brand_entry = ttk.Entry(brand_frame, width=40)
        self.brand_entry.pack(side=tk.LEFT, padx=5)

        # Price Entry
        price_frame = ttk.Frame(input_frame)
        price_frame.pack(fill=tk.X, pady=5)
        ttk.Label(price_frame, text="Price ($):").pack(side=tk.LEFT)
        self.price_entry = ttk.Entry(price_frame, width=20)
        self.price_entry.pack(side=tk.LEFT, padx=5)

        # Add Button
        ttk.Button(input_frame, text="Add Watch", command=self.add_watch).pack(pady=10)

        # Treeview for displaying watches
        tree_frame = ttk.Frame(self.root)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.tree = ttk.Treeview(tree_frame, columns=("Brand", "Price", "Category", "Date"), show="headings")
        self.tree.heading("Brand", text="Brand")
        self.tree.heading("Price", text="Price")
        self.tree.heading("Category", text="Category")
        self.tree.heading("Date", text="Date Added")

        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def categorize_watch(self, price):
        price = float(price)
        if price <= 48.95: return 5
        elif price <= 68.95: return 4
        elif price <= 124.95: return 3
        elif price <= 168.00: return 2
        else: return 1

    def add_watch(self):
        try:
            brand = self.brand_entry.get().strip()
            price = self.price_entry.get().strip()

            if not brand or not price:
                messagebox.showerror("Error", "Please fill in all fields")
                return

            try:
                price_float = float(price)
                if price_float <= 0:
                    messagebox.showerror("Error", "Price must be greater than 0")
                    return
            except ValueError:
                messagebox.showerror("Error", "Invalid price format")
                return

            category = self.categorize_watch(price)
            date_added = datetime.now().strftime("%Y-%m-%d %H:%M")

            # Add to treeview
            self.tree.insert("", tk.END, values=(brand, f"${price}", f"Category {category}", date_added))

            # Add to Excel
            wb = openpyxl.load_workbook(self.excel_file)
            sheet = wb[f"Category {category}"]
            sheet.append([brand, float(price), date_added])
            wb.save(self.excel_file)

            # Clear entries
            self.brand_entry.delete(0, tk.END)
            self.price_entry.delete(0, tk.END)
            self.brand_entry.focus()

            messagebox.showinfo("Success", f"Added {brand} to Category {category}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def load_existing_data(self):
        if os.path.exists(self.excel_file):
            wb = openpyxl.load_workbook(self.excel_file)
            for category in range(1, 6):
                sheet = wb[f"Category {category}"]
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if row[0]:  # Check if brand exists
                        self.tree.insert("", tk.END, values=(
                            row[0],  # Brand
                            f"${row[1]}" if row[1] else "",  # Price
                            f"Category {category}",
                            row[2] if row[2] else ""  # Date
                        ))

def main():
    root = tk.Tk()
    app = WatchPricingApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()