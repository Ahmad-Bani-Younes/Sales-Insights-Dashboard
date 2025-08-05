import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Function to load and analyze the file
def load_and_analyze():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls")])
    if not file_path:
        return
    
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)

        if not {'Date', 'Product', 'Quantity', 'UnitPrice'}.issubset(df.columns):
            messagebox.showerror("Missing Columns", "File must contain columns: Date, Product, Quantity, UnitPrice")
            return
        
        df['Date'] = pd.to_datetime(df['Date'])
        df['Total'] = df['Quantity'] * df['UnitPrice']
        df['Month'] = df['Date'].dt.to_period('M')

        monthly_revenue = df.groupby('Month')['Total'].sum()
        top_products = df.groupby('Product')['Total'].sum().sort_values(ascending=False)
        avg_order_value = df.groupby('Product')['Total'].mean().round(2)

        # Clear previous treeview
        for item in tree.get_children():
            tree.delete(item)

        # Insert results into treeview
        tree.insert('', 'end', text="Monthly Revenue", values=('', ''))
        for index, value in monthly_revenue.items():
            tree.insert('', 'end', values=(str(index), f"${value:.2f}"))

        tree.insert('', 'end', text="Top Products", values=('', ''))
        for index, value in top_products.items():
            tree.insert('', 'end', values=(index, f"${value:.2f}"))

        tree.insert('', 'end', text="Average Order Value", values=('', ''))
        for index, value in avg_order_value.items():
            tree.insert('', 'end', values=(index, f"${value:.2f}"))

        # Show plots
        show_plot(monthly_revenue, "Monthly Revenue", "Month", "Revenue ($)")
        show_plot(top_products, "Top-Selling Products", "Product", "Total Sales ($)")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Function to show plots
def show_plot(data, title, xlabel, ylabel):
    fig, ax = plt.subplots(figsize=(5, 3))
    data.plot(kind='bar', ax=ax)
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)

    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=10)

# GUI setup
root = tk.Tk()
root.title("Sales Data Analyzer")
root.geometry("800x600")

load_button = tk.Button(root, text="Load CSV/Excel File", command=load_and_analyze)
load_button.pack(pady=10)

tree = ttk.Treeview(root, columns=("Name", "Value"), show="headings")
tree.heading("Name", text="Name")
tree.heading("Value", text="Value")
tree.pack(expand=True, fill='both')

root.mainloop()
