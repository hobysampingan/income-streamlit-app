import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
import json

class IncomeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Income & Orders Summary with Cost Management")
        self.pesanan_data = None
        self.income_data = None
        self.merged_data = None
        self.cost_data = {}  # Dictionary to store cost per product
        self.cost_file = "product_costs.json"
        
        # Load existing cost data if available
        self.load_cost_data()
        
        # Create main frame with notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Main tab
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text="Main Analysis")
        
        # Cost management tab
        self.cost_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.cost_frame, text="Cost Management")
        
        self.setup_main_tab()
        self.setup_cost_tab()

    def setup_main_tab(self):
        # Buttons to load files and summary actions
        button_frame = ttk.Frame(self.main_frame, padding=10)
        button_frame.pack(fill="x")

        ttk.Button(button_frame, text="Load Pesanan Selesai", command=self.load_pesanan_file).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Load Income", command=self.load_income_file).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Process & Summarize", command=self.process_data).pack(side="left", padx=20)
        ttk.Button(button_frame, text="Show Overall Totals", command=self.show_overall_totals).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Show Profit Analysis", command=self.show_profit_analysis).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Export to Excel", command=self.export_to_excel).pack(side="left", padx=5)

        # Treeview for summary includes SKU, Product Name, and Profit info
        cols = ("Seller SKU", "Product Name", "Variation", "TotalQty", "Revenue", "Cost", "Profit", "Share 60%", "Share 40%")
        self.summary_tree = ttk.Treeview(self.main_frame, columns=cols, show="headings")
        
        # Set column widths
        widths = [120, 150, 120, 80, 100, 100, 100, 100, 100]
        for i, col in enumerate(cols):
            self.summary_tree.heading(col, text=col)
            self.summary_tree.column(col, width=widths[i])
        
        self.summary_tree.pack(fill="both", expand=True, padx=10, pady=10)
        self.summary_tree.bind("<Double-1>", self.on_summary_double_click)

    def on_summary_double_click(self, event):
        selected = self.summary_tree.selection()
        if not selected:
            return
        values = self.summary_tree.item(selected[0])['values']
        detail = (
            f"SKU: {values[0]}\n"
            f"Product Name: {values[1]}\n"
            f"Variation: {values[2]}\n"
            f"Total Quantity: {values[3]}\n"
            f"Revenue: {values[4]}\n"
            f"Cost: {values[5]}\n"
            f"Profit: {values[6]}\n"
            f"Share 60%: {values[7]}\n"
            f"Share 40%: {values[8]}"
        )
        messagebox.showinfo("Detail Produk", detail)

    def setup_cost_tab(self):
        frame = ttk.Frame(self.cost_frame, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Kelola Harga Modal Produk", font=("Arial", 14, "bold")).pack(pady=10)
        ttk.Label(frame, text="Pilih produk dari dropdown lalu masukkan harga modal per unit.", wraplength=500).pack(pady=5)

        input_frame = ttk.LabelFrame(frame, text="Tambah/Edit Harga Modal", padding=10)
        input_frame.pack(fill="x", pady=10)

        ttk.Label(input_frame, text="Nama Produk:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.product_name_var = tk.StringVar()
        self.product_combo = ttk.Combobox(input_frame, textvariable=self.product_name_var, width=40, state='readonly')
        self.product_combo.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(input_frame, text="Harga Modal:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.cost_var = tk.StringVar()
        self.cost_entry = ttk.Entry(input_frame, textvariable=self.cost_var, width=15)
        self.cost_entry.grid(row=0, column=3, padx=5, pady=5)

        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=1, column=0, columnspan=4, pady=10)
        ttk.Button(btn_frame, text="Simpan", command=self.save_cost).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Hapus", command=self.delete_cost).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Clear", command=self.clear_cost_inputs).pack(side="left", padx=5)

        list_frame = ttk.LabelFrame(frame, text="Daftar Harga Modal", padding=10)
        list_frame.pack(fill="both", expand=True, pady=10)
        cost_cols = ("Product Name", "Cost per Unit")
        self.cost_tree = ttk.Treeview(list_frame, columns=cost_cols, show="headings", height=15)
        self.cost_tree.heading("Product Name", text="Nama Produk")
        self.cost_tree.heading("Cost per Unit", text="Harga Modal per Unit")
        self.cost_tree.column("Product Name", width=300)
        self.cost_tree.column("Cost per Unit", width=150)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.cost_tree.yview)
        self.cost_tree.configure(yscroll=scrollbar.set)
        self.cost_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.cost_tree.bind("<<TreeviewSelect>>", self.on_cost_select)
        self.refresh_cost_list()

    def load_cost_data(self):
        try:
            if os.path.exists(self.cost_file):
                with open(self.cost_file, 'r', encoding='utf-8') as f:
                    self.cost_data = json.load(f)
        except:
            self.cost_data = {}

    def save_cost_data(self):
        with open(self.cost_file, 'w', encoding='utf-8') as f:
            json.dump(self.cost_data, f, ensure_ascii=False, indent=2)

    # New method to get product cost, preventing attribute errors
    def get_product_cost(self, product_name):
        return float(self.cost_data.get(product_name, 0.0))

    def save_cost(self):
        prod = self.product_name_var.get().strip()
        cost_str = self.cost_var.get().strip()
        if not prod or not cost_str:
            messagebox.showwarning("Warning", "Produk dan harga harus diisi!")
            return
        try:
            cost = float(cost_str)
            if cost < 0:
                raise ValueError
            self.cost_data[prod] = cost
            self.save_cost_data()
            self.refresh_cost_list()
            self.clear_cost_inputs()
            messagebox.showinfo("Sukses", f"Harga modal untuk '{prod}' disimpan.")
        except:
            messagebox.showerror("Error", "Harga modal harus angka positif.")

    def delete_cost(self):
        sel = self.cost_tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Pilih produk terlebih dahulu.")
            return
        prod = self.cost_tree.item(sel[0])['values'][0]
        if messagebox.askyesno("Confirm", f"Hapus harga modal '{prod}'?"):
            self.cost_data.pop(prod, None)
            self.save_cost_data()
            self.refresh_cost_list()
            self.clear_cost_inputs()

    def clear_cost_inputs(self):
        self.product_name_var.set("")
        self.cost_var.set("")

    def on_cost_select(self, event):
        sel = self.cost_tree.selection()
        if sel:
            prod, cost = self.cost_tree.item(sel[0])['values']
            self.product_name_var.set(prod)
            self.cost_var.set(str(cost))

    def refresh_cost_list(self):
        for i in self.cost_tree.get_children():
            self.cost_tree.delete(i)
        for prod, cost in sorted(self.cost_data.items()):
            fmt = f"{cost:,.0f}".replace(',', '.')
            self.cost_tree.insert('', 'end', values=(prod, fmt))
        # Update dropdown values
        self.update_product_list()

    def update_product_list(self):
        if self.pesanan_data is not None:
            products = sorted(self.pesanan_data['Product Name'].astype(str).unique())
        else:
            products = []
        self.product_combo['values'] = products

    def load_pesanan_file(self):
        fname = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not fname: return
        try:
            df = pd.read_excel(fname, header=0, skiprows=[1])
            df.columns = df.columns.str.strip()
            self.pesanan_data = df
            messagebox.showinfo("Loaded", f"Pesanan ({len(df)} baris) berhasil di-load.")
            self.update_product_list()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def load_income_file(self):
        fname = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not fname: return
        try:
            df = pd.read_excel(fname)
            df.columns = df.columns.str.strip()
            self.income_data = df
            messagebox.showinfo("Loaded", f"Income ({len(df)} baris) berhasil di-load.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def process_data(self):
        if self.pesanan_data is None or self.income_data is None:
            messagebox.showwarning("Warning", "Load kedua file terlebih dahulu!")
            return
        df1 = self.pesanan_data[self.pesanan_data['Order Status']=='Selesai']
        df2 = self.income_data.drop_duplicates(subset=['Order/adjustment ID'])
        merged = pd.merge(df1, df2, left_on='Order ID', right_on='Order/adjustment ID', how='inner')
        if merged.empty:
            messagebox.showwarning("Warning", "Tidak ada data cocok.")
            return
        self.merged_data = merged.copy()
        summary = merged.groupby(['Seller SKU','Product Name','Variation'], as_index=False).agg(
            TotalQty=('Quantity','sum'), Revenue=('Total settlement amount','sum')
        )
        summary['Cost per Unit'] = summary['Product Name'].apply(self.get_product_cost)
        summary['Total Cost'] = summary['TotalQty']*summary['Cost per Unit']
        summary['Profit'] = summary['Revenue']-summary['Total Cost']
        summary['Share 60%'] = summary['Profit']*0.6
        summary['Share 40%'] = summary['Profit']*0.4
        for r in self.summary_tree.get_children():
            self.summary_tree.delete(r)
        for _,r in summary.iterrows():
            fmt_vals = [
                r['Seller SKU'], r['Product Name'], r['Variation'],
                f"{r['TotalQty']:,}".replace(',','.'),
                f"{r['Revenue']:,.0f}".replace(',','.'),
                f"{r['Total Cost']:,.0f}".replace(',','.'),
                f"{r['Profit']:,.0f}".replace(',','.'),
                f"{r['Share 60%']:,.0f}".replace(',','.'),
                f"{r['Share 40%']:,.0f}".replace(',','.')
            ]
            self.summary_tree.insert('', 'end', values=fmt_vals)
        messagebox.showinfo("Done", "Summary siap. Double-click untuk detail.")

    def show_overall_totals(self):
        if self.merged_data is None:
            messagebox.showwarning("Warning", "Please process data first.")
            return

        # Calculate summary by SKU
        summary = (
            self.merged_data.groupby('Seller SKU', as_index=False)
            .agg({
                'Quantity': 'sum',
                'Total settlement amount': 'sum'
            })
            .rename(columns={
                'Quantity': 'TotalQty',
                'Total settlement amount': 'Revenue'
            })
        )
        
        # Get unique product names for each SKU to calculate costs
        sku_products = self.merged_data.groupby('Seller SKU')['Product Name'].first().to_dict()
        
        # Add cost calculations
        summary['Cost per Unit'] = summary['Seller SKU'].map(
            lambda sku: self.get_product_cost(sku_products.get(sku, ''))
        )
        summary['Total Cost'] = summary['TotalQty'] * summary['Cost per Unit']
        summary['Profit'] = summary['Revenue'] - summary['Total Cost']

        # Calculate totals
        unique_orders = self.merged_data.drop_duplicates(subset=['Order ID'])
        total_orders = unique_orders['Order ID'].nunique()
        total_revenue = unique_orders['Total settlement amount'].sum()
        total_cost = summary['Total Cost'].sum()
        total_profit = total_revenue - total_cost
        total_share_60 = total_profit * 0.6
        total_share_40 = total_profit * 0.4

        # Format detail lines
        detail_lines = []
        for _, row in summary.iterrows():
            sku = row['Seller SKU']
            qty_fmt = f"{row['TotalQty']:,}".replace(',', '.')
            revenue_fmt = f"{row['Revenue']:,.0f}".replace(',', '.')
            cost_fmt = f"{row['Total Cost']:,.0f}".replace(',', '.')
            profit_fmt = f"{row['Profit']:,.0f}".replace(',', '.')
            
            detail_lines.append(
                f"{sku}: {qty_fmt} pcs | Revenue: {revenue_fmt} | Cost: {cost_fmt} | Profit: {profit_fmt}"
            )

        # Format totals
        total_revenue_fmt = f"{total_revenue:,.0f}".replace(',', '.')
        total_cost_fmt = f"{total_cost:,.0f}".replace(',', '.')
        total_profit_fmt = f"{total_profit:,.0f}".replace(',', '.')
        total_share_60_fmt = f"{total_share_60:,.0f}".replace(',', '.')
        total_share_40_fmt = f"{total_share_40:,.0f}".replace(',', '.')

        summary_text = '\n'.join(detail_lines)
        summary_text += f"\n\n=== SUMMARY ===\n"
        summary_text += f"Total Orders: {total_orders}\n"
        summary_text += f"Total Revenue: {total_revenue_fmt}\n"
        summary_text += f"Total Cost: {total_cost_fmt}\n"
        summary_text += f"Total Profit: {total_profit_fmt}\n"
        summary_text += f"Share 60%: {total_share_60_fmt}\n"
        summary_text += f"Share 40%: {total_share_40_fmt}"

        messagebox.showinfo("Overall Totals with Profit Analysis", summary_text)

    def show_profit_analysis(self):
        """Show detailed profit analysis"""
        if self.merged_data is None:
            messagebox.showwarning("Warning", "Please process data first.")
            return
        
        # Create profit analysis window
        profit_win = tk.Toplevel(self.root)
        profit_win.title("Profit Analysis")
        profit_win.geometry("800x600")
        
        # Create notebook for different views
        profit_notebook = ttk.Notebook(profit_win)
        profit_notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Summary tab
        summary_frame = ttk.Frame(profit_notebook)
        profit_notebook.add(summary_frame, text="Summary")
        
        # Calculate profit summary
        summary = (
            self.merged_data.groupby(['Seller SKU', 'Product Name', 'Variation'], as_index=False)
            .agg({
                'Quantity': 'sum',
                'Total settlement amount': 'sum'
            })
            .rename(columns={
                'Quantity': 'TotalQty',
                'Total settlement amount': 'Revenue'
            })
        )
        
        summary['Cost per Unit'] = summary['Product Name'].apply(self.get_product_cost)
        summary['Total Cost'] = summary['TotalQty'] * summary['Cost per Unit']
        summary['Profit'] = summary['Revenue'] - summary['Total Cost']
        summary['Profit Margin %'] = (summary['Profit'] / summary['Revenue'] * 100).round(2)
        
        # Create treeview for profit summary
        profit_cols = ("SKU", "Product", "Variation", "Qty", "Revenue", "Cost", "Profit", "Margin %")
        profit_tree = ttk.Treeview(summary_frame, columns=profit_cols, show="headings")
        
        for col in profit_cols:
            profit_tree.heading(col, text=col)
            profit_tree.column(col, width=90)
        
        # Scrollbar
        profit_scrollbar = ttk.Scrollbar(summary_frame, orient="vertical", command=profit_tree.yview)
        profit_tree.configure(yscroll=profit_scrollbar.set)
        
        profit_tree.pack(side="left", fill="both", expand=True)
        profit_scrollbar.pack(side="right", fill="y")
        
        # Populate profit tree
        for _, row in summary.iterrows():
            profit_tree.insert('', 'end', values=(
                row['Seller SKU'],
                row['Product Name'][:20] + "..." if len(row['Product Name']) > 20 else row['Product Name'],
                row['Variation'][:15] + "..." if len(row['Variation']) > 15 else row['Variation'],
                f"{row['TotalQty']:,}".replace(',', '.'),
                f"{row['Revenue']:,.0f}".replace(',', '.'),
                f"{row['Total Cost']:,.0f}".replace(',', '.'),
                f"{row['Profit']:,.0f}".replace(',', '.'),
                f"{row['Profit Margin %']:.1f}%"
            ))

    def export_to_excel(self):
        if self.merged_data is None:
            messagebox.showwarning("Warning", "Please process data first.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save report as"
        )
        if not file_path:
            return

        try:
            # Prepare data for analysis
            unique_orders = self.merged_data.drop_duplicates(subset=['Order ID'])
            total_orders = unique_orders['Order ID'].nunique()
            total_revenue = unique_orders['Total settlement amount'].sum()
            total_qty = self.merged_data['Quantity'].sum()

            # Create comprehensive summaries with cost analysis
            summary_by_variation = (
                self.merged_data.groupby(['Seller SKU', 'Product Name', 'Variation'], as_index=False)
                .agg({
                    'Quantity': 'sum',
                    'Order ID': 'nunique',
                    'Total settlement amount': 'sum'
                })
                .rename(columns={
                    'Quantity': 'Total Quantity',
                    'Order ID': 'Total Orders',
                    'Total settlement amount': 'Total Revenue'
                })
            )
            
            # Add cost calculations
            summary_by_variation['Cost per Unit'] = summary_by_variation['Product Name'].apply(self.get_product_cost)
            summary_by_variation['Total Cost'] = summary_by_variation['Total Quantity'] * summary_by_variation['Cost per Unit']
            summary_by_variation['Profit'] = summary_by_variation['Total Revenue'] - summary_by_variation['Total Cost']
            summary_by_variation['Profit Margin %'] = (summary_by_variation['Profit'] / summary_by_variation['Total Revenue'] * 100).round(2)
            summary_by_variation['Share 60%'] = summary_by_variation['Profit'] * 0.6
            summary_by_variation['Share 40%'] = summary_by_variation['Profit'] * 0.4
            
            # Summary by SKU
            summary_by_sku = (
                self.merged_data.groupby('Seller SKU', as_index=False)
                .agg({
                    'Quantity': 'sum',
                    'Order ID': 'nunique',
                    'Total settlement amount': 'sum'
                })
                .rename(columns={
                    'Quantity': 'Total Quantity',
                    'Order ID': 'Total Orders',
                    'Total settlement amount': 'Total Revenue'
                })
            )
            
            # Get first product name for each SKU for cost calculation
            sku_products = self.merged_data.groupby('Seller SKU')['Product Name'].first().to_dict()
            summary_by_sku['Cost per Unit'] = summary_by_sku['Seller SKU'].map(
                lambda sku: self.get_product_cost(sku_products.get(sku, ''))
            )
            summary_by_sku['Total Cost'] = summary_by_sku['Total Quantity'] * summary_by_sku['Cost per Unit']
            summary_by_sku['Profit'] = summary_by_sku['Total Revenue'] - summary_by_sku['Total Cost']
            summary_by_sku['Profit Margin %'] = (summary_by_sku['Profit'] / summary_by_sku['Total Revenue'] * 100).round(2)
            summary_by_sku['Share 60%'] = summary_by_sku['Profit'] * 0.6
            summary_by_sku['Share 40%'] = summary_by_sku['Profit'] * 0.4

            # Calculate total cost and profit
            total_cost = summary_by_sku['Total Cost'].sum()
            total_profit = total_revenue - total_cost
            total_share_60 = total_profit * 0.6
            total_share_40 = total_profit * 0.4

            # Daily sales analysis
            date_column = None
            possible_date_columns = [
                'Order created time(UTC)', 'Order creation time', 'Order Creation Time', 
                'Creation Time', 'Date', 'Order Date', 'Order created time', 'Created time'
            ]
            
            for col in possible_date_columns:
                if col in self.merged_data.columns:
                    date_column = col
                    break
            
            if date_column:
                try:
                    self.merged_data['Order Date'] = pd.to_datetime(self.merged_data[date_column]).dt.date
                    daily_sales = (
                        self.merged_data.groupby('Order Date', as_index=False)
                        .agg({
                            'Quantity': 'sum',
                            'Order ID': 'nunique',
                            'Total settlement amount': 'sum'
                        })
                        .rename(columns={
                            'Quantity': 'Daily Quantity',
                            'Order ID': 'Daily Orders',
                            'Total settlement amount': 'Daily Revenue'
                        })
                    )
                except:
                    daily_sales = pd.DataFrame({
                        'Order Date': ['Data tidak tersedia'],
                        'Daily Quantity': [0],
                        'Daily Orders': [0],
                        'Daily Revenue': [0]
                    })
            else:
                daily_sales = pd.DataFrame({
                    'Order Date': ['Kolom tanggal tidak ditemukan'],
                    'Daily Quantity': [0],
                    'Daily Orders': [0],
                    'Daily Revenue': [0]
                })

            # Top performing products by profit
            top_products = summary_by_variation.nlargest(10, 'Profit')

            # Calculate metrics
            avg_order_value = total_revenue / total_orders if total_orders > 0 else 0
            avg_profit_per_order = total_profit / total_orders if total_orders > 0 else 0
            overall_profit_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0

            # Date range
            if date_column and date_column in self.merged_data.columns:
                try:
                    date_range_start = pd.to_datetime(self.merged_data[date_column]).min()
                    date_range_end = pd.to_datetime(self.merged_data[date_column]).max()
                except:
                    date_range_start = datetime.now()
                    date_range_end = datetime.now()
            else:
                date_range_start = datetime.now()
                date_range_end = datetime.now()

            # Create Excel writer
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            workbook = writer.book

            # Define formats
            title_format = workbook.add_format({
                'bold': True, 'font_size': 16, 'align': 'center',
                'bg_color': '#4472C4', 'font_color': 'white'
            })
            
            header_format = workbook.add_format({
                'bold': True, 'font_size': 12,
                'bg_color': '#D9E2F3', 'border': 1
            })
            
            currency_format = workbook.add_format({
                'num_format': '#,##0', 'border': 1
            })
            
            number_format = workbook.add_format({
                'num_format': '#,##0', 'border': 1
            })
            
            percent_format = workbook.add_format({
                'num_format': '0.00%', 'border': 1
            })

            # ===== OVERVIEW SHEET =====
            overview_sheet = workbook.add_worksheet('Overview')
            overview_sheet.set_column('A:B', 25)
            overview_sheet.set_column('C:C', 20)
            
            row = 0
            overview_sheet.merge_range(f'A{row+1}:C{row+1}', 'LAPORAN PENJUALAN & PROFIT ANALYSIS', title_format)
            row += 2
            
            overview_sheet.write(row, 0, f'Periode:', header_format)
            overview_sheet.write(row, 1, f'{date_range_start.strftime("%d/%m/%Y")} - {date_range_end.strftime("%d/%m/%Y")}')
            row += 1
            
            overview_sheet.write(row, 0, f'Dibuat:', header_format)
            overview_sheet.write(row, 1, f'{datetime.now().strftime("%d %B %Y %H:%M")}')
            row += 3
            
            # Key metrics
            overview_sheet.write(row, 0, 'RINGKASAN PENJUALAN & PROFIT', header_format)
            row += 1
            overview_sheet.write(row, 0, 'Total Pesanan:')
            overview_sheet.write(row, 1, total_orders, number_format)
            row += 1
            overview_sheet.write(row, 0, 'Total Kuantitas:')
            overview_sheet.write(row, 1, total_qty, number_format)
            row += 1
            overview_sheet.write(row, 0, 'Total Revenue:')
            overview_sheet.write(row, 1, total_revenue, currency_format)
            row += 1
            overview_sheet.write(row, 0, 'Total Cost:')
            overview_sheet.write(row, 1, total_cost, currency_format)
            row += 1
            overview_sheet.write(row, 0, 'Total Profit:')
            overview_sheet.write(row, 1, total_profit, currency_format)
            row += 1
            overview_sheet.write(row, 0, 'Share 60%:')
            overview_sheet.write(row, 1, total_share_60, currency_format)
            row += 1
            overview_sheet.write(row, 0, 'Share 40%:')
            overview_sheet.write(row, 1, total_share_40, currency_format)
            row += 2
            overview_sheet.write(row, 0, 'Avg Order Value:')
            overview_sheet.write(row, 1, avg_order_value, currency_format)
            row += 1
            overview_sheet.write(row, 0, 'Avg Profit per Order:')
            overview_sheet.write(row, 1, avg_profit_per_order, currency_format)
            row += 1
            overview_sheet.write(row, 0, 'Overall Profit Margin:')
            overview_sheet.write(row, 1, overall_profit_margin / 100, percent_format)

            # ===== Write Other Sheets =====
            summary_by_variation.to_excel(writer, index=False, sheet_name='Summary by Product')
            writer.sheets['Summary by Product'].set_column(0, len(summary_by_variation.columns)-1, 20)

            summary_by_sku.to_excel(writer, index=False, sheet_name='Summary by SKU')
            writer.sheets['Summary by SKU'].set_column(0, len(summary_by_sku.columns)-1, 20)

            daily_sales.to_excel(writer, index=False, sheet_name='Daily Sales')
            writer.sheets['Daily Sales'].set_column(0, len(daily_sales.columns)-1, 20)

            top_products.to_excel(writer, index=False, sheet_name='Top Products')
            writer.sheets['Top Products'].set_column(0, len(top_products.columns)-1, 20)

            # === PRODUCT COST LIST ===
            if self.cost_data:
                cost_df = pd.DataFrame(list(self.cost_data.items()), columns=["Product Name", "Cost per Unit"])
                cost_df["Cost per Unit"] = cost_df["Cost per Unit"].apply(lambda x: f"{x:,.0f}".replace(",", "."))
                cost_df = cost_df.sort_values(by="Product Name")
                cost_df.to_excel(writer, index=False, sheet_name='Product Cost List')
                writer.sheets['Product Cost List'].set_column(0, 1, 30)



            writer.close()
            messagebox.showinfo("Export Complete", f"Data berhasil diekspor ke\n{file_path}")

        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

if __name__ == '__main__':
    root = tk.Tk()
    app = IncomeApp(root)
    root.mainloop()
            
