import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
import io
import matplotlib.pyplot as plt
import numpy as np

# Set page config
st.set_page_config(
    page_title="Income & Orders Summary",
    page_icon="💰",
    layout="wide"
)

class IncomeApp:
    def __init__(self):
        self.cost_file = "product_costs.json"
        self.load_cost_data()
    
    def load_cost_data(self):
        """Load cost data from JSON file"""
        try:
            if os.path.exists(self.cost_file):
                with open(self.cost_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except:
            return {}
        return {}
    
    def save_cost_data(self, cost_data):
        """Save cost data to JSON file"""
        with open(self.cost_file, 'w', encoding='utf-8') as f:
            json.dump(cost_data, f, ensure_ascii=False, indent=2)
    
    def get_product_cost(self, product_name, cost_data):
        """Get product cost from cost data"""
        return float(cost_data.get(product_name, 0.0))
    
    def process_data(self, pesanan_data, income_data, cost_data):
        """Process and merge data"""
        # Filter completed orders
        df1 = pesanan_data[pesanan_data['Order Status'] == 'Selesai']
        
        # Remove duplicates from income data
        df2 = income_data.drop_duplicates(subset=['Order/adjustment ID'])
        
        # Merge data
        merged = pd.merge(df1, df2, left_on='Order ID', right_on='Order/adjustment ID', how='inner')
        
        if merged.empty:
            return None, None
        
        # Create summary
        summary = merged.groupby(['Seller SKU', 'Product Name', 'Variation'], as_index=False).agg(
            TotalQty=('Quantity', 'sum'),
            Revenue=('Total settlement amount', 'sum')
        )
        
        # Add cost calculations
        summary['Cost per Unit'] = summary['Product Name'].apply(
            lambda x: self.get_product_cost(x, cost_data)
        )
        summary['Total Cost'] = summary['TotalQty'] * summary['Cost per Unit']
        summary['Profit'] = summary['Revenue'] - summary['Total Cost']
        summary['Profit Margin %'] = (summary['Profit'] / summary['Revenue'] * 100).round(2)
        summary['Share 60%'] = summary['Profit'] * 0.6
        summary['Share 40%'] = summary['Profit'] * 0.4
        
        return merged, summary
    
    def create_excel_report(self, merged_data, summary_data, cost_data):
        """Create Excel report"""
        output = io.BytesIO()
        
        # Calculate totals
        unique_orders = merged_data.drop_duplicates(subset=['Order ID'])
        total_orders = unique_orders['Order ID'].nunique()
        total_revenue = unique_orders['Total settlement amount'].sum()
        total_qty = merged_data['Quantity'].sum()
        
        # Summary by SKU
        summary_by_sku = (
            merged_data.groupby('Seller SKU', as_index=False)
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
        sku_products = merged_data.groupby('Seller SKU')['Product Name'].first().to_dict()
        summary_by_sku['Cost per Unit'] = summary_by_sku['Seller SKU'].map(
            lambda sku: self.get_product_cost(sku_products.get(sku, ''), cost_data)
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
            if col in merged_data.columns:
                date_column = col
                break
        
        if date_column:
            try:
                merged_data_copy = merged_data.copy()
                merged_data_copy['Order Date'] = pd.to_datetime(merged_data_copy[date_column]).dt.date
                daily_sales = (
                    merged_data_copy.groupby('Order Date', as_index=False)
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
        top_products = summary_data.nlargest(10, 'Profit')
        
        # Create Excel writer
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
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
            
            # Overview sheet
            overview_sheet = workbook.add_worksheet('Overview')
            overview_sheet.set_column('A:B', 25)
            overview_sheet.set_column('C:C', 20)
            
            row = 0
            overview_sheet.merge_range(f'A{row+1}:C{row+1}', 'LAPORAN PENJUALAN & PROFIT ANALYSIS', title_format)
            row += 2
            
            # Date range
            if date_column and date_column in merged_data.columns:
                try:
                    date_range_start = pd.to_datetime(merged_data[date_column]).min()
                    date_range_end = pd.to_datetime(merged_data[date_column]).max()
                except:
                    date_range_start = datetime.now()
                    date_range_end = datetime.now()
            else:
                date_range_start = datetime.now()
                date_range_end = datetime.now()
            
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
            
            # Calculate additional metrics
            avg_order_value = total_revenue / total_orders if total_orders > 0 else 0
            avg_profit_per_order = total_profit / total_orders if total_orders > 0 else 0
            overall_profit_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
            
            overview_sheet.write(row, 0, 'Avg Order Value:')
            overview_sheet.write(row, 1, avg_order_value, currency_format)
            row += 1
            overview_sheet.write(row, 0, 'Avg Profit per Order:')
            overview_sheet.write(row, 1, avg_profit_per_order, currency_format)
            row += 1
            overview_sheet.write(row, 0, 'Overall Profit Margin:')
            overview_sheet.write(row, 1, overall_profit_margin / 100, percent_format)
            
            # Write other sheets
            summary_data.to_excel(writer, index=False, sheet_name='Summary by Product')
            summary_by_sku.to_excel(writer, index=False, sheet_name='Summary by SKU')
            daily_sales.to_excel(writer, index=False, sheet_name='Daily Sales')
            top_products.to_excel(writer, index=False, sheet_name='Top Products')
            
            # Product cost list
            if cost_data:
                cost_df = pd.DataFrame(list(cost_data.items()), columns=["Product Name", "Cost per Unit"])
                cost_df = cost_df.sort_values(by="Product Name")
                cost_df.to_excel(writer, index=False, sheet_name='Product Cost List')
        
        output.seek(0)
        return output


def main():
    st.title("💰 Income & Orders Summary with Cost Management")
    st.markdown("---")
    
    # Initialize app
    app = IncomeApp()
    
    # Initialize session state
    if 'cost_data' not in st.session_state:
        st.session_state.cost_data = app.load_cost_data()
    if 'pesanan_data' not in st.session_state:
        st.session_state.pesanan_data = None
    if 'income_data' not in st.session_state:
        st.session_state.income_data = None
    if 'merged_data' not in st.session_state:
        st.session_state.merged_data = None
    if 'summary_data' not in st.session_state:
        st.session_state.summary_data = None
    
    # Create tabs
    tab1, tab2, tab3 = st.tabs(["📊 Main Analysis", "💸 Cost Management", "📈 Analytics"])
    
    with tab1:
        st.header("Data Loading & Analysis")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Upload Pesanan Selesai")
            pesanan_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'], key="pesanan")
            if pesanan_file:
                try:
                    df = pd.read_excel(pesanan_file, header=0, skiprows=[1])
                    df.columns = df.columns.str.strip()
                    st.session_state.pesanan_data = df
                    st.success(f"✅ Pesanan data loaded: {len(df)} rows")
                    st.dataframe(df.head())
                except Exception as e:
                    st.error(f"Error loading pesanan file: {str(e)}")
        
        with col2:
            st.subheader("Upload Income Data")
            income_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'], key="income")
            if income_file:
                try:
                    df = pd.read_excel(income_file)
                    df.columns = df.columns.str.strip()
                    st.session_state.income_data = df
                    st.success(f"✅ Income data loaded: {len(df)} rows")
                    st.dataframe(df.head())
                except Exception as e:
                    st.error(f"Error loading income file: {str(e)}")
        
        st.markdown("---")
        
        # Process data button
        if st.button("🔄 Process & Summarize Data", type="primary"):
            if st.session_state.pesanan_data is not None and st.session_state.income_data is not None:
                merged, summary = app.process_data(
                    st.session_state.pesanan_data, 
                    st.session_state.income_data, 
                    st.session_state.cost_data
                )
                
                if merged is not None:
                    st.session_state.merged_data = merged
                    st.session_state.summary_data = summary
                    st.success("✅ Data processed successfully!")
                else:
                    st.error("❌ No matching data found between files")
            else:
                st.warning("⚠️ Please upload both files first")
        
        # Show summary if available
        if st.session_state.summary_data is not None:
            st.subheader("📋 Summary Results")
            
            # Format numbers for display
            summary_display = st.session_state.summary_data.copy()
            summary_display['TotalQty'] = summary_display['TotalQty'].apply(lambda x: f"{x:,}")
            summary_display['Revenue'] = summary_display['Revenue'].apply(lambda x: f"Rp {x:,.0f}")
            summary_display['Total Cost'] = summary_display['Total Cost'].apply(lambda x: f"Rp {x:,.0f}")
            summary_display['Profit'] = summary_display['Profit'].apply(lambda x: f"Rp {x:,.0f}")
            summary_display['Share 60%'] = summary_display['Share 60%'].apply(lambda x: f"Rp {x:,.0f}")
            summary_display['Share 40%'] = summary_display['Share 40%'].apply(lambda x: f"Rp {x:,.0f}")
            
            st.dataframe(summary_display, use_container_width=True)
            
            # Show overall totals
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📊 Show Overall Totals"):
                    # Calculate totals
                    unique_orders = st.session_state.merged_data.drop_duplicates(subset=['Order ID'])
                    total_orders = unique_orders['Order ID'].nunique()
                    total_revenue = unique_orders['Total settlement amount'].sum()
                    total_cost = st.session_state.summary_data['Total Cost'].sum()
                    total_profit = total_revenue - total_cost
                    total_share_60 = total_profit * 0.6
                    total_share_40 = total_profit * 0.4
                    
                    # Display metrics
                    st.subheader("🎯 Overall Performance")
                    
                    metric_col1, metric_col2, metric_col3 = st.columns(3)
                    with metric_col1:
                        st.metric("Total Orders", f"{total_orders:,}")
                        st.metric("Total Revenue", f"Rp {total_revenue:,.0f}")
                    with metric_col2:
                        st.metric("Total Cost", f"Rp {total_cost:,.0f}")
                        st.metric("Total Profit", f"Rp {total_profit:,.0f}")
                    with metric_col3:
                        st.metric("Share 60%", f"Rp {total_share_60:,.0f}")
                        st.metric("Share 40%", f"Rp {total_share_40:,.0f}")
            
            with col2:
                if st.button("📈 Show Profit Analysis"):
                    st.subheader("💰 Profit Analysis")
                    
                    # Top products by profit
                    top_products = st.session_state.summary_data.nlargest(10, 'Profit')
                    
                    st.write("**Top 10 Products by Profit:**")
                    profit_display = top_products[['Product Name', 'Profit', 'Profit Margin %']].copy()
                    profit_display['Profit'] = profit_display['Profit'].apply(lambda x: f"Rp {x:,.0f}")
                    profit_display['Profit Margin %'] = profit_display['Profit Margin %'].apply(lambda x: f"{x:.1f}%")
                    st.dataframe(profit_display, use_container_width=True)
            
            # Export button
            if st.button("📥 Export to Excel"):
                try:
                    excel_data = app.create_excel_report(
                        st.session_state.merged_data,
                        st.session_state.summary_data,
                        st.session_state.cost_data
                    )
                    
                    st.download_button(
                        label="💾 Download Excel Report",
                        data=excel_data,
                        file_name=f"income_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("✅ Excel report generated successfully!")
                except Exception as e:
                    st.error(f"Error generating Excel report: {str(e)}")
    
    with tab2:
        st.header("💸 Cost Management")
        st.markdown("Kelola harga modal produk untuk menghitung profit secara akurat")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.subheader("Add/Edit Product Cost")
            
            # Product selection
            if st.session_state.pesanan_data is not None:
                products = sorted(st.session_state.pesanan_data['Product Name'].astype(str).unique())
                selected_product = st.selectbox("Select Product", options=products, key="product_select")
            else:
                selected_product = st.text_input("Product Name", key="product_input")
            
            cost_input = st.number_input("Cost per Unit", min_value=0.0, format="%.2f", key="cost_input")
            
            col_btn1, col_btn2, col_btn3 = st.columns(3)
            with col_btn1:
                if st.button("💾 Save Cost"):
                    if selected_product and cost_input >= 0:
                        st.session_state.cost_data[selected_product] = cost_input
                        app.save_cost_data(st.session_state.cost_data)
                        st.success(f"✅ Cost saved for {selected_product}")
                        st.rerun()
                    else:
                        st.warning("⚠️ Please enter valid product and cost")
            
            with col_btn2:
                if st.button("🗑️ Delete Cost"):
                    if selected_product in st.session_state.cost_data:
                        del st.session_state.cost_data[selected_product]
                        app.save_cost_data(st.session_state.cost_data)
                        st.success(f"✅ Cost deleted for {selected_product}")
                        st.rerun()
                    else:
                        st.warning("⚠️ Product not found in cost data")
            
            with col_btn3:
                if st.button("🔄 Clear Form"):
                    st.rerun()

        st.download_button(
            label="⬇️ Backup (JSON)",
            data=json.dumps(st.session_state.cost_data, ensure_ascii=False, indent=2),
            file_name="product_costs_backup.json",
            mime="application/json"
        )
        
        with col2:
            st.subheader("Quick Stats")
            if st.session_state.cost_data:
                st.metric("Total Products", len(st.session_state.cost_data))
                avg_cost = sum(st.session_state.cost_data.values()) / len(st.session_state.cost_data)
                st.metric("Average Cost", f"Rp {avg_cost:,.0f}")
            else:
                st.metric("Total Products", 0)
        
        st.markdown("---")
        
        # Cost data table
        if st.session_state.cost_data:
            st.subheader("📋 Current Cost Data")
            cost_df = pd.DataFrame(
                list(st.session_state.cost_data.items()),
                columns=["Product Name", "Cost per Unit"]
            )
            cost_df = cost_df.sort_values("Product Name")
            cost_df['Cost per Unit'] = cost_df['Cost per Unit'].apply(lambda x: f"Rp {x:,.0f}")
            st.dataframe(cost_df, use_container_width=True)
        else:
            st.info("ℹ️ No cost data available. Add some product costs to get started.")
    
    with tab3:
        st.header("📈 Analytics Dashboard")
        
        if st.session_state.summary_data is not None:
            # Revenue vs Profit chart
            st.subheader("📊 Revenue vs Profit Analysis")
            
            chart_data = st.session_state.summary_data[['Product Name', 'Revenue', 'Profit']].copy()
            chart_data = chart_data.sort_values('Revenue', ascending=False).head(10)
            
            st.bar_chart(chart_data.set_index('Product Name')[['Revenue', 'Profit']])
            
            # Profit margin distribution using matplotlib
            st.subheader("📈 Profit Margin Distribution")
            
            # Create histogram using matplotlib
            fig, ax = plt.subplots(figsize=(10, 6))
            profit_margins = st.session_state.summary_data['Profit Margin %'].dropna()
            
            if not profit_margins.empty:
                ax.hist(profit_margins, bins=20, alpha=0.7, color='skyblue', edgecolor='black')
                ax.set_xlabel('Profit Margin (%)')
                ax.set_ylabel('Frequency')
                ax.set_title('Distribution of Profit Margins')
                ax.grid(True, alpha=0.3)
                
                # Add statistics
                mean_margin = profit_margins.mean()
                median_margin = profit_margins.median()
                ax.axvline(mean_margin, color='red', linestyle='--', label=f'Mean: {mean_margin:.1f}%')
                ax.axvline(median_margin, color='green', linestyle='--', label=f'Median: {median_margin:.1f}%')
                ax.legend()
                
                st.pyplot(fig)
                
                # Additional statistics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Mean Margin", f"{mean_margin:.1f}%")
                with col2:
                    st.metric("Median Margin", f"{median_margin:.1f}%")
                with col3:
                    st.metric("Min Margin", f"{profit_margins.min():.1f}%")
                with col4:
                    st.metric("Max Margin", f"{profit_margins.max():.1f}%")
            else:
                st.warning("⚠️ No profit margin data available")
            
            # Top performing products
            st.subheader("🏆 Top Performing Products")
            
            top_by_profit = st.session_state.summary_data.nlargest(5, 'Profit')
            top_by_margin = st.session_state.summary_data.nlargest(5, 'Profit Margin %')
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**By Total Profit:**")
                profit_display = top_by_profit[['Product Name', 'Profit']].copy()
                profit_display['Profit'] = profit_display['Profit'].apply(lambda x: f"Rp {x:,.0f}")
                st.dataframe(profit_display, use_container_width=True)
            
            with col2:
                st.write("**By Profit Margin:**")
                margin_display = top_by_margin[['Product Name', 'Profit Margin %']].copy()
                margin_display['Profit Margin %'] = margin_display['Profit Margin %'].apply(lambda x: f"{x:.1f}%")
                st.dataframe(margin_display, use_container_width=True)
        
        else:
            st.info("ℹ️ Please process data first to see analytics")

if __name__ == "__main__":
    main()
