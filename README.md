# 📊 Income & Order Analytics Dashboard

A **Streamlit-based** business-intelligence tool that merges **order data** and **settlement data** to deliver real-time profit analysis, cost tracking, and strategic recommendations for e-commerce sellers.

---

## 🚀 Features

| Feature | Description |
|---------|-------------|
| **📁 Drag-and-Drop Upload** | Accepts two Excel files: “Completed Orders” & “Settlement/Income”. |
| **💸 Cost Management** | Maintain a JSON-backed cost database per SKU. |
| **📊 Live Dashboard** | Key KPIs, profit margins, order counts, and revenue splits (60 % / 40 %). |
| **📈 Advanced Analytics** | Scatter plots, Pareto charts, quadrant analysis (Stars / Workhorses / Niche / Problem). |
| **🤖 AI Summary** | One-click prompt generator for ChatGPT with curated strategic questions. |
| **📥 Excel Export** | Full multi-sheet workbook (`Ringkasan`, `Penjualan Harian`, `Produk Teratas`, etc.). |

---

## 🏁 Quick Start

1. **Clone / download** the repository.
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt

   📋 File Requirements

   | File               | Required Columns (sample names)                                                   |
| ------------------ | --------------------------------------------------------------------------------- |
| **Orders** (Excel) | `Order ID`, `Order Status`, `Seller SKU`, `Product Name`, `Variation`, `Quantity` |
| **Income** (Excel) | `Order/adjustment ID`, `Total settlement amount`                                  |

✅ Rows must be UTF-8 clean; extra columns are ignored.


💰 Cost JSON Format
The app automatically stores product costs in product_costs.json:

{
  "White T-Shirt XL": 25000,
  "Black Hoodie M": 45000
}

You can import/export this file from the Cost Management tab.

🧪 Example Workflow
Upload orders_2024_07.xlsx → ✅ 3 412 rows loaded.
Upload settlement_2024_07.xlsx → ✅ 3 399 rows loaded.
Click 🔄 Process Data → Dashboard populated with KPIs.
Add missing product costs in 💸 Manajemen Biaya.
Hit 📥 Ekspor Laporan to download income_report_20240717_145522.xlsx.
Use 💬 Ringkas & Lanjut ke ChatGPT for strategic insights.
🛠️ Development Tips
All styling is in-line via st.markdown(..., unsafe_allow_html=True)—edit the <style> block in income.py to customize themes quickly.
The app is stateless except for st.session_state, so it scales well on Streamlit Cloud or Docker.
To extend AI features, add your OpenAI key and uncomment calls to OpenAI() – everything else is ready.

📄 License
MIT © 2024 – feel free to fork, improve, and commercialize.




