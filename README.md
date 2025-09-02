SME Finance Analyzer

An interactive dashboard for small businesses to analyze transactions, track performance, and export results.  
Built with Python and Streamlit, the tool generates monthly P&L, revenue/expense trends, and category breakdowns.

---

Features

• Upload CSV or Excel transaction files  
• Clean and normalize data automatically  
• Calculate revenue, expenses, profit, and margin  
• Generate monthly P&L with visualizations  
• Breakdown of expenses by category  
• Export results as clean CSV or PowerPoint report  

---

Tech Stack

Python (pandas, numpy)  
Streamlit (dashboard UI)  
Plotly (interactive charts)  
python-pptx (PowerPoint export)  

---

Sample Data

A sample dataset is included under `data/sample_transactions.csv`.  
Columns: date, type (sale/expense), category, description, amount, tax  

---
SME Finance Analyzer

Interactive dashboard to upload transaction data (CSV/Excel), auto-build monthly P&L, visualize trends, and export a consulting-style PowerPoint.

Data inputs: date, type (sale/expense), category, description, amount (negative for expenses), tax (optional).
Outputs: revenue, expenses, profit, margin, monthly trends, expense breakdown, exportable PPT.
---

How to Run

Clone the repository:
git clone https://github.com/<your-username>/sme-finance-analyzer.git
cd sme-finance-analyzer

Create and activate a virtual environment (optional but recommended):
python -m venv .venv
.venv\Scripts\activate   # Windows
source .venv/bin/activate   # Mac/Linux

Install dependencies:
pip install -r requirements.txt

Launch the dashboard:
streamlit run app.py
