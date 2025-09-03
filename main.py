import pandas as pd          # for data handling
import matplotlib.pyplot as plt  # for charts
import os                     # for file/folder operations
from fpdf import FPDF          # for PDF generation
from datetime import datetime  # for timestamp
import platform                # to detect OS
import subprocess              # to open PDF
import time                    # for small delays


# --- 1. Loading data ---
data_file = os.path.join("data", "marketing_data.csv")
df = pd.read_csv(data_file)
print("First lines of data:")
print(df.head())

# --- 2. Calculating the main metrics ---
df["CTR"] = df["clicks"] / df["impressions"] * 100
df["CPC"] = df["cost"] / df["clicks"]
df["ConversionRate"] = df["conversions"] / df["clicks"] * 100

print("\nCampaign metrics:")
print(df[["date", "campaign", "CTR", "CPC", "ConversionRate"]])

# --- 3. Chart: Campaign Costs: ---
# Group data by campaigns and sum up expenses
total_costs_per_campaign = df.groupby("campaign")["cost"].sum()

# Create a bar chart
plt.bar(total_costs_per_campaign.index, total_costs_per_campaign.values, color='orange')
plt.title("Expenses per Campaign")
plt.ylabel("Cost")
plt.show()


# --- 4. Graph: Clicks over Time ---
klicks_zeit = df.groupby("date")["clicks"].sum()
plt.plot(klicks_zeit.index, klicks_zeit.values, marker='o', color='green')
plt.title("Graph: Clicks over Time")
plt.xlabel("Date")
plt.ylabel("Clicks")
plt.xticks(rotation=45)
plt.show()

# --- 5.  Graph: Traffic share by campaigns---
traffic_quellen = df.groupby("campaign")["clicks"].sum()
plt.pie(traffic_quellen, labels=traffic_quellen.index, autopct="%1.1f%%", startangle=90)
plt.title("Graph: Traffic share by campaigns")
plt.show()

# --- 6. Save results in separate folder ---

# Create folder "charts" if it does not exist
charts_dir = "charts"
os.makedirs(charts_dir, exist_ok=True)

# 6.1 Save metrics to Excel
output_excel = os.path.join(charts_dir, "metrics.xlsx")
df.to_excel(output_excel, index=False)
print(f"Excel file saved: {output_excel}")

# 6.2 Save charts as PNG images
# 6.2.1 Bar chart: Expenses per Campaign
plt.bar(total_costs_per_campaign.index, total_costs_per_campaign.values, color='orange')
plt.title("Expenses per Campaign")
plt.ylabel("Cost")
plt.savefig(os.path.join(charts_dir, "expenses_per_campaign.png"))
plt.close()

# 6.2.2 Line chart: Clicks over Time
plt.plot(klicks_zeit.index, klicks_zeit.values, marker='o', color='green')
plt.title("Graph: Clicks over Time")
plt.xlabel("Date")
plt.ylabel("Clicks")
plt.xticks(rotation=45)
plt.savefig(os.path.join(charts_dir, "clicks_over_time.png"))
plt.close()

# 6.2.3 Pie chart: Traffic share by campaigns
plt.pie(traffic_quellen, labels=traffic_quellen.index, autopct="%1.1f%%", startangle=90)
plt.title("Graph: Traffic share by campaigns")
plt.savefig(os.path.join(charts_dir, "traffic_share.png"))
plt.close()

print("All results saved in the 'charts' folder.")


charts_dir = "charts"  # folder with PNG graphics


# --- Create PDF ---
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.alias_nb_pages()

# --- 1. Title page ---
pdf.add_page()
pdf.set_font("Arial", "B", 16)
pdf.cell(0, 10, "Marketing Report", ln=True, align="C")
pdf.ln(5)
pdf.set_font("Arial", "", 12)
pdf.cell(0, 10, "Author: Mariia Stepura", ln=True)
pdf.cell(0, 10, f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}", ln=True)

# --- 2. Table with metrics ---
pdf.ln(10)
pdf.set_font("Arial", "B", 12)
pdf.cell(0, 10, "Campaign Metrics", ln=True)
pdf.set_font("Arial", "", 10)

columns = ["date", "campaign", "CTR", "CPC", "ConversionRate"]
row_height = 8
col_widths = [pdf.w * 0.15, pdf.w * 0.25, pdf.w * 0.15, pdf.w * 0.15, pdf.w * 0.25]

# Headlines
for i, col in enumerate(columns):
    pdf.cell(col_widths[i], row_height, col, border=1, align="C")
pdf.ln(row_height)

# Data
for _, row in df.iterrows():
    pdf.cell(col_widths[0], row_height, str(row["date"]), border=1)
    pdf.cell(col_widths[1], row_height, str(row["campaign"]), border=1)
    pdf.cell(col_widths[2], row_height, f"{row['CTR']:.2f}", border=1, align="C")
    pdf.cell(col_widths[3], row_height, f"{row['CPC']:.2f}", border=1, align="C")
    pdf.cell(col_widths[4], row_height, f"{row['ConversionRate']:.2f}", border=1, align="C")
    pdf.ln(row_height)

# --- 3. Inserting charts with captions ---
charts = [
    ("expenses_per_campaign.png", "Figure 1: Expenses per Campaign", 25),
    ("clicks_over_time.png", "Figure 2: Clicks over Time", 40),  # сдвиг выше для цифр X
    ("traffic_share.png", "Figure 3: Traffic Share", 25)
]

for chart_file, caption, y_start in charts:
    pdf.add_page()
    img_path = os.path.join(charts_dir, chart_file)

    # Insert chart with page width and automatic height
    pdf.image(img_path, x=15, y=y_start, w=pdf.w-30)

    # Indent for signature under graph
    pdf.set_y(pdf.get_y() + 5)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, caption, ln=True, align="C")

# --- 4. Saving PDF ---
pdf_output = os.path.join(charts_dir, "Marketing_Report.pdf")
pdf.output(pdf_output)
print(f"PDF report saved: {pdf_output}")


# --- Auto open PDF ---
time.sleep(0.5)  # a short pause to allow the file to save
try:
    system = platform.system()
    if system == "Windows":
        os.startfile(pdf_output)
    elif system == "Darwin":  # macOS
        subprocess.run(["open", pdf_output])
    else:  # Linux
        subprocess.run(["xdg-open", pdf_output])
except Exception as e:
    print(f"Failed to open PDF automatically: {e}")