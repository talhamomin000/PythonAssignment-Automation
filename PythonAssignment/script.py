import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
from fpdf import FPDF


# Function to calculate KPIs and generate visualizations
def generate_report():
    # Load the dataset
    df = pd.read_excel('file:///C:/Users/Asus/Downloads/Sales%20Data.xls')  # Adjust the path as needed

    # Convert 'Date' column to datetime type
    df['Date'] = pd.to_datetime(df['Date'])

    # Calculate yearly KPIs
    df['Year'] = df['Date'].dt.year

    # Total Sales per Category
    total_sales_per_category = df.groupby('Category')['TotalSales'].sum()

    # Average Order Value (AOV) per Category
    df['AOV'] = df['TotalSales'] / df['QuantitySold']
    aov_per_category = df.groupby('Category')['AOV'].mean()

    # Return on Marketing Spend (ROMS) per Category
    roms_per_category = total_sales_per_category / (df.groupby('Category')['QuantitySold'].sum())

    # Combine all KPIs into a single DataFrame for better clarity
    kpi_df = pd.DataFrame({
        'Total Sales': total_sales_per_category,
        'AOV': aov_per_category,
        'ROMS': roms_per_category
    })

    # Save Total Sales bar chart
    total_sales_per_category.plot(kind='bar', color='skyblue')
    plt.title('Total Sales per Category')
    plt.xlabel('Category')
    plt.ylabel('Total Sales ($)')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig('total_sales_per_category.png')

    # Save AOV line graph
    aov_per_category.plot(kind='line', marker='o', color='green')
    plt.title('Average Order Value (AOV) per Category')
    plt.xlabel('Category')
    plt.ylabel('Average Order Value ($)')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig('aov_per_category.png')

    # Save ROMS bar chart
    roms_per_category.plot(kind='bar', color='orange')
    plt.title('Return on Marketing Spend (ROMS) per Category')
    plt.xlabel('Category')
    plt.ylabel('ROMS')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig('roms_per_category.png')

    # Initialize PDF
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Set title
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(200, 10, txt="Sales Data Analysis Report", ln=True, align='C')

    # Add KPIs table
    pdf.ln(10)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(40, 10, 'Category', 1)
    pdf.cell(60, 10, 'Total Sales ($)', 1)
    pdf.cell(60, 10, 'Average Order Value ($)', 1)
    pdf.cell(60, 10, 'ROMS', 1)
    pdf.ln()

    # Add KPI values
    pdf.set_font('Arial', '', 12)
    for category, row in kpi_df.iterrows():
        pdf.cell(40, 10, category, 1)
        pdf.cell(60, 10, f"${row['Total Sales']:,.2f}", 1)
        pdf.cell(60, 10, f"${row['AOV']:,.2f}", 1)
        pdf.cell(60, 10, f"{row['ROMS']:.2f}", 1)
        pdf.ln()

    # Add visualizations to the PDF
    pdf.ln(10)
    pdf.image('total_sales_per_category.png', x=10, w=180)

    pdf.ln(10)
    pdf.image('aov_per_category.png', x=10, w=180)

    pdf.ln(10)
    pdf.image('roms_per_category.png', x=10, w=180)

    # Output the PDF
    pdf.output('Sales_Data_Analysis_Report.pdf')
    print("Report generated successfully!")

# Run the report generation function
generate_report()


