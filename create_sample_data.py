"""
Script to create a sample Excel file for testing transpose functionality
"""

from openpyxl import Workbook

def create_sample_excel():
    """Create a sample Excel file with test data"""
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Original Data"
    
    # Add sample data in rows
    # This will be transposed to columns
    sample_data = [
        ["Product", "Q1", "Q2", "Q3", "Q4"],
        ["Sales", 10000, 15000, 12000, 18000],
        ["Expenses", 5000, 6000, 5500, 7000],
        ["Profit", 5000, 9000, 6500, 11000],
        ["Growth %", 15, 22, 18, 25]
    ]
    
    # Write data to cells
    for row_idx, row_data in enumerate(sample_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    # Adjust column widths for better visibility
    ws.column_dimensions['A'].width = 15
    for col_num in range(2, 6):
        ws.column_dimensions[chr(64 + col_num)].width = 12
    
    # Save the file
    output_file = "sample_data.xlsx"
    wb.save(output_file)
    
    print(f"✓ Sample Excel file created: {output_file}")
    print("\nSample data structure:")
    print("Original format (Rows):")
    for row in sample_data:
        print(f"  {row}")


if __name__ == "__main__":
    create_sample_excel()
