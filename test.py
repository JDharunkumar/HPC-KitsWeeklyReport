import os
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill
import smtplib
from email.message import EmailMessage

# ================================
# 1. SQL Server Configuration (Imported)
# ================================
from config import DB_SERVER, DB_DATABASE

# ================================
# 2. Email Configuration (Imported)
# ================================
from config import SMTP_SERVER, SMTP_PORT, EMAIL_TO, EMAIL_CC

# ================================
# 3. SQL Queries (Imported)
# ================================
from sql_queries import MODELS_QUERY, OPTIONS_QUERY, SPECIALS_QUERY

# ================================
# 4. Database Helper
# ================================
def get_db_engine():
    conn_str = (
        f"mssql+pyodbc://{DB_SERVER}/{DB_DATABASE}"
        "?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
    )
    return create_engine(conn_str)

def fetch_data(product_line_id=None, product_id=None):
    engine = get_db_engine()
    
    params = {}
    if product_line_id:
        params['product_line_id'] = product_line_id
    if product_id:
        params['product_id'] = product_id
    
    def get_query(base_sql, include_product_id=True, alias_pl='pl.DBID', alias_p='p.DBID'):
        filters = []
        if product_line_id:
            filters.append(f"AND {alias_pl} = :product_line_id")
        if product_id and include_product_id:
            filters.append(f"AND {alias_p} = :product_id")
            
        if not filters:
            return base_sql
            
        # Insert filters before ORDER BY
        if "ORDER BY" in base_sql:
            parts = base_sql.rsplit("ORDER BY", 1)
            return f"{parts[0]} {' '.join(filters)} ORDER BY {parts[1]}"
        return f"{base_sql} {' '.join(filters)}"

    # Fetch Models data (Only supports ProductLine filter)
    models_sql = get_query(MODELS_QUERY, include_product_id=False)
    models_df = pd.read_sql(models_sql, engine, params=params)
    
    # Fetch Options and Specials data
    options_df = pd.read_sql(get_query(OPTIONS_QUERY), engine, params=params)
    specials_df = pd.read_sql(get_query(SPECIALS_QUERY), engine, params=params)
    
    return models_df, options_df, specials_df

# ================================
# 5. Excel Generation with Two Sheets
# ================================
def create_excel_report(product_line_id=None, product_id=None):
    models_df, options_df, specials_df = fetch_data(product_line_id, product_id)
    
    # Prepare Kits data (Options + Specials)
    options_df['Type'] = 'Option'
    specials_df['Type'] = 'Special'
    
    # Add missing Notes column to specials if not present
    if 'Notes' not in specials_df.columns:
        specials_df['Notes'] = ''
    
    # Combine options and specials
    kits_df = pd.concat([options_df, specials_df], ignore_index=True)
    
    # Format currency for both sheets
    def format_currency(x):
        try:
            return f"${float(x):,.2f}"
        except:
            return "" if pd.isna(x) else str(x)
    
    # Format Kits data
    kits_rows = []
    for _, row in kits_df.iterrows():
        kits_rows.append({
            'ProductLine Name': row['ProductLineName'],
            'Option Name': row['OptionName'],
            'Description': row['Description'],
            'Date Created': row['DateCreated'],
            'Date Modified': row['DateModified'],
            'Current List Price': format_currency(row['CurrentListPrice']),
            'Standard Cost': format_currency(row['stdCost'])
        })
    
    # Format Models data
    models_rows = []
    for _, row in models_df.iterrows():
        models_rows.append({
            'ProductLine Name': row['ProductLineName'],
            'Model Name': row['Name'],
            'Description': row['Description'],
            'Date Created': row['DateCreated'],
            'Date Modified': row['DateModified'],
            'Current List Price': format_currency(row['CurrentListPrice']),
            'Standard Cost': format_currency(row['stdCost']),
            'Site Name': row['SiteName']
        })
    
    # Create DataFrames
    kits_final_df = pd.DataFrame(kits_rows)
    models_final_df = pd.DataFrame(models_rows)
    
    # Sort data
    kits_final_df = kits_final_df.sort_values(['ProductLine Name', 'Option Name'], ascending=[True, True])
    models_final_df = models_final_df.sort_values(['ProductLine Name', 'Model Name'], ascending=[True, True])
    
    # Reset index
    kits_final_df = kits_final_df.reset_index(drop=True)
    models_final_df = models_final_df.reset_index(drop=True)
    
    # Create Excel file
    os.makedirs("reports", exist_ok=True)
    file_name = f"reports/UniverseKits And Model Report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    
    # Write to Excel with multiple sheets
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        kits_final_df.to_excel(writer, sheet_name='Kits', index=False)
        models_final_df.to_excel(writer, sheet_name='Models', index=False)
    
    # Format Excel sheets
    wb = load_workbook(file_name)
    
    # Format both sheets
    for sheet_name in ['Kits', 'Models']:
        ws = wb[sheet_name]
        
        # Define font styles
        header_font = Font(bold=True, size=11)
        data_font = Font(size=8)
        center_alignment = Alignment(vertical="center")
        
        # Header style (row 1)
        header_fill = PatternFill("solid", fgColor="FFD700")
        for col in ws.iter_cols(min_row=1, max_row=1):
            for cell in col:
                cell.font = header_font
                cell.alignment = center_alignment
                cell.fill = header_fill
        
        # Data style (rows 2 onwards)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.font = data_font
                cell.alignment = center_alignment
        
        # Auto-fit columns
        for col in ws.columns:
            w = max(len(str(c.value)) for c in col if c.value) + 2
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(w, 50)
        
        # Add auto filter
        ws.auto_filter.ref = ws.dimensions
        
        # Freeze the first row
        ws.freeze_panes = 'A2'
    
    wb.save(file_name)
    return file_name

# ================================
# 6. Email Sender
# ================================
def send_email(file_path):
    msg = EmailMessage()
    msg['Subject'] = f'Kit Usage Report ({datetime.now():%d-%b-%Y})'
    msg['From'] = 'reports@hussmann.com'
    msg['To'] = ','.join(EMAIL_TO)
    
    # Add CC recipients if they exist
    if 'EMAIL_CC' in globals() and EMAIL_CC:
        msg['Cc'] = ', '.join(EMAIL_CC)
    
    # Create HTML content with blue disclaimer
    html_content = """
    <html>
    <body>
        <p>Hello,</p>
        <p>Please find attached the UniverseKits And Model Report</p>
        <p>Regards,<br>
        Automated Reporting System</p>
        <p>This is an automated email, please do not reply.</p>
        <p><span style="color: blue;"><strong>Disclaimer:</strong><br>
        The content of this email is intended solely for the person or entity to whom it is addressed. It may contain confidential information. If you are not the intended recipient or have received this message in error, please contact HPC Support immediately.</span></p>
    </body>
    </html>
    """
    
    # Set HTML content
    msg.set_content(html_content, subtype='html')
    
    with open(file_path, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=os.path.basename(file_path)
        )
    
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
        s.send_message(msg)

# ================================
# 7. Main
# ================================
if __name__ == '__main__':
    try:
        path = create_excel_report()
        send_email(path)
        print("Report generated:", path)
    except Exception as e:
        print("Error:", e)
