import os
import pandas as pd
import requests
from pypdf import PdfReader
from datetime import datetime, timedelta

# === CONFIGURATION ===
excel_path = r'C:\Users\kshere\Downloads\DBRS IN\sample SFI.xlsx'
column_name = 'URL'
download_folder = r'C:\Users\kshere\Downloads\DBRS IN\Download'

# === CREATE DOWNLOAD FOLDER IF NOT EXISTS ===
if not os.path.exists(download_folder):
    os.makedirs(download_folder)

# === READ EXCEL FILE ===
df = pd.read_excel(excel_path)

# Add necessary columns if not present
for col in ['Create Date', 'Update Date', 'File Name', 'Choose File', 'Document Date', 'Effective Date']:
    if col not in df.columns:
        df[col] = ''

# === FUNCTION TO PARSE PDF DATE ===
def parse_pdf_date(pdf_date):
    if pdf_date and pdf_date.startswith('D:'):
        try:
            return datetime.strptime(pdf_date[2:16], '%Y%m%d%H%M%S').strftime('%Y-%m-%d %H:%M:%S')
        except ValueError:
            return 'Invalid Format'
    return 'Not Available'

# === DOWNLOAD, EXTRACT METADATA, RENAME, AND UPDATE EXCEL ===
for index, row in df.iterrows():
    url = row[column_name]
    try:
        response = requests.get(url)
        response.raise_for_status()

        # Temporary filename
        original_filename = os.path.basename(url)
        file_ext = os.path.splitext(original_filename)[1]
        temp_path = os.path.join(download_folder, original_filename)

        # Save the file temporarily
        with open(temp_path, 'wb') as f:
            f.write(response.content)

        # Extract metadata
        create_date, update_date, document_date = '', '', ''
        if file_ext.lower() == '.pdf':
            reader = PdfReader(temp_path)
            metadata = reader.metadata
            create_date = parse_pdf_date(metadata.get('/CreationDate'))
            update_date = parse_pdf_date(metadata.get('/ModDate'))

            # Extract YYYY-MM-DD only for Document Date
            if update_date not in ('', 'Invalid Format', 'Not Available', 'Error'):
                document_date = update_date.split(' ')[0]

        # Prepare renamed filename
        safe_update_date = update_date.replace(':', '-').replace(' ', '_') if update_date not in ('', 'Invalid Format', 'Not Available', 'Error') else 'UnknownDate'
        new_filename = f"{os.path.splitext(original_filename)[0]}__{safe_update_date}{file_ext}"
        final_path = os.path.join(download_folder, new_filename)

        # Rename the file
        os.rename(temp_path, final_path)

        # Update Excel columns
        df.at[index, 'Create Date'] = create_date
        df.at[index, 'Update Date'] = update_date
        df.at[index, 'Document Date'] = document_date
        df.at[index, 'File Name'] = new_filename
        df.at[index, 'Choose File'] = final_path

        # Calculate Effective Date
        if document_date not in ('', 'Error', 'Invalid Format', 'Not Available'):
            try:
                doc_date_obj = datetime.strptime(document_date, '%Y-%m-%d')
                effective_date = doc_date_obj + timedelta(days=1)
                df.at[index, 'Effective Date'] = effective_date.strftime('%Y-%m-%d')
            except Exception:
                df.at[index, 'Effective Date'] = 'Error'
        else:
            df.at[index, 'Effective Date'] = 'Error'

        print(f"✅ {new_filename} | Created: {create_date} | Updated: {update_date} | Saved to: {final_path}")

    except Exception as e:
        print(f"❌ Failed to process {url} - Error: {e}")
        df.at[index, 'Create Date'] = 'Error'
        df.at[index, 'Update Date'] = 'Error'
        df.at[index, 'Document Date'] = 'Error'
        df.at[index, 'File Name'] = 'Error'
        df.at[index, 'Choose File'] = 'Error'
        df.at[index, 'Effective Date'] = 'Error'


# === SAVE FINAL EXCEL FILE WITH "Effective Date" AS 3RD COLUMN ===
# Reorder columns to place 'Effective Date' at index 2 (3rd position)
cols = df.columns.tolist()
if 'Effective Date' in cols:
    cols.insert(2, cols.pop(cols.index('Effective Date')))
    df = df[cols]

output_path = r'C:\Users\kshere\Downloads\DBRS IN\updated_sample_SFI.xlsx'
df.to_excel(output_path, index=False)
print(f"\n📄 Excel updated and saved to: {output_path}")