import os
import io
import zipfile
import requests
from datetime import datetime
import pandas as pd
from flask import Flask, render_template, send_file, flash, redirect, url_for
from dotenv import load_dotenv

# Load environment variables from .env file (for local testing)
load_dotenv()

app = Flask(__name__)
# Secret key is needed for flashing messages in Flask
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'default-dev-key')

def fetch_documents():
    api_key = os.environ.get('BOLDSIGN_API_KEY')
    if not api_key:
        return []

    headers = {
        'X-API-KEY': api_key,
        'Accept': 'application/json'
    }
    
    all_docs = []
    page = 1
    page_size = 100
    total_pages = 1
    
    while page <= total_pages:
        url = f"https://api.boldsign.com/v1/document/list?page={page}&pageSize={page_size}"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            docs = data.get('result', [])
            all_docs.extend(docs)
            
            page_details = data.get('pageDetails', {})
            total_pages = page_details.get('totalPages', 1)
            page += 1
        else:
            print(f"Error fetching page {page}: HTTP {response.status_code} - {response.text}")
            break
            
    return all_docs

def extract_inn_code(title):
    parts = str(title).split('-')
    if len(parts) >= 2:
        return parts[1].strip()
    return str(title).strip()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate')
def generate_reports():
    api_key = os.environ.get('BOLDSIGN_API_KEY')
    if not api_key:
        flash('API Key missing. Please configure BOLDSIGN_API_KEY.', 'error')
        return redirect(url_for('index'))

    docs = fetch_documents()
    filtered_docs = []
    
    for doc in docs:
        if doc.get('displayStatus') == 'Waiting for me':
            title = doc.get('messageTitle', 'Unknown')
            timestamp = doc.get('createdDate')
            time_str = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S') if timestamp else 'Unknown'
            
            filtered_docs.append({
                'Title': title,
                'Time': time_str
            })
            
    if not filtered_docs:
        flash('No documents found with status "Waiting for me".', 'warning')
        return redirect(url_for('index'))

    df = pd.DataFrame(filtered_docs)
    
    # Process duplicates
    df['INN Code'] = df['Title'].apply(extract_inn_code)
    duplicates_mask = df.duplicated(subset=['INN Code'], keep=False)
    duplicates_df = df[duplicates_mask].copy()
    if not duplicates_df.empty:
        duplicates_df = duplicates_df.sort_values(by=['INN Code', 'Time'])

    # Create Zip File in memory
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        
        # Write Main Document
        main_excel = io.BytesIO()
        with pd.ExcelWriter(main_excel, engine='openpyxl') as writer:
            df.drop(columns=['INN Code']).to_excel(writer, index=False)
        zf.writestr('Waiting_for_Me_Documents.xlsx', main_excel.getvalue())
        
        # Write Duplicates Document if they exist
        if not duplicates_df.empty:
            dup_excel = io.BytesIO()
            with pd.ExcelWriter(dup_excel, engine='openpyxl') as writer:
                duplicates_df.to_excel(writer, index=False)
            zf.writestr('Duplicates_Waiting_for_Me.xlsx', dup_excel.getvalue())

    memory_file.seek(0)
    
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M")
    
    return send_file(
        memory_file,
        mimetype='application/zip',
        as_attachment=True,
        download_name=f'BoldSign_Reports_{timestamp_str}.zip'
    )

# Render expects the application to run on 0.0.0.0 and dynamically assign the port
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
