import pandas as pd
import matplotlib.pyplot as plt
import io
import base64
import os
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

# Create uploads folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            return render_template('index.html', error='No file part')
        
        file = request.files['file']
        
        # If user does not select file, browser also submits an empty part without filename
        if file.filename == '':
            return render_template('index.html', error='No selected file')
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Read the Excel file
            try:
                df = pd.read_excel(filepath)
                columns = df.columns.tolist()
                return render_template('select_field.html', columns=columns, filename=filename)
            except Exception as e:
                return render_template('index.html', error=f'Error reading Excel file: {str(e)}')
        else:
            return render_template('index.html', error='File type not allowed. Please upload an Excel file (.xlsx or .xls)')
    
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    filename = request.form.get('filename')
    selected_field = request.form.get('field')
    
    if not filename or not selected_field:
        return render_template('index.html', error='Missing filename or field selection')
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    try:
        # Read the Excel file
        df = pd.read_excel(filepath)
        
        # Create frequency table
        freq_table = df[selected_field].value_counts().reset_index()
        freq_table.columns = [selected_field, 'Frequency']
        freq_table['Percentage'] = (freq_table['Frequency'] / freq_table['Frequency'].sum() * 100).round(2)
        
        # Create visualizations
        plt.figure(figsize=(10, 6))
        plt.bar(freq_table[selected_field].astype(str), freq_table['Frequency'])
        plt.title(f'Bar Chart of {selected_field}')
        plt.xlabel(selected_field)
        plt.ylabel('Frequency')
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        # Save bar chart to a bytes buffer
        bar_chart_buffer = io.BytesIO()
        plt.savefig(bar_chart_buffer, format='png')
        plt.close()
        
        # Create pie chart
        plt.figure(figsize=(10, 6))
        plt.pie(freq_table['Frequency'], labels=freq_table[selected_field].astype(str), autopct='%1.1f%%')
        plt.title(f'Pie Chart of {selected_field}')
        plt.axis('equal')
        plt.tight_layout()
        
        # Save pie chart to a bytes buffer
        pie_chart_buffer = io.BytesIO()
        plt.savefig(pie_chart_buffer, format='png')
        plt.close()
        
        # Create Excel file with results
        output_file = io.BytesIO()
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Write frequency table
            freq_table.to_excel(writer, sheet_name='Results', index=False, startrow=1, startcol=1)
            
            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Results']
            
            # Insert bar chart
            bar_chart_buffer.seek(0)
            worksheet.insert_image('A10', 'bar_chart.png', {'image_data': bar_chart_buffer})
            
            # Insert pie chart
            pie_chart_buffer.seek(0)
            worksheet.insert_image('A30', 'pie_chart.png', {'image_data': pie_chart_buffer})
            
            # Add title
            title_format = workbook.add_format({'bold': True, 'font_size': 16})
            worksheet.write('B1', f'Analysis of {selected_field}', title_format)
        
        # Prepare the file for download
        output_file.seek(0)
        return send_file(
            output_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'analysis_{selected_field}.xlsx'
        )
        
    except Exception as e:
        return render_template('index.html', error=f'Error during analysis: {str(e)}')

if __name__ == '__main__':
    app.run(debug=True)