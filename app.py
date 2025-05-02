import pandas as pd
import matplotlib.pyplot as plt
import io
import base64
import os
import math
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}
app.config['ROWS_PER_PAGE'] = 10  # Number of rows to display per page

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
                
                # Get preview data (first page)
                preview_data = df.head(app.config['ROWS_PER_PAGE']).to_dict('records')
                total_rows = len(df)
                total_pages = math.ceil(total_rows / app.config['ROWS_PER_PAGE'])
                
                return render_template('select_field.html', 
                                      columns=columns, 
                                      filename=filename, 
                                      preview_data=preview_data,
                                      total_rows=total_rows,
                                      total_pages=total_pages,
                                      current_page=1)
            except Exception as e:
                return render_template('index.html', error=f'Error reading Excel file: {str(e)}')
        else:
            return render_template('index.html', error='File type not allowed. Please upload an Excel file (.xlsx or .xls)')
    
    return render_template('index.html')

@app.route('/preview/<filename>/<int:page>', methods=['GET'])
def preview_data(filename, page):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    try:
        df = pd.read_excel(filepath)
        rows_per_page = app.config['ROWS_PER_PAGE']
        start_idx = (page - 1) * rows_per_page
        end_idx = start_idx + rows_per_page
        
        page_data = df.iloc[start_idx:end_idx].to_dict('records')
        total_rows = len(df)
        total_pages = math.ceil(total_rows / rows_per_page)
        
        return jsonify({
            'data': page_data,
            'columns': df.columns.tolist(),
            'total_pages': total_pages,
            'current_page': page
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

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
        bar_chart_buffer.seek(0)
        bar_chart_data = base64.b64encode(bar_chart_buffer.getvalue()).decode('utf-8')
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
        pie_chart_buffer.seek(0)
        pie_chart_data = base64.b64encode(pie_chart_buffer.getvalue()).decode('utf-8')
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
        
        # Get preview data (first page)
        preview_data = freq_table.head(app.config['ROWS_PER_PAGE']).to_dict('records')
        total_rows = len(freq_table)
        total_pages = math.ceil(total_rows / app.config['ROWS_PER_PAGE'])
        
        return render_template('results.html',
                              selected_field=selected_field,
                              preview_data=preview_data,
                              total_rows=total_rows,
                              total_pages=total_pages,
                              current_page=1,
                              bar_chart=bar_chart_data,
                              pie_chart=pie_chart_data,
                              filename=filename)
    except Exception as e:
        return render_template('index.html', error=f'Error during analysis: {str(e)}')

@app.route('/download_results/<filename>/<field>', methods=['GET'])
def download_results(filename, field):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    try:
        # Read the Excel file
        df = pd.read_excel(filepath)
        
        # Create frequency table
        freq_table = df[field].value_counts().reset_index()
        freq_table.columns = [field, 'Frequency']
        freq_table['Percentage'] = (freq_table['Frequency'] / freq_table['Frequency'].sum() * 100).round(2)
        
        # Create visualizations (same as in analyze function)
        plt.figure(figsize=(10, 6))
        plt.bar(freq_table[field].astype(str), freq_table['Frequency'])
        plt.title(f'Bar Chart of {field}')
        plt.xlabel(field)
        plt.ylabel('Frequency')
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        bar_chart_buffer = io.BytesIO()
        plt.savefig(bar_chart_buffer, format='png')
        plt.close()
        
        plt.figure(figsize=(10, 6))
        plt.pie(freq_table['Frequency'], labels=freq_table[field].astype(str), autopct='%1.1f%%')
        plt.title(f'Pie Chart of {field}')
        plt.axis('equal')
        plt.tight_layout()
        
        pie_chart_buffer = io.BytesIO()
        plt.savefig(pie_chart_buffer, format='png')
        plt.close()
        
        # Create Excel file with results
        output_file = io.BytesIO()
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            freq_table.to_excel(writer, sheet_name='Results', index=False, startrow=1, startcol=1)
            
            workbook = writer.book
            worksheet = writer.sheets['Results']
            
            bar_chart_buffer.seek(0)
            worksheet.insert_image('A10', 'bar_chart.png', {'image_data': bar_chart_buffer})
            
            pie_chart_buffer.seek(0)
            worksheet.insert_image('A30', 'pie_chart.png', {'image_data': pie_chart_buffer})
            
            title_format = workbook.add_format({'bold': True, 'font_size': 16})
            worksheet.write('B1', f'Analysis of {field}', title_format)
        
        output_file.seek(0)
        return send_file(
            output_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'analysis_{field}.xlsx'
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/results_data/<filename>/<field>/<int:page>', methods=['GET'])
def results_data(filename, field, page):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    try:
        df = pd.read_excel(filepath)
        freq_table = df[field].value_counts().reset_index()
        freq_table.columns = [field, 'Frequency']
        freq_table['Percentage'] = (freq_table['Frequency'] / freq_table['Frequency'].sum() * 100).round(2)
        
        rows_per_page = app.config['ROWS_PER_PAGE']
        start_idx = (page - 1) * rows_per_page
        end_idx = min(start_idx + rows_per_page, len(freq_table))
        
        page_data = freq_table.iloc[start_idx:end_idx].to_dict('records')
        total_rows = len(freq_table)
        total_pages = math.ceil(total_rows / rows_per_page)
        
        return jsonify({
            'data': page_data,
            'total_pages': total_pages,
            'current_page': page,
            'field': field
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# At the bottom of your app.py file, change:
#if __name__ == '__main__':
 #   app.run(debug=True)

# To:
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)