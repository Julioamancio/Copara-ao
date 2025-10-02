from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename
from rapidfuzz import fuzz, process
import json
from datetime import datetime
import tempfile
import io

app = Flask(__name__)
app.config['SECRET_KEY'] = 'sua-chave-secreta-aqui'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Criar pasta de uploads se não existir
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Extensões permitidas
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

class NameComparator:
    def __init__(self):
        # Brazilian Portuguese specific configurations
        self.stopwords = {'de', 'da', 'do', 'dos', 'das', 'e', 'del', 'la', 'el', 'von', 'van'}
        
    def normalize_name(self, name):
        """Normalize name for better comparison"""
        if pd.isna(name):
            return ""
        
        # Convert to string and normalize
        name = str(name).strip()
        
        # Remove extra spaces and convert to lowercase
        name = ' '.join(name.split()).lower()
        
        # Remove accents - simple replacement for common Portuguese accents
        replacements = {
            'á': 'a', 'à': 'a', 'ã': 'a', 'â': 'a',
            'é': 'e', 'ê': 'e',
            'í': 'i',
            'ó': 'o', 'ô': 'o', 'õ': 'o',
            'ú': 'u', 'ü': 'u',
            'ç': 'c'
        }
        
        for old, new in replacements.items():
            name = name.replace(old, new)
        
        # Remove punctuation except commas (important for TOEFL format)
        import re
        name = re.sub(r'[^\w\s,]', '', name)
        
        return name
    
    def parse_toefl_name(self, toefl_name):
        """Parse TOEFL format name (LASTNAME, FIRSTNAME [MIDDLE])"""
        normalized = self.normalize_name(toefl_name)
        
        if ',' in normalized:
            parts = normalized.split(',')
            if len(parts) >= 2:
                lastname = parts[0].strip()
                firstname_parts = parts[1].strip().split()
                firstname = firstname_parts[0] if firstname_parts else ''
                
                # Return both possible combinations
                return {
                    'lastname': lastname,
                    'firstname': firstname,
                    'full_name_normal': f"{firstname} {lastname}",
                    'full_name_reverse': f"{lastname} {firstname}"
                }
        
        # If no comma, treat as regular name
        return {
            'lastname': '',
            'firstname': normalized,
            'full_name_normal': normalized,
            'full_name_reverse': normalized
        }
    
    def parse_base_name(self, base_name):
        """Parse base name (full name format)"""
        normalized = self.normalize_name(base_name)
        parts = normalized.split()
        
        if len(parts) >= 2:
            # Remove stopwords
            filtered_parts = [part for part in parts if part not in self.stopwords]
            
            if len(filtered_parts) >= 2:
                firstname = filtered_parts[0]
                lastname = filtered_parts[-1]
                
                return {
                    'firstname': firstname,
                    'lastname': lastname,
                    'full_name': normalized,
                    'parts': filtered_parts
                }
        
        return {
            'firstname': normalized,
            'lastname': '',
            'full_name': normalized,
            'parts': parts
        }
    
    def compare_names(self, toefl_name, base_name, algorithm='token_sort_ratio'):
        """Compare TOEFL format name with base name"""
        toefl_parsed = self.parse_toefl_name(toefl_name)
        base_parsed = self.parse_base_name(base_name)
        
        scores = []
        
        # Compare different combinations
        comparisons = [
            # Direct full name comparisons
            (toefl_parsed['full_name_normal'], base_parsed['full_name']),
            (toefl_parsed['full_name_reverse'], base_parsed['full_name']),
            
            # First name + Last name comparisons
            (f"{toefl_parsed['firstname']} {toefl_parsed['lastname']}", base_parsed['full_name']),
            (f"{toefl_parsed['lastname']} {toefl_parsed['firstname']}", base_parsed['full_name']),
        ]
        
        # Individual name part comparisons (higher weight)
        if toefl_parsed['firstname'] and base_parsed['firstname']:
            firstname_score = self._calculate_similarity(toefl_parsed['firstname'], base_parsed['firstname'], algorithm)
            scores.append(firstname_score * 0.4)  # 40% weight for first name match
        
        if toefl_parsed['lastname'] and base_parsed['lastname']:
            lastname_score = self._calculate_similarity(toefl_parsed['lastname'], base_parsed['lastname'], algorithm)
            scores.append(lastname_score * 0.4)  # 40% weight for last name match
        
        # Full name comparisons (lower individual weight)
        for toefl_variant, base_full in comparisons:
            if toefl_variant and base_full:
                score = self._calculate_similarity(toefl_variant, base_full, algorithm)
                scores.append(score * 0.2)  # 20% weight for full name variants
        
        # Return the highest score
        return max(scores) if scores else 0
    
    def _calculate_similarity(self, str1, str2, algorithm='token_sort_ratio'):
        """Calculate similarity between two strings"""
        if not str1 or not str2:
            return 0
        
        if algorithm == 'ratio':
            return fuzz.ratio(str1, str2)
        elif algorithm == 'partial_ratio':
            return fuzz.partial_ratio(str1, str2)
        elif algorithm == 'token_sort_ratio':
            return fuzz.token_sort_ratio(str1, str2)
        elif algorithm == 'token_set_ratio':
            return fuzz.token_set_ratio(str1, str2)
        else:
            return fuzz.token_sort_ratio(str1, str2)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        if 'file1' not in request.files or 'file2' not in request.files:
            return jsonify({'error': 'Ambos os arquivos são obrigatórios'}), 400
        
        file1 = request.files['file1']
        file2 = request.files['file2']
        
        if file1.filename == '' or file2.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        if not (allowed_file(file1.filename) and allowed_file(file2.filename)):
            return jsonify({'error': 'Formato de arquivo não suportado'}), 400
        
        # Salvar arquivos temporariamente
        filename1 = secure_filename(file1.filename)
        filename2 = secure_filename(file2.filename)
        
        filepath1 = os.path.join(app.config['UPLOAD_FOLDER'], filename1)
        filepath2 = os.path.join(app.config['UPLOAD_FOLDER'], filename2)
        
        file1.save(filepath1)
        file2.save(filepath2)
        
        # Read the uploaded files
        try:
            # First file is the base file with names (column A) and classes (column B)
            if filename1.endswith('.csv'):
                df1 = pd.read_csv(filepath1)
            else:
                df1 = pd.read_excel(filepath1)
            # Second file contains TOEFL students names for comparison
            if filename2.endswith('.csv'):
                df2 = pd.read_csv(filepath2)
            else:
                df2 = pd.read_excel(filepath2)
        except Exception as e:
            return jsonify({'error': f'Erro ao ler planilhas: {str(e)}'}), 400
        
        # Retornar informações das planilhas
        response = {
            'success': True,
            'file1_info': {
                'name': filename1,
                'rows': len(df1),
                'columns': list(df1.columns)
            },
            'file2_info': {
                'name': filename2,
                'rows': len(df2),
                'columns': list(df2.columns)
            }
        }
        
        return jsonify(response)
        
    except Exception as e:
        return jsonify({'error': f'Erro interno: {str(e)}'}), 500

@app.route('/compare', methods=['POST'])
def compare_names():
    try:
        data = request.get_json()
        threshold = float(data.get('threshold', 80)) / 100
        algorithm = data.get('algorithm', 'levenshtein')
        
        # Read the uploaded files directly without asking for columns
        file1_path = os.path.join(app.config['UPLOAD_FOLDER'], 'file1.xlsx')
        file2_path = os.path.join(app.config['UPLOAD_FOLDER'], 'file2.xlsx')
        
        if not os.path.exists(file1_path) or not os.path.exists(file2_path):
            return jsonify({'success': False, 'error': 'Arquivos não encontrados. Faça o upload novamente.'})
        
        # Read the uploaded files
        try:
            # Base file: Column A = Names, Column B = Classes
            df1 = pd.read_excel(file1_path)
            # TOEFL file: Column A = Names (format: LASTNAME, FIRSTNAME or LASTNAME, FIRSTNAME MIDDLE)
            df2 = pd.read_excel(file2_path)
            
            # Get names and classes from base file (columns A and B)
            base_names = df1.iloc[:, 0].dropna().astype(str).tolist()  # Column A: Full names
            base_classes = df1.iloc[:, 1].fillna('').astype(str).tolist() if df1.shape[1] > 1 else [''] * len(base_names)  # Column B: Classes
            
            # Get TOEFL names (column A only)
            toefl_names = df2.iloc[:, 0].dropna().astype(str).tolist()  # Column A: Names in TOEFL format
            
            # Initialize comparator
            comparator = NameComparator()
            
            # Perform comparison
            results = []
            for i, toefl_name in enumerate(toefl_names):
                best_match = None
                best_score = 0
                best_class = ''
                
                for j, base_name in enumerate(base_names):
                     # Compare TOEFL name with base name
                     score = comparator.compare_names(toefl_name, base_name, algorithm)
                     
                     if score >= (threshold * 100) and score > best_score:
                         best_match = base_name
                         best_score = score
                         best_class = base_classes[j] if j < len(base_classes) else ''
                 
                 if best_match:
                     results.append({
                         'toefl_name': toefl_name,
                         'matched_name': best_match,
                         'class': best_class,
                         'score': round(best_score, 2)
                     })
            
            # Calculate statistics
            total_toefl = len(toefl_names)
            matched_count = len(results)
            unmatched_count = total_toefl - matched_count
            match_percentage = (matched_count / total_toefl * 100) if total_toefl > 0 else 0
            
            return jsonify({
                'success': True,
                'results': results,
                'statistics': {
                    'total_toefl': total_toefl,
                    'matched': matched_count,
                    'unmatched': unmatched_count,
                    'match_percentage': round(match_percentage, 2)
                }
            })
            
        except Exception as e:
            return jsonify({'success': False, 'error': f'Erro ao processar planilhas: {str(e)}'})
            
    except Exception as e:
        return jsonify({'success': False, 'error': f'Erro na comparação: {str(e)}'})

@app.route('/export', methods=['POST'])
def export_results():
    try:
        data = request.get_json()
        results = data.get('results', [])
        
        # Criar DataFrame para exportação
        export_data = []
        
        for result in results:
            base_row = {
                'Nome_Planilha1': result['nome_planilha1'],
                'Index_Planilha1': result['index_planilha1'],
                'Total_Matches': result['total_matches']
            }
            
            if result['matches']:
                for i, match in enumerate(result['matches'][:3]):  # Top 3 matches
                    row = base_row.copy()
                    row.update({
                        f'Match_{i+1}_Nome': match['nome_planilha2'],
                        f'Match_{i+1}_Index': match['index_planilha2'],
                        f'Match_{i+1}_Score': match['score']
                    })
                    export_data.append(row)
            else:
                export_data.append(base_row)
        
        df_export = pd.DataFrame(export_data)
        
        # Criar arquivo Excel em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, sheet_name='Resultados_Comparacao', index=False)
        
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'comparacao_nomes_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
        
    except Exception as e:
        return jsonify({'error': f'Erro na exportação: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)