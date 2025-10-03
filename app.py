from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename
from rapidfuzz import fuzz, process
import json
from datetime import datetime
import re
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
        """Compare TOEFL format name with base name and return score 0-100 (weighted average)."""
        toefl_parsed = self.parse_toefl_name(toefl_name)
        base_parsed = self.parse_base_name(base_name)

        # Raw component scores
        firstname_score = None
        lastname_score = None
        full_name_scores = []

        # First/last name components
        if toefl_parsed['firstname'] and base_parsed['firstname']:
            firstname_score = self._calculate_similarity(
                toefl_parsed['firstname'], base_parsed['firstname'], algorithm
            )

        if toefl_parsed['lastname'] and base_parsed['lastname']:
            lastname_score = self._calculate_similarity(
                toefl_parsed['lastname'], base_parsed['lastname'], algorithm
            )

        # Full name variants
        comparisons = [
            (toefl_parsed['full_name_normal'], base_parsed['full_name']),
            (toefl_parsed['full_name_reverse'], base_parsed['full_name']),
            (f"{toefl_parsed['firstname']} {toefl_parsed['lastname']}", base_parsed['full_name']),
            (f"{toefl_parsed['lastname']} {toefl_parsed['firstname']}", base_parsed['full_name']),
        ]

        for toefl_variant, base_full in comparisons:
            if toefl_variant and base_full:
                full_name_scores.append(
                    self._calculate_similarity(toefl_variant, base_full, algorithm)
                )

        max_full = max(full_name_scores) if full_name_scores else None

        # Weighted average with dynamic weights based on available components
        weights = []
        values = []

        if firstname_score is not None:
            weights.append(0.4)
            values.append(firstname_score)
        if lastname_score is not None:
            weights.append(0.4)
            values.append(lastname_score)
        if max_full is not None:
            weights.append(0.2)
            values.append(max_full)

        if values:
            total_weight = sum(weights)
            weighted_avg = sum(v * w for v, w in zip(values, weights)) / total_weight
        else:
            weighted_avg = 0

        # Space-insensitive comparison to handle concatenated names (e.g., "oliveiraclimenia")
        import re
        compact_toefl = re.sub(r"\s+", "", self.normalize_name(toefl_name))
        compact_base = re.sub(r"\s+", "", base_parsed.get('full_name', ''))
        compact_score = self._calculate_similarity(compact_toefl, compact_base, 'ratio') if compact_base else 0

        # Reforço baseado em tokens para lidar com casos de apenas um nome coincidente
        toefl_tokens = set(self.normalize_name(toefl_name).split())
        base_tokens = set(base_parsed.get('parts', []))
        overlap = toefl_tokens.intersection(base_tokens)
        jaccard_score = (len(overlap) / max(1, len(base_tokens))) * 100

        max_first_token_score = 0
        max_last_token_score = 0
        base_first = base_parsed.get('firstname', '')
        base_last = base_parsed.get('lastname', '')
        if base_first:
            for tok in toefl_tokens:
                s = self._calculate_similarity(base_first, tok, algorithm)
                if s > max_first_token_score:
                    max_first_token_score = s
        if base_last:
            for tok in toefl_tokens:
                s = self._calculate_similarity(base_last, tok, algorithm)
                if s > max_last_token_score:
                    max_last_token_score = s

        final_score = max(
            weighted_avg,
            max_full if max_full is not None else 0,
            jaccard_score,
            max_first_token_score,
            max_last_token_score,
            compact_score,
        )
        # Penalizar correspondências de um único token quando o nome TOEFL tem 2+ tokens
        # Mas evitar penalidade quando a comparação sem espaços indicar alta similaridade
        compact_match_high = compact_score >= 90
        if len(toefl_tokens) >= 2 and len(overlap) < 2 and not compact_match_high:
            final_score = min(final_score, 75)
        return final_score
    
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
        
        # Salvar arquivos temporariamente com nomes fixos para posterior comparação
        filename1 = secure_filename(file1.filename)
        filename2 = secure_filename(file2.filename)
        
        ext1 = os.path.splitext(filename1)[1].lower()
        ext2 = os.path.splitext(filename2)[1].lower()
        
        filepath1 = os.path.join(app.config['UPLOAD_FOLDER'], f'file1{ext1}')
        filepath2 = os.path.join(app.config['UPLOAD_FOLDER'], f'file2{ext2}')
        
        file1.save(filepath1)
        file2.save(filepath2)
        
        # Read the uploaded files
        try:
            # First file is the base file with names (column A) and classes (column B)
            if ext1 == '.csv':
                df1 = pd.read_csv(filepath1)
            else:
                df1 = pd.read_excel(filepath1)
            # Second file contains TOEFL students names for comparison
            if ext2 == '.csv':
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
        threshold = float(data.get('threshold', 80))  # threshold em escala 0-100
        algorithm = data.get('algorithm', 'token_sort_ratio')
        column1 = data.get('column1')  # coluna de nomes na planilha base
        column2 = data.get('column2')  # coluna de nomes na planilha TOEFL
        
        # Função auxiliar para localizar arquivos por base name e extensões suportadas
        def find_uploaded_file(base_name):
            for ext in ['.xlsx', '.xls', '.csv']:
                path = os.path.join(app.config['UPLOAD_FOLDER'], f'{base_name}{ext}')
                if os.path.exists(path):
                    return path, ext
            return None, None
        
        file1_path, ext1 = find_uploaded_file('file1')
        file2_path, ext2 = find_uploaded_file('file2')
        
        if not file1_path or not file2_path:
            return jsonify({'success': False, 'error': 'Arquivos não encontrados. Faça o upload novamente.'})
        
        # Read the uploaded files
        try:
            # Ler planilha base: todas as abas (Sheet1..N) quando for Excel
            if ext1 in ['.xlsx', '.xls']:
                df1_sheets = pd.read_excel(file1_path, sheet_name=None)
            else:
                df1_sheets = {'CSV': pd.read_csv(file1_path)}

            # Ler planilha TOEFL (aba única)
            df2 = pd.read_excel(file2_path) if ext2 in ['.xlsx', '.xls'] else pd.read_csv(file2_path)

            # Inicializar comparador/normalizador para apoiar filtros e deduplicação
            comparator = NameComparator()

            # Agregar nomes, turmas, professor e nível da base a partir de todas as abas
            base_names = []
            base_classes = []
            base_professors = []
            base_levels = []
            for sheet_name, df in df1_sheets.items():
                # Não ignorar nenhuma linha: usar todas as linhas da aba
                df1_data = df

                # Determinar colunas de nomes e turmas
                if column1 and column1 in df1_data.columns:
                    names_col = column1
                else:
                    names_col = df1_data.columns[0]

                # Detectar dinamicamente TODAS as colunas relacionadas à turma/serie/letra (NÃO inclui nível)
                class_keywords = {'turma','classe','serie','série','ano','grau','turno','sala'}
                class_cols = []
                for col in df1_data.columns:
                    norm_col = comparator.normalize_name(str(col))
                    if any(kw in norm_col for kw in class_keywords):
                        class_cols.append(col)

                # Escolher coluna primária de turma/classe quando disponível
                primary_class_col = None
                for col in df1_data.columns:
                    norm_col = comparator.normalize_name(str(col))
                    if any(kw in norm_col for kw in {'turma','classe'}):
                        primary_class_col = col
                        break
                if primary_class_col is None and class_cols:
                    primary_class_col = class_cols[0]

                # Detectar coluna de Professor e Nível
                professor_col = None
                nivel_col = None
                for col in df1_data.columns:
                    norm_col = comparator.normalize_name(str(col))
                    if professor_col is None and any(kw in norm_col for kw in {'professor','docente','prof','teacher'}):
                        professor_col = col
                    if nivel_col is None and any(kw in norm_col for kw in {'nivel','nível'}):
                        nivel_col = col
                # Fallbacks explícitos: Coluna C=professor, D=nivel
                if professor_col is None and len(df1_data.columns) >= 3:
                    professor_col = df1_data.columns[2]
                if nivel_col is None and len(df1_data.columns) >= 4:
                    nivel_col = df1_data.columns[3]

                # Subconjunto e limpeza (inclui todas as colunas de classe encontradas)
                subset_cols = [names_col] + class_cols + ([professor_col] if professor_col else []) + ([nivel_col] if nivel_col else [])
                df_sub = df1_data[subset_cols].copy()
                df_sub = df_sub[df_sub[names_col].notna()]

                # Construir um campo normalizado combinando possíveis colunas de classe + nome da aba
                sheet_norm = comparator.normalize_name(sheet_name)
                def build_class_norm(row):
                    parts = []
                    for c in class_cols:
                        val = str(row[c]) if c in row else ''
                        parts.append(val)
                    parts.append(sheet_norm)
                    combo = ' '.join([p for p in parts if p and p != 'nan'])
                    return comparator.normalize_name(combo)
                df_sub['__class_norm__'] = df_sub.apply(build_class_norm, axis=1)
                # Guardar turma crua priorizando coluna primária
                if primary_class_col:
                    df_sub['__class_raw__'] = df_sub[primary_class_col].astype(str).fillna('')
                else:
                    df_sub['__class_raw__'] = ''
                # Guardar professor e nível como metadados diretos
                if professor_col:
                    df_sub['__professor__'] = df_sub[professor_col].astype(str).fillna('')
                else:
                    df_sub['__professor__'] = ''
                if nivel_col:
                    df_sub['__nivel__'] = df_sub[nivel_col].astype(str).fillna('')
                else:
                    df_sub['__nivel__'] = ''

                # Filtrar extracurriculares e manter entradas relevantes
                before_count = len(df_sub)
                # Excluir termos extracurriculares comuns (normalizados, sem acentos)
                banned_terms = [
                    'violino','danca','teatro','musica','ballet','coral','flauta','piano','canto',
                    'judo','capoeira','arte','artes','basquete','futsal','handebol','volei','xadrez'
                ]
                pattern = '|'.join(banned_terms)
                df_sub = df_sub[~df_sub['__class_norm__'].str.contains(pattern, na=False)]
                after_count = len(df_sub)
                print(f"[DEBUG] Sheet='{sheet_name}' extracurricular filter: kept {after_count}/{before_count}")

                # Limpar rótulo da turma para mostrar FUND-<numero><letra>, usando o texto combinado
                roman_map = {
                    'i': '1','ii': '2','iii': '3','iv': '4','v': '5',
                    'vi': '6','vii': '7','viii': '8','ix': '9','x': '10'
                }

                # Normaliza a visualização do Nível para formato fechado (ex.: 6.1/6.2/6.3)
                def normalize_nivel_display(raw_nivel):
                    if raw_nivel is None:
                        return ''
                    s = str(raw_nivel).strip().lower()
                    # Padrões com separador (.,;:-)
                    m = re.match(r"^\s*(\d{1,2})\s*[\.,;:-]\s*(\d+)\s*$", s)
                    if m:
                        major = m.group(1)
                        minor_first = m.group(2)[0]  # primeiro dígito após o separador
                        # Exibir apenas quando minor ∈ {1,2,3}
                        if minor_first in {'1','2','3'}:
                            return f"{major}.{minor_first}"
                        else:
                            return ''
                    # Inteiro simples
                    m2 = re.match(r"^\s*(\d{1,2})\s*$", s)
                    if m2:
                        # Não exibir inteiro puro para 6º ano fechado; deixar vazio
                        return ''
                    return str(raw_nivel)

                def clean_fund_label(norm, raw_class=None, raw_nivel=None, sheet_norm=None):

                    def extract_num_letter(text):
                        # 6 a | 6º a | 6° a | 6 ano a | 6-a | 6a | 6 a manha
                        pats = [
                            r"\b(\d{1,2})\s*(?:ano|serie|grau|º|°)?\s*([a-z])\b",
                            r"\b(\d{1,2})\s*[-_]\s*([a-z])\b",
                            r"\b(\d{1,2}[a-z])\b",
                            # Não usar padrões sem separador; focar em maior.minor apenas
                        ]
                        for p in pats:
                            m = re.search(p, text)
                            if m:
                                if m.lastindex == 2:
                                    num = m.group(1)
                                    let = m.group(2)
                                    # Se 'let' é dígito (padrão 6.1), converter para letra
                                    if re.fullmatch(r"\d", let):
                                        idx = int(let)
                                        if 1 <= idx <= 26:
                                            let = chr(ord('a') + idx - 1)
                                    return num, let
                                else:
                                    g = m.group(1)
                                    num = re.findall(r"\d{1,2}", g)
                                    let = re.findall(r"[a-z]", g)
                                    if num and let:
                                        return num[0], let[0]
                        return None

                    # 0) Tentar primeiro a turma crua vinda da planilha
                    if raw_class:
                        rc_norm = comparator.normalize_name(str(raw_class))
                        res = extract_num_letter(rc_norm)
                        if res:
                            num, let = res
                            return f"FUND-{num}{let.upper()}"

                    # 1) Tenta no texto combinado (todas as colunas de classe)
                    res = extract_num_letter(norm)
                    # Em seguida, tenta diretamente no Nível bruto (sem normalização), preservando pontuação
                    if not res and raw_nivel:
                        raw_nivel_str = str(raw_nivel).strip().lower()
                        # 1) Aceitar apenas formato estrito X.Y com um dígito após o separador
                        m_single = re.match(r"^\s*(\d{1,2})\s*[\.,;:-]\s*(\d)\s*$", raw_nivel_str)
                        if m_single:
                            num = m_single.group(1)
                            minor = m_single.group(2)
                            idx = int(minor)
                            if 1 <= idx <= 26:
                                let = chr(ord('a') + idx - 1)
                                res = (num, let)
                        else:
                            # 2) Se vier com mais de um dígito (ex.: 6.31), usar o primeiro dígito após o separador
                            m_multi = re.match(r"^\s*(\d{1,2})\s*[\.,;:-]\s*(\d{2,})\s*$", raw_nivel_str)
                            if m_multi:
                                num = m_multi.group(1)
                                minor = m_multi.group(2)[0]  # primeiro dígito
                                idx = int(minor)
                                if 1 <= idx <= 26:
                                    let = chr(ord('a') + idx - 1)
                                    res = (num, let)
                            else:
                                # 3) Caso seja apenas número inteiro (ex.: "6"), não acrescentar letra
                                m_int = re.match(r"^\s*(\d{1,2})\s*$", raw_nivel_str)
                                if m_int:
                                    res = (m_int.group(1), '')
                    if not res:
                        # Fallback: tenta no nome da aba
                        res = extract_num_letter(sheet_norm or '')
                    if res:
                        num, let = res
                        return f"FUND-{num}{let.upper()}" if let else f"FUND-{num}"

                    # Fallback: detectar romano
                    m_rom = re.search(r"\bfund[a-z]*\b\s*(i{1,3}|iv|v|vi|vii|viii|ix|x)\b", norm)
                    if not m_rom and sheet_norm:
                        m_rom = re.search(r"\bfund[a-z]*\b\s*(i{1,3}|iv|v|vi|vii|viii|ix|x)\b", sheet_norm)
                    if m_rom:
                        num = roman_map.get(m_rom.group(1), '')
                        return f"FUND-{num}" if num else "FUND"

                    return "FUND"
                # Calcular rótulo limpo FUND usando turma crua da planilha + classe normalizada + Nível bruto
                df_sub['__class_clean__'] = df_sub.apply(
                    lambda row: clean_fund_label(
                        comparator.normalize_name(row['__class_norm__']),
                        row['__class_raw__'],
                        row['__nivel__'],
                        sheet_norm
                    ),
                    axis=1
                )

                # Após limpar, manter apenas linhas de FUND
                df_sub = df_sub[df_sub['__class_clean__'].str.startswith('FUND')]

                # Agregar
                base_names.extend(df_sub[names_col].astype(str).tolist())
                # Usar rótulo limpo FUND (com fallback quando necessário)
                base_classes.extend(df_sub['__class_clean__'].fillna('').astype(str).tolist())
                base_professors.extend(df_sub['__professor__'].astype(str).tolist())
                base_levels.extend(df_sub['__nivel__'].astype(str).tolist())

            # Deduplicar nomes da base por normalização POR TURMA (preserva primeira ocorrência por turma)
            # Isso evita perder alunos presentes em múltiplas turmas (ex.: FUND-6A e FUND-6B).
            seen = set()
            dedup_names = []
            dedup_classes = []
            dedup_professors = []
            dedup_levels = []
            for i, name in enumerate(base_names):
                key = comparator.normalize_name(str(name))
                cls = base_classes[i] if i < len(base_classes) else ''
                comp_key = (key, cls)
                if comp_key not in seen:
                    seen.add(comp_key)
                    dedup_names.append(name)
                    dedup_classes.append(cls)
                    dedup_professors.append(base_professors[i] if i < len(base_professors) else '')
                    dedup_levels.append(base_levels[i] if i < len(base_levels) else '')
            base_names = dedup_names
            base_classes = dedup_classes
            base_professors = dedup_professors
            base_levels = dedup_levels

            # Obter nomes TOEFL
            if column2 and column2 in df2.columns:
                toefl_names = df2[column2].dropna().astype(str).tolist()
            else:
                toefl_names = df2.iloc[:, 0].dropna().astype(str).tolist()  # padrão: primeira coluna
            
            
        # Perform comparison
            results = []
            suggestions = []
            # Amostras de diagnóstico: primeiros 5 itens
            debug_limit = 5
            print(f"[DEBUG] Base nomes: {len(base_names)} | TOEFL nomes: {len(toefl_names)} | threshold={threshold} | algorithm={algorithm}")
            for i, toefl_name in enumerate(toefl_names):
                best_match = None
                best_score = 0
                best_class = ''
                best_professor = ''
                best_nivel = ''
                cand_scores = []
                
                for j, base_name in enumerate(base_names):
                    # Compare TOEFL name with base name
                    score = comparator.compare_names(toefl_name, base_name, algorithm)
                    # Collect candidate scores for suggestions
                    cand_scores.append((
                        base_name,
                        base_classes[j] if j < len(base_classes) else '',
                        base_professors[j] if j < len(base_professors) else '',
                        base_levels[j] if j < len(base_levels) else '',
                        score
                    ))
                    
                    if score >= threshold and score > best_score:
                        best_match = base_name
                        best_score = score
                        best_class = base_classes[j] if j < len(base_classes) else ''
                        best_professor = base_professors[j] if j < len(base_professors) else ''
                        best_nivel = base_levels[j] if j < len(base_levels) else ''
                
                # Log diagnóstico para os primeiros itens
                if i < debug_limit:
                    # Mesmo se não atingir o limiar, recalcula melhor absoluto para entender proximidade
                    abs_best_score = 0
                    abs_best_match = None
                    abs_best_class = ''
                    for j, base_name in enumerate(base_names):
                        s = comparator.compare_names(toefl_name, base_name, algorithm)
                        if s > abs_best_score:
                            abs_best_score = s
                            abs_best_match = base_name
                            abs_best_class = base_classes[j] if j < len(base_classes) else ''
                    print(f"[DEBUG] TOEFL='{toefl_name}' | best_above_threshold={best_score} | abs_best={abs_best_score} -> '{abs_best_match}' turma='{abs_best_class}'")

                if best_match:
                    results.append({
                        'toefl_name': toefl_name,
                        'matched_name': best_match,
                        'class': best_class,
                        'professor': best_professor,
                        'nivel': normalize_nivel_display(best_nivel),
                        'score': round(best_score, 2)
                    })
                else:
                    # Build suggestions (top 3 by score) when no match above threshold
                    if cand_scores:
                        cand_scores.sort(key=lambda x: x[4], reverse=True)
                        top = cand_scores[:3]
                        suggestions.append({
                            'toefl_name': toefl_name,
                            'candidates': [
                                {
                                    'name': n,
                                    'class': c,
                                    'professor': p,
                                    'nivel': normalize_nivel_display(lv),
                                    'score': round(s, 2)
                                } for (n, c, p, lv, s) in top
                            ]
                        })
            
            # Calcular lista de não encontrados
            matched_toefl_set = set([r['toefl_name'] for r in results])
            unmatched_list = [name for name in toefl_names if name not in matched_toefl_set]

            # Calcular estatísticas
            total_toefl = len(toefl_names)
            matched_count = len(results)
            unmatched_count = total_toefl - matched_count
            match_percentage = (matched_count / total_toefl * 100) if total_toefl > 0 else 0
            print(f"[DEBUG] matched={matched_count}/{total_toefl} ({round(match_percentage,2)}%)")
            if matched_count == 0:
                print("[DEBUG] Nenhuma correspondência acima do limiar. Sugestões: reduzir limiar para 60–70; testar algoritmo 'token_set_ratio'; confirmar colunas de nomes.")
            
            return jsonify({
                'success': True,
                'results': results,
                'unmatched_list': unmatched_list,
                'suggestions': suggestions,
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
        unmatched = data.get('unmatched_list', [])
        
        # Criar DataFrame para exportação (sem coluna de Similaridade)
        export_rows = []
        for r in results:
            export_rows.append({
                'Nome': r.get('toefl_name', ''),
                'Nome Completo': r.get('matched_name', ''),
                'Turma': r.get('class', ''),
                'Professor': r.get('professor', ''),
                'Nivel': r.get('nivel', '')
            })
        
        df_export = pd.DataFrame(export_rows)
        df_unmatched = pd.DataFrame({'Alunos_Nao_Encontrados': unmatched})
        
        # Criar arquivo Excel em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Escrever resultados
            sheet_name = 'Resultados_Comparacao'
            df_export.to_excel(writer, sheet_name=sheet_name, index=False)

            # Se houver não encontrados, escrever ao final da mesma aba
            if not df_unmatched.empty:
                start_row = len(df_export) + 3  # duas linhas em branco + cabeçalho
                df_unmatched.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
        
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
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))