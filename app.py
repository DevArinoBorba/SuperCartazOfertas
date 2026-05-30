import csv
import io
import os
import json
import sqlite3
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, g, session, send_from_directory
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'cartazes_dev_secret_key')
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
)
DATABASE = os.path.join(app.root_path, 'ofertas.db')

PASSWORD_HASH_PREFIXES = ('scrypt:', 'pbkdf2:', 'argon2:')
ALLOWED_IMPORT_EXTENSIONS = {'.csv', '.xlsx'}
REQUIRED_IMPORT_COLUMNS = {'produto', 'preco_oferta'}
IMPORT_COLUMN_ALIASES = {
    'produto': 'produto',
    'nome': 'produto',
    'nome_produto': 'produto',
    'produto_nome': 'produto',
    'unidade': 'unidade',
    'und': 'unidade',
    'preco_oferta': 'preco_oferta',
    'preco': 'preco_oferta',
    'oferta': 'preco_oferta',
    'preco_club': 'preco_club',
    'preco_clube': 'preco_club',
    'clube': 'preco_club',
    'data_cadastro': 'data_cadastro',
    'data': 'data_cadastro',
    'data_oferta': 'data_cadastro',
    'data_validade': 'data_validade',
    'validade': 'data_validade',
    'destinatario': 'destinatario',
    'destinatarios': 'destinatario',
    'filiais': 'destinatario',
    'filial': 'destinatario',
}

LOJAS_DISPONIVEIS = ['10', '20', '30', '40']


def is_password_hash(value):
    return bool(value) and value.startswith(PASSWORD_HASH_PREFIXES)


def normalize_key(value):
    value = str(value or '').strip().lower()
    replacements = {
        'á': 'a', 'à': 'a', 'ã': 'a', 'â': 'a',
        'é': 'e', 'ê': 'e',
        'í': 'i',
        'ó': 'o', 'ô': 'o', 'õ': 'o',
        'ú': 'u',
        'ç': 'c',
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    value = ''.join(char if char.isalnum() else '_' for char in value)
    return '_'.join(part for part in value.split('_') if part)


def parse_date_or_default(value, default=None):
    if not value:
        return default or datetime.now().strftime('%Y-%m-%d')
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d')
    if hasattr(value, 'strftime'):
        return value.strftime('%Y-%m-%d')
    value = str(value).strip()
    try:
        datetime.strptime(value, '%Y-%m-%d')
        return value
    except ValueError:
        pass
    for fmt in ('%d/%m/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(value, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return default or datetime.now().strftime('%Y-%m-%d')


def parse_price(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    value = value.strip()
    if not value:
        return None
    return float(value.replace(',', '.'))


def log_audit(action, entity_type, entity_id=None, details=None):
    if 'username' not in session:
        return
    db = get_db()
    db.execute(
        'INSERT INTO audit_log (action, entity_type, entity_id, username, details) VALUES (?, ?, ?, ?, ?)',
        (action, entity_type, entity_id, session['username'], details)
    )
    db.commit()


def get_config(chave, default=''):
    db = get_db()
    resultado = db.execute('SELECT valor FROM config WHERE chave = ?', (chave,)).fetchone()
    return resultado['valor'] if resultado else default


def user_can_view_offer(oferta, loja, role):
    if role == 'admin':
        return True
    destinatario = oferta['destinatario'] or 'todos'
    if destinatario == 'todos':
        return True
    destinatarios = [d.strip() for d in destinatario.split(',') if d.strip()]
    return str(loja) in destinatarios


def insert_oferta(db, produto, unidade, preco_oferta, preco_club, data_cadastro, data_validade, destinatario,
                  layout_custom=None, layout_custom_pequeno=None):
    cursor = db.execute(
        '''
        INSERT INTO ofertas (
            produto, unidade, preco_oferta, preco_club, data_cadastro, data_validade,
            destinatario, layout_custom, layout_custom_pequeno
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''',
        (
            produto, unidade, preco_oferta, preco_club, data_cadastro, data_validade,
            destinatario, layout_custom, layout_custom_pequeno
        )
    )
    return cursor.lastrowid


def normalize_import_row(row):
    normalized = {}
    for key, value in row.items():
        canonical_key = IMPORT_COLUMN_ALIASES.get(normalize_key(key))
        if canonical_key:
            normalized[canonical_key] = value
    return normalized


def read_csv_import(file_storage):
    raw = file_storage.read()
    try:
        text = raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        text = raw.decode('latin-1')
    sample = text[:2048]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=',;')
    except csv.Error:
        dialect = csv.excel
        dialect.delimiter = ';'
    reader = csv.DictReader(io.StringIO(text), dialect=dialect)
    return list(reader)


def read_xlsx_import(file_storage):
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError('Para importar Excel (.xlsx), instale as dependências com: pip install -r requirements.txt') from exc

    workbook = load_workbook(file_storage, read_only=True, data_only=True)
    sheet = workbook.active
    rows = sheet.iter_rows(values_only=True)
    headers = next(rows, None)
    if not headers:
        return []
    result = []
    for row in rows:
        result.append({headers[index]: value for index, value in enumerate(row) if index < len(headers)})
    return result


def build_offer_from_import_row(row):
    row = normalize_import_row(row)
    missing = REQUIRED_IMPORT_COLUMNS - set(key for key, value in row.items() if value not in (None, ''))
    if missing:
        raise ValueError('colunas obrigatórias ausentes: produto e preco_oferta')

    produto = str(row.get('produto', '')).strip()
    unidade = str(row.get('unidade') or 'un').strip()
    preco_oferta = parse_price(row.get('preco_oferta'))
    preco_club = parse_price(row.get('preco_club'))
    data_cadastro = parse_date_or_default(row.get('data_cadastro'))
    data_validade = row.get('data_validade')
    data_validade = parse_date_or_default(data_validade, default='') if data_validade else None
    destinatario = str(row.get('destinatario') or 'todos').strip() or 'todos'

    if not produto:
        raise ValueError('produto vazio')
    if preco_oferta is None or preco_oferta < 0:
        raise ValueError('preço de oferta inválido')
    if preco_club is not None and preco_club < 0:
        raise ValueError('preço clube inválido')

    return {
        'produto': produto,
        'unidade': unidade,
        'preco_oferta': preco_oferta,
        'preco_club': preco_club,
        'data_cadastro': data_cadastro,
        'data_validade': data_validade,
        'destinatario': destinatario,
    }


def buscar_historico_produto(db, termo, exclude_id=None, limit=6):
    termo = (termo or '').strip()
    if len(termo) < 2:
        return []

    query = '''
        SELECT id, produto, unidade, preco_oferta, preco_club, data_cadastro, data_validade
        FROM ofertas
        WHERE produto LIKE ?
    '''
    params = [f'%{termo}%']
    if exclude_id:
        query += ' AND id != ?'
        params.append(exclude_id)
    query += ' ORDER BY data_cadastro DESC, id DESC LIMIT ?'
    params.append(limit)

    return db.execute(query, params).fetchall()

def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DATABASE)
        db.row_factory = sqlite3.Row
    return db

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

def load_layout_config(tipo='a4'):
    filename = 'layout_config_pequeno.json' if tipo == 'pequeno' else 'layout_config.json'
    try:
        with open(filename, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

def save_layout_config(config, tipo='a4'):
    filename = 'layout_config_pequeno.json' if tipo == 'pequeno' else 'layout_config.json'
    with open(filename, 'w') as f:
        json.dump(config, f)

def init_db():
    with app.app_context():
        db = get_db()
        db.execute('''
            CREATE TABLE IF NOT EXISTS ofertas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                produto TEXT NOT NULL,
                unidade TEXT NOT NULL,
                preco_oferta REAL NOT NULL,
                preco_club REAL,
                data_cadastro DATE DEFAULT CURRENT_DATE
            )
        ''')
        try:
            db.execute('ALTER TABLE ofertas ADD COLUMN layout_custom TEXT')
        except sqlite3.OperationalError:
            pass # Coluna já existe
        try:
            db.execute('ALTER TABLE ofertas ADD COLUMN layout_custom_pequeno TEXT')
        except sqlite3.OperationalError:
            pass # Coluna já existe
        try:
            db.execute('ALTER TABLE ofertas ADD COLUMN data_validade TEXT')
        except sqlite3.OperationalError:
            pass # Coluna já existe
        try:
            db.execute("ALTER TABLE ofertas ADD COLUMN destinatario TEXT NOT NULL DEFAULT 'todos'")
        except sqlite3.OperationalError:
            pass # Coluna já existe
        db.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user'
            )
        ''')
        try:
            db.execute('ALTER TABLE users ADD COLUMN loja TEXT')
        except sqlite3.OperationalError:
            pass # Coluna já existe
        # Inserir admin padrão se não houver usuários
        cursor = db.execute('SELECT COUNT(*) FROM users')
        if cursor.fetchone()[0] == 0:
            db.execute(
                'INSERT INTO users (username, password, role) VALUES (?, ?, ?)',
                ('admin', generate_password_hash('admin'), 'admin')
            )
        # Migração automática: se role for 'user' (filial) e loja estiver nula, e o username for um número de loja, salvar na coluna loja
        db.execute('''
            UPDATE users 
            SET loja = username 
            WHERE role = 'user' AND (loja IS NULL OR loja = '') AND username IN ('10', '20', '30', '40')
        ''')

        db.execute('''
            CREATE TABLE IF NOT EXISTS config (
                chave TEXT PRIMARY KEY,
                valor TEXT NOT NULL
            )
        ''')
        try:
            db.execute("INSERT OR IGNORE INTO config (chave, valor) VALUES ('club_label', 'ARAPONGAS PRIME')")
        except sqlite3.OperationalError:
            pass

        db.execute('''
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                action TEXT NOT NULL,
                entity_type TEXT NOT NULL,
                username TEXT NOT NULL,
                entity_id INTEGER,
                details TEXT,
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        try:
            db.execute('CREATE INDEX IF NOT EXISTS idx_audit_timestamp ON audit_log(timestamp DESC)')
        except sqlite3.OperationalError:
            pass

        usuarios = db.execute('SELECT id, password FROM users').fetchall()
        for usuario in usuarios:
            if not is_password_hash(usuario['password']):
                db.execute(
                    'UPDATE users SET password = ? WHERE id = ?',
                    (generate_password_hash(usuario['password']), usuario['id'])
                )

        db.commit()

# Rota para servir o Service Worker com os headers corretos para PWA
@app.route('/sw.js')
@app.route('/static/sw.js')
def serve_sw():
    response = send_from_directory('static', 'sw.js')
    response.headers['Content-Type'] = 'application/javascript'
    response.headers['Service-Worker-Allowed'] = '/'
    response.headers['Cache-Control'] = 'no-cache'
    return response

# Verifica se o banco existe antes do primeiro request, se não, cria.
with app.app_context():
    init_db()

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session or session.get('role') != 'admin':
            flash('Acesso negado. Apenas o Administrador pode acessar esta área.', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('admin' if session.get('role') == 'admin' else 'index'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        db = get_db()
        user = db.execute('SELECT * FROM users WHERE username = ?', (username,)).fetchone()
        
        if user and check_password_hash(user['password'], password):
            session.clear()
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['role'] = user['role']
            session['loja'] = user['loja']
            
            if user['role'] == 'admin':
                return redirect(url_for('admin'))
            else:
                return redirect(url_for('index'))
        else:
            flash('Usuário ou senha inválidos.', 'error')
            
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/admin/usuarios', methods=['GET', 'POST'])
@admin_required
def admin_usuarios():
    db = get_db()
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        role = request.form.get('role', 'user')
        loja = request.form.get('loja', '').strip()
        
        if role != 'admin' and not loja:
            flash('Por favor, selecione a loja do funcionário.', 'error')
            return redirect(url_for('admin_usuarios'))
            
        if not username or not password:
            flash('Usuário e senha são obrigatórios.', 'error')
            return redirect(url_for('admin_usuarios'))
        if len(password) < 4:
            flash('A senha deve ter pelo menos 4 caracteres.', 'error')
            return redirect(url_for('admin_usuarios'))
        try:
            db.execute(
                'INSERT INTO users (username, password, role, loja) VALUES (?, ?, ?, ?)',
                (username, generate_password_hash(password), role, loja if role != 'admin' else None)
            )
            db.commit()
            log_audit('create', 'usuario', None, f'Username: {username}, Role: {role}, Loja: {loja}')
            flash('Usuário cadastrado com sucesso!', 'success')
        except sqlite3.IntegrityError:
            flash('Este nome de usuário já existe.', 'error')
        return redirect(url_for('admin_usuarios'))
        
    usuarios = db.execute('SELECT * FROM users').fetchall()
    return render_template('usuarios.html', usuarios=usuarios, lojas_disponiveis=LOJAS_DISPONIVEIS)

@app.route('/admin/usuarios/delete/<int:id>', methods=['POST'])
@admin_required
def delete_usuario(id):
    if session.get('user_id') == id:
        flash('Você não pode excluir a si mesmo!', 'error')
        return redirect(url_for('admin_usuarios'))
        
    db = get_db()
    usuario = db.execute('SELECT username FROM users WHERE id = ?', (id,)).fetchone()
    if usuario:
        log_audit('delete', 'usuario', id, f'Username: {usuario["username"]}')
    db.execute('DELETE FROM users WHERE id = ?', (id,))
    db.commit()
    flash('Usuário excluído com sucesso!', 'success')
    return redirect(url_for('admin_usuarios'))

@app.route('/')
@login_required
def index():
    db = get_db()
    data_selecionada = parse_date_or_default(request.args.get('data'))
    busca = request.args.get('busca', '').strip()
    loja = session.get('loja', '')
    role = session.get('role', '')
    page = request.args.get('page', 1, type=int)
    per_page = 50
    ativas_apenas = request.args.get('ativas_apenas', '0') == '1'

    if ativas_apenas:
        date_condition = "data_cadastro <= ? AND (data_validade >= ? OR data_validade IS NULL OR data_validade = '')"
        date_params_list = [data_selecionada, data_selecionada]
    else:
        date_condition = "data_cadastro = ?"
        date_params_list = [data_selecionada]

    count_query = f'SELECT COUNT(*) FROM ofertas WHERE {date_condition}'
    params = list(date_params_list)
    if role != 'admin':
        count_query += " AND (destinatario = 'todos' OR ',' || destinatario || ',' LIKE ?)"
        params.append(f'%,{loja},%')
    if busca:
        count_query += ' AND produto LIKE ?'
        params.append(f'%{busca}%')

    total = db.execute(count_query, params).fetchone()[0]

    data_query = f'SELECT * FROM ofertas WHERE {date_condition}'
    data_params = list(date_params_list)
    if role != 'admin':
        data_query += " AND (destinatario = 'todos' OR ',' || destinatario || ',' LIKE ?)"
        data_params.append(f'%,{loja},%')
    if busca:
        data_query += ' AND produto LIKE ?'
        data_params.append(f'%{busca}%')
    data_query += ' ORDER BY id DESC LIMIT ? OFFSET ?'
    data_params.extend([per_page, (page - 1) * per_page])

    ofertas = db.execute(data_query, data_params).fetchall()

    total_pages = max(1, (total + per_page - 1) // per_page)

    data_obj = datetime.strptime(data_selecionada, '%Y-%m-%d')
    data_hoje_formatada = data_obj.strftime('%d/%m/%Y')
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('index.html', ofertas=ofertas, data_hoje=data_hoje_formatada,
                           data_selecionada=data_selecionada, page=page, busca=busca,
                           total_pages=total_pages, total=total, today=today,
                           ativas_apenas=ativas_apenas)

@app.route('/admin', methods=('GET', 'POST'))
@admin_required
def admin():
    db = get_db()
    if request.method == 'POST':
        produto = request.form.get('produto', '').strip()
        unidade = request.form.get('unidade', '').strip()
        preco_oferta = request.form.get('preco_oferta', '')
        preco_club = request.form.get('preco_club', '')
        data_cadastro = parse_date_or_default(request.form.get('data_cadastro'))
        data_validade = request.form.get('data_validade')
        dest_todos = request.form.get('dest_todos')
        if dest_todos:
            destinatario = 'todos'
        else:
            selecionados = request.form.getlist('destinatarios')
            destinatario = ','.join(selecionados) if selecionados else 'todos'
        if not produto or not preco_oferta:
            flash('Produto e Preço de Oferta são obrigatórios!', 'error')
        else:
            try:
                preco_oferta_val = parse_price(preco_oferta)
                preco_club_val = parse_price(preco_club)
                if preco_oferta_val is None or preco_oferta_val < 0:
                    raise ValueError
                if preco_club_val is not None and preco_club_val < 0:
                    raise ValueError
                
                novo_id = insert_oferta(
                    db, produto, unidade, preco_oferta_val, preco_club_val,
                    data_cadastro, data_validade, destinatario
                )
                db.commit()
                log_audit('create', 'oferta', novo_id, f'Produto: {produto}, Preço: {preco_oferta_val}')
                flash('Oferta cadastrada com sucesso!', 'success')
                return redirect(url_for('admin', data=data_cadastro))
            except ValueError:
                flash('Erro de formato no preço. Use números (ex: 10.50 ou 10,50).', 'error')

    data_selecionada = request.args.get('data', '')
    if data_selecionada:
        data_selecionada = parse_date_or_default(data_selecionada, default='')
    busca = request.args.get('busca', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 50

    count_query = "SELECT COUNT(*) FROM ofertas WHERE 1=1"
    params = []

    if data_selecionada:
        count_query += " AND data_cadastro = ?"
        params.append(data_selecionada)
    if busca:
        count_query += " AND produto LIKE ?"
        params.append(f"%{busca}%")

    total = db.execute(count_query, params).fetchone()[0]

    data_query = "SELECT * FROM ofertas WHERE 1=1"
    if data_selecionada:
        data_query += " AND data_cadastro = ?"
    if busca:
        data_query += " AND produto LIKE ?"
    data_query += " ORDER BY id DESC LIMIT ? OFFSET ?"
    data_params = params + [per_page, (page - 1) * per_page]

    ofertas = db.execute(data_query, data_params).fetchall()
    total_pages = max(1, (total + per_page - 1) // per_page)

    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('admin.html', ofertas=ofertas, data_selecionada=data_selecionada,
                           busca=busca, lojas_disponiveis=LOJAS_DISPONIVEIS,
                           page=page, total_pages=total_pages, total=total, today=today)

@app.route('/admin/delete/<int:id>', methods=('POST',))
@admin_required
def delete_oferta(id):
    db = get_db()
    oferta = db.execute('SELECT produto FROM ofertas WHERE id = ?', (id,)).fetchone()
    if oferta:
        log_audit('delete', 'oferta', id, f'Produto: {oferta["produto"]}')
    db.execute('DELETE FROM ofertas WHERE id = ?', (id,))
    db.commit()
    flash('Oferta excluída com sucesso!', 'success')
    return redirect(url_for('admin'))

@app.route('/admin/duplicar/<int:id>', methods=('POST',))
@admin_required
def duplicar_oferta(id):
    db = get_db()
    oferta = db.execute('SELECT * FROM ofertas WHERE id = ?', (id,)).fetchone()
    if not oferta:
        flash('Oferta não encontrada.', 'error')
        return redirect(url_for('admin'))

    data_cadastro = parse_date_or_default(request.form.get('data_cadastro'))
    novo_id = insert_oferta(
        db,
        oferta['produto'],
        oferta['unidade'],
        oferta['preco_oferta'],
        oferta['preco_club'],
        data_cadastro,
        oferta['data_validade'],
        oferta['destinatario'] or 'todos',
        oferta['layout_custom'] if 'layout_custom' in oferta.keys() else None,
        oferta['layout_custom_pequeno'] if 'layout_custom_pequeno' in oferta.keys() else None,
    )
    db.commit()
    flash(f'Oferta #{id} duplicada como #{novo_id}.', 'success')
    return redirect(url_for('editar_oferta', id=novo_id))

@app.route('/admin/importar', methods=('POST',))
@admin_required
def importar_ofertas():
    arquivo = request.files.get('arquivo')
    if not arquivo or not arquivo.filename:
        flash('Selecione um arquivo CSV ou Excel para importar.', 'error')
        return redirect(url_for('admin'))

    filename = secure_filename(arquivo.filename)
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_IMPORT_EXTENSIONS:
        flash('Formato inválido. Use arquivos .csv ou .xlsx.', 'error')
        return redirect(url_for('admin'))

    try:
        rows = read_xlsx_import(arquivo) if ext == '.xlsx' else read_csv_import(arquivo)
    except UnicodeDecodeError:
        flash('Não foi possível ler o CSV. Salve o arquivo em UTF-8 e tente novamente.', 'error')
        return redirect(url_for('admin'))
    except RuntimeError as exc:
        flash(str(exc), 'error')
        return redirect(url_for('admin'))
    except Exception as exc:
        flash(f'Erro ao ler o arquivo: {exc}', 'error')
        return redirect(url_for('admin'))

    if not rows:
        flash('O arquivo não possui linhas para importar.', 'error')
        return redirect(url_for('admin'))

    db = get_db()
    importadas = 0
    erros = []
    primeira_data = None
    for numero_linha, row in enumerate(rows, start=2):
        try:
            oferta = build_offer_from_import_row(row)
            insert_oferta(db, **oferta)
            importadas += 1
            primeira_data = primeira_data or oferta['data_cadastro']
        except (ValueError, TypeError) as exc:
            erros.append(f'linha {numero_linha}: {exc}')

    if importadas:
        db.commit()
    else:
        db.rollback()

    if importadas:
        mensagem = f'{importadas} oferta(s) importada(s) com sucesso.'
        if erros:
            mensagem += f' {len(erros)} linha(s) ignorada(s): ' + '; '.join(erros[:5])
        flash(mensagem, 'success' if not erros else 'error')
        return redirect(url_for('admin', data=primeira_data or ''))

    flash('Nenhuma oferta foi importada. ' + '; '.join(erros[:5]), 'error')
    return redirect(url_for('admin'))

@app.route('/admin/editar/<int:id>', methods=('GET', 'POST'))
@admin_required
def editar_oferta(id):
    db = get_db()
    if request.method == 'POST':
        produto = request.form.get('produto', '').strip()
        unidade = request.form.get('unidade', '').strip()
        preco_oferta = request.form.get('preco_oferta', '')
        preco_club = request.form.get('preco_club', '')
        data_cadastro = parse_date_or_default(request.form.get('data_cadastro'), default='')
        data_validade = request.form.get('data_validade')
        # Lê os checkboxes de filiais selecionadas
        dest_todos = request.form.get('dest_todos')
        if dest_todos:
            destinatario = 'todos'
        else:
            selecionados = request.form.getlist('destinatarios')
            destinatario = ','.join(selecionados) if selecionados else 'todos'
        
        if not produto or not preco_oferta or not data_cadastro:
            flash('Produto, Preço de Oferta e Data são obrigatórios!', 'error')
        else:
            try:
                preco_oferta_val = parse_price(preco_oferta)
                preco_club_val = parse_price(preco_club)
                if preco_oferta_val is None or preco_oferta_val < 0:
                    raise ValueError
                if preco_club_val is not None and preco_club_val < 0:
                    raise ValueError
                
                db.execute(
                    'UPDATE ofertas SET produto = ?, unidade = ?, preco_oferta = ?, preco_club = ?, data_cadastro = ?, data_validade = ?, destinatario = ? WHERE id = ?',
                    (produto, unidade, preco_oferta_val, preco_club_val, data_cadastro, data_validade, destinatario, id)
                )
                db.commit()
                log_audit('update', 'oferta', id, f'Produto: {produto}, Preço: {preco_oferta_val}')
                flash('Oferta atualizada com sucesso!', 'success')
                return redirect(url_for('admin', data=data_cadastro))
            except ValueError:
                flash('Erro de formato no preço. Use números (ex: 10.50 ou 10,50).', 'error')
                
    oferta = db.execute('SELECT * FROM ofertas WHERE id = ?', (id,)).fetchone()
    if not oferta:
        return "Oferta não encontrada", 404
    return render_template('editar_oferta.html', oferta=oferta, lojas_disponiveis=LOJAS_DISPONIVEIS)

@app.route('/admin/duplicar_lote', methods=['POST'])
@admin_required
def duplicar_ofertas_lote():
    ids = request.form.getlist('ids')
    data_destino = parse_date_or_default(request.form.get('data_destino'))
    if not ids or not data_destino:
        flash('Selecione ofertas e informe a data destino.', 'error')
        return redirect(url_for('admin'))

    db = get_db()
    duplicadas = 0
    for sid in ids:
        try:
            oferta = db.execute('SELECT * FROM ofertas WHERE id = ?', (int(sid),)).fetchone()
            if not oferta:
                continue
            insert_oferta(
                db, oferta['produto'], oferta['unidade'],
                oferta['preco_oferta'], oferta['preco_club'],
                data_destino, oferta['data_validade'],
                oferta['destinatario'] or 'todos'
            )
            duplicadas += 1
        except (ValueError, TypeError):
            continue

    if duplicadas:
        db.commit()
        flash(f'{duplicadas} oferta(s) duplicada(s) para {data_destino}.', 'success')
    else:
        flash('Nenhuma oferta foi duplicada.', 'error')
    return redirect(url_for('admin', data=data_destino))



@app.route('/admin/dashboard')
@admin_required
def admin_dashboard():
    db = get_db()
    hoje = datetime.now().strftime('%Y-%m-%d')

    total_ofertas = db.execute('SELECT COUNT(*) FROM ofertas').fetchone()[0]
    ofertas_hoje = db.execute('SELECT COUNT(*) FROM ofertas WHERE data_cadastro = ?', (hoje,)).fetchone()[0]
    ofertas_vencendo = db.execute("SELECT COUNT(*) FROM ofertas WHERE data_validade = ?", (hoje,)).fetchone()[0]
    ofertas_vencidas = db.execute("SELECT COUNT(*) FROM ofertas WHERE data_validade < ? AND data_validade != ''", (hoje,)).fetchone()[0]

    por_filial = db.execute('''
        SELECT destinatario, COUNT(*) as total FROM ofertas GROUP BY destinatario ORDER BY total DESC LIMIT 10
    ''').fetchall()

    ultimas = db.execute('SELECT * FROM ofertas ORDER BY id DESC LIMIT 5').fetchall()

    stats = {
        'total_ofertas': total_ofertas,
        'ofertas_hoje': ofertas_hoje,
        'ofertas_vencendo': ofertas_vencendo,
        'ofertas_vencidas': ofertas_vencidas,
    }
    return render_template('dashboard.html', stats=stats, por_filial=por_filial, ultimas=ultimas)


@app.route('/admin/usuarios/senha/<int:id>', methods=['POST'])
@admin_required
def admin_usuario_senha(id):
    db = get_db()
    nova_senha = request.form.get('nova_senha', '')
    if len(nova_senha) < 4:
        flash('A senha deve ter pelo menos 4 caracteres.', 'error')
        return redirect(url_for('admin_usuarios'))
    db.execute('UPDATE users SET password = ? WHERE id = ?', (generate_password_hash(nova_senha), id))
    db.commit()
    log_audit('update', 'usuario', id, 'Senha alterada pelo admin')
    flash('Senha alterada com sucesso!', 'success')
    return redirect(url_for('admin_usuarios'))


@app.route('/perfil', methods=['GET', 'POST'])
@login_required
def perfil():
    db = get_db()
    if request.method == 'POST':
        senha_atual = request.form.get('senha_atual', '')
        nova_senha = request.form.get('nova_senha', '')
        confirmar_senha = request.form.get('confirmar_senha', '')

        user = db.execute('SELECT * FROM users WHERE id = ?', (session['user_id'],)).fetchone()
        if not user or not check_password_hash(user['password'], senha_atual):
            flash('Senha atual incorreta.', 'error')
        elif nova_senha != confirmar_senha:
            flash('Nova senha e confirmação não conferem.', 'error')
        elif len(nova_senha) < 4:
            flash('A nova senha deve ter pelo menos 4 caracteres.', 'error')
        else:
            db.execute('UPDATE users SET password = ? WHERE id = ?', (generate_password_hash(nova_senha), session['user_id']))
            db.commit()
            flash('Senha alterada com sucesso!', 'success')
        return redirect(url_for('perfil'))

    return render_template('perfil.html')


@app.route('/admin/exportar')
@admin_required
def exportar_ofertas():
    db = get_db()
    data_selecionada = request.args.get('data', '')
    if data_selecionada:
        data_selecionada = parse_date_or_default(data_selecionada, default='')
    busca = request.args.get('busca', '').strip()

    query = "SELECT * FROM ofertas WHERE 1=1"
    params = []

    if data_selecionada:
        query += " AND data_cadastro = ?"
        params.append(data_selecionada)
    if busca:
        query += " AND produto LIKE ?"
        params.append(f"%{busca}%")
    query += " ORDER BY id DESC"

    ofertas = db.execute(query, params).fetchall()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['ID', 'Produto', 'Unidade', 'Preco_Oferta', 'Preco_Club', 'Data_Cadastro', 'Data_Validade', 'Destinatario'])
    for o in ofertas:
        writer.writerow([o['id'], o['produto'], o['unidade'], o['preco_oferta'],
                         o['preco_club'] or '', o['data_cadastro'], o['data_validade'] or '', o['destinatario'] or 'todos'])

    log_audit('export', 'oferta', None, f'Registros exportados: {len(ofertas)}')
    response = app.make_response(output.getvalue())
    response.headers['Content-Type'] = 'text/csv; charset=utf-8-sig'
    response.headers['Content-Disposition'] = 'attachment; filename=ofertas_exportadas.csv'
    return response


@app.route('/admin/auditoria')
@admin_required
def admin_auditoria():
    db = get_db()
    page = request.args.get('page', 1, type=int)
    per_page = 100

    total = db.execute('SELECT COUNT(*) FROM audit_log').fetchone()[0]
    logs = db.execute(
        'SELECT * FROM audit_log ORDER BY timestamp DESC LIMIT ? OFFSET ?',
        (per_page, (page - 1) * per_page)
    ).fetchall()
    total_pages = max(1, (total + per_page - 1) // per_page)

    return render_template('auditoria.html', logs=logs, page=page, total_pages=total_pages, total=total)


@app.route('/cartaz/<int:id>')
@login_required
def cartaz(id):
    tipo = request.args.get('tipo', 'a4')
    db = get_db()
    oferta = db.execute('SELECT * FROM ofertas WHERE id = ?', (id,)).fetchone()
    if oferta is None:
        return "Oferta não encontrada", 404
    if session.get('role') != 'admin' and not user_can_view_offer(oferta, session.get('loja'), session.get('role')):
        flash('Esta oferta não está disponível para sua loja.', 'error')
        return redirect(url_for('index'))
    
    col_name = 'layout_custom_pequeno' if tipo == 'pequeno' else 'layout_custom'
    if col_name in oferta.keys() and oferta[col_name]:
        layout = json.loads(oferta[col_name])
    else:
        layout = load_layout_config(tipo)
        
    club_label = get_config('club_label', 'ARAPONGAS PRIME')
    return render_template('cartaz.html', oferta=oferta, layout=layout, tipo=tipo, club_label=club_label)

@app.route('/cartazes_multi')
@login_required
def cartazes_multi():
    tipo = request.args.get('tipo', 'a4')
    ids = request.args.getlist('ids')
    if not ids:
        return "Nenhuma oferta selecionada para impressão", 400
    
    # Valida e converte para inteiros para evitar SQL Injection
    try:
        valid_ids = [int(i) for i in ids]
    except ValueError:
        return "IDs de oferta inválidos", 400
        
    placeholders = ','.join('?' for _ in valid_ids)
    db = get_db()
    ofertas = db.execute(f'SELECT * FROM ofertas WHERE id IN ({placeholders}) ORDER BY id DESC', valid_ids).fetchall()
    if session.get('role') != 'admin':
        ofertas = [o for o in ofertas if user_can_view_offer(o, session.get('loja'), session.get('role'))]
    if not ofertas:
        return "Nenhuma oferta disponível para impressão", 403
    
    itens = []
    col_name = 'layout_custom_pequeno' if tipo == 'pequeno' else 'layout_custom'
    layout_padrao = load_layout_config(tipo)
    
    for oferta in ofertas:
        if col_name in oferta.keys() and oferta[col_name]:
            layout = json.loads(oferta[col_name])
        else:
            layout = layout_padrao
        itens.append({'oferta': oferta, 'layout': layout})
        
    club_label = get_config('club_label', 'ARAPONGAS PRIME')
    return render_template('cartazes_multi.html', itens=itens, tipo=tipo, club_label=club_label)

@app.route('/admin/editor', defaults={'id': None})
@app.route('/admin/editor/<int:id>')
@admin_required
def editor(id):
    tipo = request.args.get('tipo', 'a4')
    db = get_db()
    if id is not None:
        oferta = db.execute('SELECT * FROM ofertas WHERE id = ?', (id,)).fetchone()
        if not oferta:
            return "Oferta não encontrada", 404
        
        col_name = 'layout_custom_pequeno' if tipo == 'pequeno' else 'layout_custom'
        if col_name in oferta.keys() and oferta[col_name]:
            layout = json.loads(oferta[col_name])
        else:
            layout = load_layout_config(tipo)
        
        oferta_data = oferta
        save_url = url_for('salvar_layout_id', id=id, tipo=tipo)
        is_custom = True
    else:
        layout = load_layout_config(tipo)
        oferta_data = {
            'produto': 'PRODUTO DE TESTE',
            'unidade': 'KG',
            'preco_oferta': 99.99,
            'preco_club': 89.99,
            'data_validade': datetime.now().strftime('%Y-%m-%d')
        }
        save_url = url_for('salvar_layout', tipo=tipo)
        is_custom = False

    club_label = get_config('club_label', 'ARAPONGAS PRIME')
    return render_template('editor.html', layout=layout, oferta=oferta_data, save_url=save_url, is_custom=is_custom, tipo=tipo, club_label=club_label)

@app.route('/admin/api/salvar_layout', methods=['POST'])
@admin_required
def salvar_layout():
    tipo = request.args.get('tipo', 'a4')
    novo_layout = request.json
    save_layout_config(novo_layout, tipo)
    return {'status': 'success'}

@app.route('/admin/api/salvar_layout/<int:id>', methods=['POST'])
@admin_required
def salvar_layout_id(id):
    tipo = request.args.get('tipo', 'a4')
    novo_layout = request.json
    db = get_db()
    col_name = 'layout_custom_pequeno' if tipo == 'pequeno' else 'layout_custom'
    if novo_layout is None:
        db.execute(f'UPDATE ofertas SET {col_name} = NULL WHERE id = ?', (id,))
    else:
        db.execute(f'UPDATE ofertas SET {col_name} = ? WHERE id = ?', (json.dumps(novo_layout), id))
    db.commit()
    return {'status': 'success'}

@app.route('/admin/api/historico_produto')
@admin_required
def historico_produto():
    termo = request.args.get('q', '').strip()
    exclude_id = request.args.get('exclude_id', type=int)
    db = get_db()
    historico = buscar_historico_produto(db, termo, exclude_id=exclude_id)
    return {
        'items': [
            {
                'id': item['id'],
                'produto': item['produto'],
                'unidade': item['unidade'],
                'preco_oferta': format_moeda(item['preco_oferta']),
                'preco_club': format_moeda(item['preco_club']) if item['preco_club'] is not None else '',
                'data_cadastro': item['data_cadastro'],
                'data_validade': item['data_validade'] or '',
            }
            for item in historico
        ]
    }

# Filtro customizado para formatar moeda no Jinja2
@app.template_filter('moeda')
def format_moeda(value):
    if value is None:
        return ""
    # Formata como R$ 0,00
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

if __name__ == '__main__':
    # Roda em 0.0.0.0 para ficar acessível para outras máquinas da LAN
    app.run(host='0.0.0.0', port=5000, debug=True)
