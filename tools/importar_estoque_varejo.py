"""Uso: python importar_estoque_varejo.py  (le dados\\estoque_varejo\\*.xlsx e injeta DADOS_VENDAS no dashboard.html)"""
import json, os, re, sys, unicodedata, collections, warnings
from datetime import datetime, timedelta
import openpyxl

warnings.filterwarnings('ignore')
sys.stdout.reconfigure(encoding='utf-8')

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA = os.path.join(RAIZ, 'dados', 'estoque_varejo')
HTML = os.path.join(RAIZ, 'dashboard.html')
DE_PARA = os.path.join(RAIZ, 'dados', 'de_para_produtos_varejo.csv')
CADASTRO = os.path.join(os.environ.get('TEMP', ''), 'nintendo_lojas_cadastro.json')
HISTORICO = os.path.join(RAIZ, 'dados', 'estoque_varejo_historico.json')
CAMPOS_HIST = ('cod', 'rede', 'bandeira', 'nome', 'nomeVarejo', 'uf', 'cidade', 'fonte', 'dataRef', 'codVarejo', 'produtos')
CRUZAMENTO = os.path.join(RAIZ, 'dados', 'base_lojas_cruzamento.json')
DIAS_SEM_VENDA_PARADO = 30   # Carrefour/Sam's: produto com estoque e sem venda ha 30+ dias
COBERTURA_PARADO = 90        # Havan: estoque pra mais de 90 dias de venda
DIAS_VISITA_X_ARQUIVO = 21   # visita e arquivo de estoque ate 21 dias de distancia pra cruzar
CATEGORIA = {'CONSOLE': 'CONSOLE', 'GAMES': 'SOFTWARE', 'ACESSÓRIOS': 'ACESSÓRIO', 'ACESSORIOS': 'ACESSÓRIO'}


def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', str(s or '')) if unicodedata.category(c) != 'Mn').upper().strip()


def num(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def data_no_nome(caminho):
    # "..._28_9.xlsx" -> 28/09 do ano do arquivo (None se o nome nao traz data)
    mtime = datetime.fromtimestamp(os.path.getmtime(caminho))
    m = re.findall(r'(\d{1,2})[._](\d{1,2})(?=\D*$)', os.path.basename(caminho))
    if m:
        dia, mes = (int(x) for x in m[-1])
        try:
            return datetime(mtime.year, mes, dia).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return None


def data_ref(caminho):
    return data_no_nome(caminho) or datetime.fromtimestamp(os.path.getmtime(caminho)).strftime('%Y-%m-%d')


def fim_semana(semana=None, periodo=None, ano=None):
    """Ultimo dia da semana do arquivo: 'W40' (+ ano) ou '28/9 a 4/10' -> '2026-10-04'."""
    hoje = datetime.now()
    if periodo:
        m = re.search(r'(\d{1,2})/(\d{1,2})\s*$', str(periodo).strip())
        if m:
            a = int(ano) if ano else hoje.year
            try:
                d = datetime(a, int(m.group(2)), int(m.group(1)))
            except ValueError:
                d = None
            if d:
                if not ano and d > hoje + timedelta(days=10):
                    d = d.replace(year=a - 1)
                return d.strftime('%Y-%m-%d')
    if semana:
        m = re.search(r'(\d{1,2})', str(semana))
        if m:
            a = int(ano) if ano else hoje.year
            return datetime.strptime(f'{a}-W{int(m.group(1))}-7', '%G-W%V-%u').strftime('%Y-%m-%d')
    return None


def eh_nintendo_ean(ean):
    return re.sub(r'\D', '', str(ean or '')).lstrip('0').startswith('45496')


def visita_proxima(visita, arquivo):
    try:
        return abs((datetime.strptime(str(visita)[:10], '%Y-%m-%d') - datetime.strptime(arquivo, '%Y-%m-%d')).days) <= DIAS_VISITA_X_ARQUIVO
    except ValueError:
        return False


def cabecalho(linha):
    return {sem_acento(c): i for i, c in enumerate(linha) if c not in (None, '')}


def ler_carrefour_sams(caminho, aba=None):
    """Relatorio 'Estoque Loja a Loja': aba BASE (arquivo antigo, um por bandeira) ou abas CRF e SAMS
    (arquivo novo, as duas bandeiras juntas e com produtos de outras marcas: fica so o EAN Nintendo)."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    nome_aba = 'BASE' if 'BASE' in wb.sheetnames else aba
    if nome_aba not in wb.sheetnames:
        return []
    linhas = list(wb[nome_aba].iter_rows(values_only=True))
    c = cabecalho(linhas[0])
    tem_categoria = 'CATEGORIA' in c
    pega = lambda r, nome: r[c[nome]] if nome in c and c[nome] < len(r) else None
    col_atend = next((k for k in c if k.startswith('ATENDIMENTO')), None)
    itens = []
    for r in linhas[1:]:
        if not pega(r, 'PRODUTO'):
            continue
        nome = str(pega(r, 'NOME DA LOJA') or '').strip()
        atend = sem_acento(pega(r, col_atend)) if col_atend else ''
        ean = re.sub(r'\D', '', str(pega(r, 'EAN DO PRODUTO') or ''))
        if not tem_categoria and not eh_nintendo_ean(ean):
            continue
        produto = str(pega(r, 'PRODUTO')).strip()
        itens.append({
            'chave': sem_acento(pega(r, 'SIGLA DA LOJA')), 'codLoja': str(pega(r, 'COD DA LOJA') or '').strip(),
            'nomeLoja': nome, 'uf': sem_acento(pega(r, 'UF')),
            'cd': atend == 'CD' or sem_acento(nome).startswith('CD') or 'CENTRO DIST' in sem_acento(nome),
            'produto': produto, 'ean': ean,
            'categoria': CATEGORIA.get(sem_acento(pega(r, 'CATEGORIA')), sem_acento(pega(r, 'CATEGORIA'))) if tem_categoria else categoria_por_nome(produto),
            'estoque': num(pega(r, 'ESTOQUE DISPONIVEL')) or 0, 'vendaDia': None,
            'diasSemVenda': num(pega(r, 'DIAS SEM VENDA')), 'diasSemAbast': num(pega(r, 'DIAS SEM ABASTECIMENTO')),
            'dataRef': fim_semana(pega(r, 'SEMANA'), pega(r, 'PERIODO')),
        })
    return itens


def ler_havan(caminho):
    """Relatorio estoque x vendas: 2 linhas de cabecalho (metricas em cima, identificacao embaixo)."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    linhas = list(wb.worksheets[0].iter_rows(values_only=True))
    # linha com LOJA e ITEM = fim do cabecalho; as de cima trazem as metricas
    i_id = next(i for i, r in enumerate(linhas[:8]) if r and {'LOJA', 'ITEM'} <= set(cabecalho(r)))
    c = {}
    for r in linhas[:i_id + 1]:
        c.update(cabecalho(r))
    pega = lambda r, nome: r[c[nome]] if nome in c and c[nome] < len(r) else None
    col_vmd = next(k for k in c if k.startswith('VENDA MEDIA DIARIA'))
    itens = []
    for r in linhas[i_id + 1:]:
        if not r or not pega(r, 'LOJA') or not pega(r, 'ITEM'):
            continue
        loja = str(pega(r, 'LOJA')).strip()
        produto = str(pega(r, 'ITEM')).strip()
        # "Estoque atual" soma posicoes fora da loja (bate com PDV + deposito em ~85% das linhas);
        # o giro do proprio relatorio usa o estoque do ponto de venda, entao usamos PDV + deposito
        pdv, dep = num(pega(r, 'ESTOQUE PONTO DE VENDA')), num(pega(r, 'ESTOQUE DEPOSITO'))
        cd = bool(re.search(r'\bCDH?\b', sem_acento(loja)))
        estoque = (pdv or 0) + (dep or 0) if (pdv is not None or dep is not None) and not cd else (num(pega(r, 'ESTOQUE ATUAL')) or 0)
        itens.append({
            'chave': (re.match(r'\s*(\d+)', loja) or [None, ''])[1], 'codLoja': '', 'nomeLoja': loja,
            'uf': sem_acento(pega(r, 'UF')), 'cd': cd, 'produto': produto,
            'categoria': 'CONSOLE' if sem_acento(produto).startswith('CONSOLE') else categoria_por_nome(produto),
            'estoque': estoque, 'vendaDia': num(pega(r, col_vmd)) or 0, 'diasSemVenda': None, 'diasSemAbast': None,
            'dataRef': fim_semana(pega(r, 'SEMANA'), pega(r, 'PERIODO')),
        })
    return itens


# rede do arquivo -> (palavra no nome do arquivo, leitor, rede na base do roteiro, rede no dashboard, bandeira)
def categoria_por_nome(produto):
    p = sem_acento(produto)
    if p.startswith('CONSOLE') or re.match(r'(NOVO )?(BUNDLE )?NINTENDO SWITCH\b(?!.*(JOGO|GAME))', p) and not re.search(r'CONTROLE|CAMERA|CAMARA|JOY', p):
        return 'CONSOLE'
    if re.match(r'(CONTROLE|CAMERA|CAMARA|CARTAO|CARTO |MICROSD|ESTOJO|CASE|PELICULA|CARREGADOR|BASE|SUPORTE|FONE|HEADSET|JOY)', p):
        return 'ACESSÓRIO'
    return 'SOFTWARE'


def ler_livrarias(caminho):
    """Estoque (SALDO), venda (VENDA, periodo do relatorio) e preco (PV) por filial."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    linhas = list(wb.worksheets[0].iter_rows(values_only=True))
    c = cabecalho(linhas[0])
    pega = lambda r, nome: r[c[nome]] if nome in c and c[nome] < len(r) else None
    itens = []
    for r in linhas[1:]:
        if not r or not pega(r, 'FILIAL') or not pega(r, 'PRODUTO'):
            continue
        filial = str(pega(r, 'FILIAL')).strip()
        produto = str(pega(r, 'PRODUTO')).strip()
        cat = {'GAMES': 'SOFTWARE', 'ACESSORIO': 'ACESSÓRIO', 'CONSOLE': 'CONSOLE'}.get(sem_acento(pega(r, 'GAMES')))
        itens.append({'chave': '', 'codLoja': str(pega(r, 'DEPOSITO') or ''), 'nomeLoja': filial, 'uf': '',
                      'cd': sem_acento(filial) in ('DISTRIBUIDORA', 'CD'), 'produto': produto,
                      'categoria': cat or categoria_por_nome(produto), 'estoque': num(pega(r, 'SALDO')) or 0,
                      'vendaDia': None, 'vendaPeriodo': num(pega(r, 'VENDA')) or 0, 'preco': num(pega(r, 'PV')),
                      'diasSemVenda': None, 'diasSemAbast': None, 'dataRef': fim_semana(periodo=pega(r, 'SEMANA'))})
    return itens


def ler_vivo(caminho):
    """Aba 'ESTOQUE Wnn' (muda o numero toda semana): estoque livre por loja e material."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = next(w for w in wb.worksheets if sem_acento(w.title).startswith('ESTOQUE'))
    linhas = list(ws.iter_rows(values_only=True))
    c = cabecalho(linhas[0])
    pega = lambda r, nome: r[c[nome]] if nome in c and c[nome] < len(r) else None
    itens = []
    for r in linhas[1:]:
        loja = str(pega(r, 'LOJA') or '').strip() if r else ''
        if loja in ('', '-'):
            continue
        produto = str(pega(r, 'TEXTO BREVE MATERIAL') or '').strip()
        itens.append({'chave': str(pega(r, 'CENTRO') or ''), 'codLoja': str(pega(r, 'CENTRO') or ''), 'nomeLoja': loja,
                      'uf': sem_acento(pega(r, 'UF')), 'cd': sem_acento(pega(r, 'ATENDIMENTO')) == 'CD' or sem_acento(loja).startswith('CD '),
                      'produto': produto, 'categoria': categoria_por_nome(produto),
                      'estoque': num(pega(r, 'UTILIZACAO LIVRE')) or 0, 'vendaDia': None,
                      'diasSemVenda': None, 'diasSemAbast': None})
    return itens


def vivo_venda_nacional(caminho):
    """Aba RESUMO_VENDA: venda por semana do Brasil todo (a Vivo nao abre por loja)."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    if 'RESUMO_VENDA' not in wb.sheetnames:
        return None
    linhas = [r for r in wb['RESUMO_VENDA'].iter_rows(values_only=True) if any(r)]
    i_cab = next((i for i, r in enumerate(linhas) if r and sem_acento(r[0]) == 'EAN'), None)
    if i_cab is None:
        return None
    cab = [str(c).strip() if c else '' for c in linhas[i_cab]]
    semanas = [c for c in cab[2:] if re.match(r'\d+-\d{4}$', c)]
    produtos = []
    for r in linhas[i_cab + 1:]:
        if not r[1] or 'TOTAL' in sem_acento(r[0]):
            continue
        produtos.append({'produto': str(r[1]).strip(), 'semanas': {s: num(r[cab.index(s)]) or 0 for s in semanas}})
    return {'semanas': semanas, 'produtos': produtos,
            'total': sum(sum(p['semanas'].values()) for p in produtos)}


def ler_miranda(caminho):
    """Tabela produto x loja: linha 1 com o apelido da loja, linha 2 com 'LOJA n'."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    linhas = list(wb.worksheets[0].iter_rows(values_only=True))
    i_cab = next(i for i, r in enumerate(linhas) if r and any(sem_acento(c) == 'PRODUTO' for c in r))
    cab, apelidos = linhas[i_cab], linhas[i_cab - 1] if i_cab else [None] * len(linhas[i_cab])
    col_prod = next(i for i, c in enumerate(cab) if sem_acento(c) == 'PRODUTO')
    lojas = [(i, str(apelidos[i] or cab[i]).strip(), str(cab[i]).strip()) for i, c in enumerate(cab) if sem_acento(c).startswith('LOJA')]
    itens = []
    for r in linhas[i_cab + 1:]:
        if not r or not r[col_prod]:
            continue
        produto = str(r[col_prod]).strip()
        for i, apelido, cod in lojas:
            if i >= len(r) or r[i] is None:
                continue  # vazio = loja nao trabalha o produto; 0 = trabalha e esta zerado
            itens.append({'chave': cod, 'codLoja': cod, 'nomeLoja': f'{apelido.upper()} ({cod})', 'uf': '',
                          'cd': sem_acento(apelido) == 'CD', 'produto': produto, 'categoria': categoria_por_nome(produto),
                          'estoque': num(r[i]) or 0, 'vendaDia': None, 'diasSemVenda': None, 'diasSemAbast': None})
    return itens


def ler_lasa_estoque(caminho):
    """Estoque Americanas por loja ('L736 - CAPAO DA CANOA') x item. O arquivo nao traz data."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    linhas = list(wb.worksheets[0].iter_rows(values_only=True))
    c = cabecalho(linhas[0])
    pega = lambda r, nome: r[c[nome]] if nome in c and c[nome] < len(r) else None
    itens = []
    for r in linhas[1:]:
        if not r or not pega(r, 'LOJA') or not pega(r, 'ITEM'):
            continue
        loja = str(pega(r, 'LOJA')).strip()
        m = re.match(r'\s*(L?\d+)\s*-', loja)
        produto = str(pega(r, 'ITEM')).strip()
        itens.append({'chave': sem_acento(m.group(1)) if m else '', 'codLoja': sem_acento(m.group(1)) if m else '', 'nomeLoja': loja,
                      'uf': '', 'cd': False, 'produto': produto, 'categoria': categoria_por_nome(produto),
                      'estoque': num(pega(r, 'ESTOQUE LOJA')) or 0, 'vendaDia': None, 'diasSemVenda': None, 'diasSemAbast': None})
    return itens


def ler_casas_bahia(caminho):
    """Casas Bahia manda o total da rede (sem abrir por loja): venda e estoque por produto e semana,
    e a venda por canal (loja fisica x online)."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb['CONSOLIDADO'] if 'CONSOLIDADO' in wb.sheetnames else wb.worksheets[0]
    linhas = list(ws.iter_rows(values_only=True))
    c = cabecalho(linhas[0])
    pega = lambda r, nome: r[c[nome]] if nome in c and c[nome] < len(r) else None
    semanas = {}
    for r in linhas[1:]:
        if not r or not pega(r, 'SEMANA') or not pega(r, 'ANO'):
            continue
        n = re.search(r'\d+', str(pega(r, 'SEMANA')))
        chave = f"{int(pega(r, 'ANO'))}-S{int(n.group()):02d}"
        s = semanas.setdefault(chave, {'chave': chave, 'periodo': str(pega(r, 'PERIODO (DATA)') or ''), 'venda': 0, 'estoque': 0, 'produtos': []})
        produto = str(pega(r, 'DESCRICAO PRODUTO') or '').strip()
        venda, estoque = num(pega(r, 'VENDA')) or 0, num(pega(r, 'ESTOQUE')) or 0
        s['venda'] += venda; s['estoque'] += estoque
        s['produtos'].append({'produto': produto, 'categoria': categoria_por_nome(produto), 'venda': venda, 'estoque': estoque})
    canal = None
    if 'VENDA POR CANAL' in wb.sheetnames:
        lc = list(wb['VENDA POR CANAL'].iter_rows(values_only=True))
        cc = cabecalho(lc[0])
        fis = sum(num(r[cc['VENDA LF']]) or 0 for r in lc[1:] if r and 'VENDA LF' in cc)
        onl = sum(num(r[cc['VENDA ON']]) or 0 for r in lc[1:] if r and 'VENDA ON' in cc)
        canal = {'lojaFisica': fis, 'online': onl}
    ordem = sorted(semanas)
    ultima = semanas[ordem[-1]] if ordem else None
    return {'arquivo': os.path.basename(caminho), 'canal': canal,
            'semanas': [{k: v for k, v in semanas[ch].items() if k != 'produtos'} for ch in ordem],
            'ultimaSemana': ultima}


def ler_sellout_crf_sams(caminho):
    """Venda loja a loja (sell-out) de Carrefour e Sam's: uma linha por dia x loja x produto.
    Conta loja fisica e retirada na loja; fica de fora e-commerce puro e televendas."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb['BASE VALIDA'] if 'BASE VALIDA' in wb.sheetnames else wb.worksheets[0]
    linhas = list(ws.iter_rows(values_only=True))
    c = cabecalho(linhas[0])
    pega = lambda r, *nomes: next((r[c[n]] for n in nomes if n in c and c[n] < len(r)), None)
    dados = {}
    for r in linhas[1:]:
        if not r or pega(r, 'DATE') is None:
            continue
        marca = pega(r, 'PRODUCT_BRAND')
        if marca and sem_acento(marca) != 'NINTENDO':
            continue
        if sem_acento(pega(r, 'SALES_ORIGIN')) in ('ECOMMERCE', 'TELEVENDAS'):
            continue
        d = pega(r, 'DATE')
        d = d if isinstance(d, datetime) else datetime.strptime(str(d)[:10], '%Y-%m-%d')
        ano, sem, _ = d.isocalendar()
        rede = sem_acento(pega(r, 'REDE', 'COMPANY'))
        fonte = "Sam's Club" if 'SAM' in rede else 'Carrefour'
        cod = str(pega(r, 'STORE_CODE') or '').strip()
        loja = dados.setdefault((f'{ano}-S{sem:02d}', fonte), {}).setdefault(cod, {
            'cnpj': re.sub(r'\D', '', str(pega(r, 'CNPJ', 'STORE_CNPJ') or '')).zfill(14), 'nome': str(pega(r, 'STORE_NAME') or '').strip(),
            'uf': sem_acento(pega(r, 'UF', 'STORE_UF')), 'cidade': sem_acento(pega(r, 'CIDADE', 'STORE_CITY')), 'produtos': {}})
        ean = re.sub(r'\D', '', str(pega(r, 'CODIGO DE BARRAS EAN', 'PRODUCT_EAN_BARCODE') or '')).lstrip('0')
        prod = loja['produtos'].setdefault(ean, {'produto': str(pega(r, 'PRODUCT_NAME') or '').strip(), 'qtd': 0})
        prod['qtd'] += num(pega(r, 'QUANTIDADE TOTAL VENDIDA', 'QUANTITY_TOTAL_UNIT_SOLD')) or 0
    return {'arquivo': os.path.basename(caminho), 'dados': dados}


def ler_modulacao_lasa(caminho):
    """Modulacao Americanas: o mix oficial de cada loja (Id Loja x item)."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = next((w for w in wb.worksheets if sem_acento(w.title).startswith('MODULA')), wb.worksheets[0])
    linhas = list(ws.iter_rows(values_only=True))
    c = cabecalho(linhas[0])
    pega = lambda r, nome: r[c[nome]] if nome in c and c[nome] < len(r) else None
    col_mod = next((k for k in c if k.startswith('MODULACAO')), None)
    itens = []
    for r in linhas[1:]:
        if not r or not pega(r, 'ID LOJA') or not pega(r, 'ITEM'):
            continue
        if col_mod and sem_acento(pega(r, col_mod)) not in ('SIM', 'S', '1'):
            continue
        itens.append({'idLoja': sem_acento(pega(r, 'ID LOJA')), 'loja': str(pega(r, 'LOJA') or '').strip(),
                      'item': str(pega(r, 'ITEM')).strip(), 'cidade': sem_acento(pega(r, 'CIDADE')),
                      'cnpj': re.sub(r'\D', '', str(pega(r, 'CNPJ') or '')).zfill(14)})
    return itens


# rede do arquivo -> (o nome do arquivo e dessa rede?, leitor, rede na base do roteiro, rede no dashboard, bandeira)
FONTES = [
    ('Carrefour', lambda n: ('CARREFOUR' in n and 'ESTOQUE LOJA' in n) or ('CRF' in n and 'ESTOQUE' in n),
     lambda a: ler_carrefour_sams(a, 'CRF'), 'CARREFOUR', 'CARREFOUR', 'CARREFOUR'),
    ("Sam's Club", lambda n: ('SAMS' in n and 'ESTOQUE LOJA' in n) or ('CRF' in n and 'SAMS' in n and 'ESTOQUE' in n),
     lambda a: ler_carrefour_sams(a, 'SAMS'), "SAM'S CLUB", 'CARREFOUR', "SAM'S CLUB"),
    ('Havan', lambda n: 'HAVAN' in n, ler_havan, 'HAVAN', 'HAVAN', 'HAVAN'),
    ('Livrarias Curitiba', lambda n: 'CURITIBA' in n, ler_livrarias, 'LIVRARIAS CURITIBA', 'LIVRARIAS CURITIBA', 'LIVRARIAS CURITIBA'),
    ('Vivo', lambda n: 'VIVO' in n, ler_vivo, 'VIVO', 'VIVO', 'VIVO'),
    ('Miranda', lambda n: 'MIRANDA' in n, ler_miranda, 'MIRANDA', 'MIRANDA', 'MIRANDA'),
    ('Americanas', lambda n: 'LASA' in n and 'ESTOQUE' in n, ler_lasa_estoque, 'LOJAS AMERICANAS', 'LOJAS AMERICANAS', 'LOJAS AMERICANAS'),
]
eh_sellout = lambda n: 'VENDA' in n and ('SAMS' in n or 'CARREFOUR' in n) and 'ESTOQUE' not in n
eh_casas_bahia = lambda n: 'CASAS BAHIA' in n
eh_modulacao = lambda n: 'LASA' in n and 'ESTOQUE' not in n
# redes sem codigo em comum com a base do roteiro: casa direto no cadastro pelo nome da loja
CASA_POR_NOME = {'Livrarias Curitiba', 'Vivo', 'Miranda', 'Americanas'}
PARAR_REDE = {'Livrarias Curitiba': {'LIVRARIA', 'LIVRARIAS', 'CURITIBA'}, 'Vivo': {'VIVO', 'VVO'},
              'Miranda': {'MIRANDA', 'PISO', 'COMPUTACAO'},
              'Americanas': {'LASA', 'SHC', 'AMERICANAS', 'LOJAS', 'EXPRESS', 'SHOPP'}}


def palavras_sams(s):
    return {t for t in re.split(r'[^A-Z0-9]+', sem_acento(s)) if len(t) > 1 and t not in {'SAM', 'SAMS', 'CLUB', 'LOJA'}}


def casar_loja(fonte, item, base_rede):
    """Linha da base do roteiro correspondente a loja do arquivo do varejo (ou None)."""
    if fonte in CASA_POR_NOME:
        return None
    if fonte == 'Carrefour':
        return base_rede.get(item['chave'])
    if fonte == 'Havan':
        return base_rede.get(item['chave'])
    # Sam's nao tem sigla: mesma UF + nome da cidade/loja
    pn = palavras_sams(item['nomeLoja'])
    melhor, nota = None, 0
    for b in base_rede.values():
        if item['uf'] and b['uf'] and sem_acento(b['uf']) != item['uf']:
            continue
        pb = palavras_sams(b['loja']) | palavras_sams(b['cidade'])
        j = len(pn & pb) / len(pn | pb) if pn and pb else 0
        if j > nota:
            melhor, nota = b, j
    return melhor if nota >= 0.5 else None


PARAR = {'CRFO', 'CARREFOUR', 'HAVAN', 'SAM', 'SAMS', 'CLUB', 'LOJA', 'SHOPPING', 'SHOP', 'HIPER', 'DE', 'DA', 'DO', 'DOS', 'DAS', 'E'}


def palavras(s):
    return {t for t in re.split(r'[^A-Z0-9]+', sem_acento(s)) if len(t) > 1 and t not in PARAR and not t.isdigit()}


def do_grupo(fonte, l):
    txt = sem_acento(f"{l.get('rede')} {l.get('bandeira')} {l.get('nome')}")
    sams = 'SAM' in txt and 'CLUB' in txt
    if fonte == "Sam's Club":
        return sams
    if fonte == 'Carrefour':
        return 'CARREFOUR' in txt and not sams
    if fonte == 'Livrarias Curitiba':
        return 'LIVRARIA' in sem_acento(l.get('rede'))
    if fonte == 'Americanas':
        return 'AMERICANAS' in sem_acento(l.get('rede'))
    if fonte in ('Vivo', 'Miranda'):
        return sem_acento(l.get('rede')) == sem_acento(fonte)
    return 'HAVAN' in txt


def bate(a, b):
    return a == b or (len(a) >= 3 and len(b) >= 3 and (a.startswith(b) or b.startswith(a)))


UFS = {'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI',
       'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'}
# apelidos do arquivo que nao aparecem no cadastro (o cadastro traz o endereco)
ROMANOS = {'I', 'II', 'III', 'IV', 'V', 'VI'}
APELIDOS = {'Miranda': {'NORTE': 'POTENGI'}}


# casamento pelo nome da loja: todas as palavras do nome no arquivo (ou 2/3, se o nome e longo)
# precisam aparecer no nome/cidade da loja no cadastro, com pelo menos uma no nome; "MID" bate com "MIDWAY".
# Empate: fica a loja cujo nome no cadastro tem menos palavras sobrando, depois a atendida
def casar_por_nome(fonte, item, lojas_dash, usadas):
    extra = PARAR_REDE.get(fonte, set()) | UFS
    alvo = {APELIDOS.get(fonte, {}).get(a, a) for a in palavras(item['nomeLoja'].split('(')[0])
            if not re.match(r'^L\d+$', a)} - extra  # 'L736' e codigo da Americanas, nao nome
    if not alvo:
        return None
    melhor, nota = None, None
    for l in lojas_dash.values():
        if str(l['cod']) in usadas or not do_grupo(fonte, l):
            continue
        if item['uf'] in UFS and sem_acento(l.get('uf')) != item['uf']:
            continue
        pn = palavras(l.get('nome')) - extra
        pl = pn | (palavras(l.get('cidade')) - extra)
        achou = [a for a in alvo if any(bate(a, b) for b in pl)]
        no_nome = [b for b in pn if b not in ROMANOS and any(bate(a, b) for a in alvo)]
        if not no_nome:
            continue
        chave = (len(achou) / len(alvo), len(no_nome) / len(pn), bool(l.get('atendida')))
        if nota is None or chave > nota:
            melhor, nota = l, chave
    return melhor if nota and nota[0] >= (1.0 if len(alvo) <= 2 else 0.66) else None


# segundo passo quando a base do roteiro nao leva a uma loja do cadastro: mesma rede,
# mesma UF, e nome parecido ou mesma cidade (unica loja da rede na cidade)
def casar_cadastro(fonte, item, b, lojas_dash, usadas):
    cidade = sem_acento(b['cidade']) if b else ''
    uf = sem_acento(b['uf']) if b and b.get('uf') else item['uf']
    alvos = [p for p in (palavras(item['nomeLoja']), palavras(b['loja']) if b else set()) if p]
    jac = lambda a, c: len(a & c) / len(a | c) if a and c else 0
    cands = [l for l in lojas_dash.values() if str(l['cod']) not in usadas and do_grupo(fonte, l)
             and (not uf or sem_acento(l.get('uf')) == uf)]
    melhor, nota = None, 0
    for l in cands:
        pl = palavras(l.get('nome'))
        j = max((jac(a, pl) for a in alvos), default=0)
        if cidade and sem_acento(l.get('cidade')) == cidade:
            j += 0.3
        if j > nota:
            melhor, nota = l, j
    if nota >= 0.5:
        return melhor
    na_cidade = [l for l in cands if cidade and sem_acento(l.get('cidade')) == cidade]
    return na_cidade[0] if len(na_cidade) == 1 else None


def indexar_base(base, rede_base, fonte):
    idx = {}
    for b in base:
        if sem_acento(b['rede']) != sem_acento(rede_base):
            continue
        if fonte == 'Carrefour':
            chave = sem_acento(b['codBase'])
        elif fonte == 'Havan':
            chave = (re.match(r'\s*(\d+)', b['loja']) or [None, ''])[1]
        else:
            chave = sem_acento(b['loja'])
        if chave:
            idx[chave] = b
    return idx


def ler_de_para():
    # dados\de_para_produtos_varejo.csv (rede;produto_varejo;sku_campo): nome do produto no arquivo -> SKU da pesquisa de campo
    if not os.path.exists(DE_PARA):
        return {}
    de_para = {}
    for linha in open(DE_PARA, encoding='utf-8-sig').read().splitlines()[1:]:
        partes = linha.split(';')
        if len(partes) >= 3:
            de_para[(sem_acento(partes[0]), sem_acento(partes[1]))] = partes[2].strip()
    return de_para


codigo_lasa = lambda s: (lambda m: f"{m.group(1)}{int(m.group(2))}" if m else None)(re.match(r'(L?)\s*(\d+)$', str(s or '').strip()))


def mapa_lasa(lojas_dash):
    """'L227' / '1451' -> cod da loja no dashboard (codigo no nome do cadastro: 'LASA - L227 - SHC BANGU')."""
    por_codigo = {}
    for l in lojas_dash.values():
        m = re.match(r'LASA\s*-\s*(L?\s*\d+)\s*-', sem_acento(l.get('nome')))
        if m and 'AMERICANAS' in sem_acento(l.get('rede')):
            por_codigo[codigo_lasa(m.group(1))] = str(l['cod'])
    return por_codigo


def mapa_cnpj(lojas_dash):
    """CNPJ -> cod da loja no dashboard (cadastro exportado pelo atualizar_dashboard.ps1 antes deste script)."""
    por_cnpj = {}
    if os.path.exists(CADASTRO):
        for l in json.load(open(CADASTRO, encoding='utf-8-sig')):
            cnpj = re.sub(r'\D', '', str(l.get('cnpj') or ''))
            if cnpj and str(l['cod']) in lojas_dash:
                por_cnpj[cnpj.zfill(14)] = str(l['cod'])
    return por_cnpj


def modulacao_americanas(arquivos, lojas_dash):
    """Mix oficial de cada loja Americanas (arquivo 'Lasa'), ligado a loja do dashboard pelo codigo 'Lnnn'."""
    cands = [a for a in arquivos if eh_modulacao(sem_acento(os.path.basename(a)))]
    if not cands:
        return None
    arq = max(cands, key=lambda a: (data_ref(a), os.path.getmtime(a)))
    de_para = ler_de_para()
    # codigo "L227" ou "1451" (lojas novas vem sem o L), no nome do cadastro: "LASA - L227 - SHC BANGU"
    codigo = lambda s: (lambda m: f"{m.group(1)}{int(m.group(2))}" if m else None)(re.match(r'(L?)\s*(\d+)$', s.strip()))
    por_codigo = {}
    for l in lojas_dash.values():
        m = re.match(r'LASA\s*-\s*(L?\s*\d+)\s*-', sem_acento(l.get('nome')))
        if m and 'AMERICANAS' in sem_acento(l.get('rede')):
            por_codigo[codigo(m.group(1))] = str(l['cod'])
    # CNPJ do cadastro (exportado pelo atualizar_dashboard.ps1 antes deste script)
    por_cnpj = {}
    if os.path.exists(CADASTRO):
        for l in json.load(open(CADASTRO, encoding='utf-8-sig')):
            cnpj = re.sub(r'\D', '', str(l.get('cnpj') or ''))
            if cnpj and str(l['cod']) in lojas_dash:
                por_cnpj[cnpj.zfill(14)] = str(l['cod'])
    # sem CNPJ nem codigo no cadastro ("LASA - NITEROI"): casa pelo nome dentro da cidade
    # (a coluna UF do arquivo nao e confiavel: traz DF pra lojas do RJ)
    ligadas, por_id = set(por_codigo.values()) | set(por_cnpj.values()), {}
    lojas, sem_cadastro = {}, set()
    for it in ler_modulacao_lasa(arq):
        if it['idLoja'] not in por_id:
            cod = por_cnpj.get(it['cnpj']) or por_codigo.get(codigo(it['idLoja']))
            if not cod:
                nome = re.sub(r'^\s*L?\d+\s*-\s*', '', it['loja'])
                na_cidade = {k: l for k, l in lojas_dash.items() if sem_acento(l.get('cidade')) == it['cidade']}
                achada = casar_por_nome('Americanas', {'nomeLoja': nome, 'uf': ''}, na_cidade, ligadas)
                cod = str(achada['cod']) if achada else None
                if cod:
                    ligadas.add(cod)
            por_id[it['idLoja']] = cod
        cod = por_id[it['idLoja']]
        if not cod:
            sem_cadastro.add(it['idLoja'])
            continue
        sku = de_para.get(('AMERICANAS', sem_acento(it['item'])), '')
        mix = lojas.setdefault(cod, [])
        if all(i['item'] != it['item'] for i in mix):
            mix.append({'item': it['item'], 'sku': sku})
    print(f"  Americanas (modulacao): {os.path.basename(arq)} | {len(lojas)} lojas ligadas ao dashboard, "
          f"{len(sem_cadastro)} codigos sem loja no cadastro")
    return {'arquivo': os.path.basename(arq), 'dataRef': data_ref(arq), 'lojas': lojas}


def semana_iso(data):
    """'2026-09-28' -> ('2026-S40', segunda-feira da semana)"""
    d = datetime.strptime(data, '%Y-%m-%d')
    ano, sem, dia = d.isocalendar()
    return f'{ano}-S{sem:02d}', (d - timedelta(days=dia - 1)).strftime('%Y-%m-%d')


def atualizar_historico(lojas_out, fontes_out):
    """Guarda o que chegou de cada rede na semana (ISO) da data do arquivo. Os arquivos em dados\\estoque_varejo
    sao trocados toda semana; o historico (local, NDA) e o que permite o filtro de semana e a evolucao semanal."""
    hist = json.load(open(HISTORICO, encoding='utf-8')) if os.path.exists(HISTORICO) else {'semanas': {}, 'vivoNacional': {}}
    for f in fontes_out:
        chave, _ = semana_iso(f['dataRef'])
        hist['semanas'].setdefault(chave, {})[f['rede']] = {
            'dataRef': f['dataRef'], 'arquivo': f['arquivo'], 'estoqueCD': f['estoqueCD'],
            'lojas': [{k: v for k, v in l.items() if k in CAMPOS_HIST} for l in lojas_out if l['fonte'] == f['rede']]}
        vn = f.get('vendaNacional')
        if vn:  # semanas da Vivo vem como '37-2026'
            for p in vn['produtos']:
                for s, qtd in p['semanas'].items():
                    num_sem, ano = s.split('-')
                    hist['vivoNacional'].setdefault(f'{ano}-S{int(num_sem):02d}', {})[p['produto']] = qtd
    return hist


def salvar_historico(hist):
    # guarda no maximo 1 ano
    for chave in sorted(hist['semanas'])[:-52]:
        del hist['semanas'][chave]
    json.dump(hist, open(HISTORICO, 'w', encoding='utf-8'), ensure_ascii=False)


def aplicar_sellout(hist, so, lojas_dash, por_cnpj):
    """Distribui a venda do sell-out (Carrefour/Sam's) nas semanas do historico, loja a loja.
    Semana sem arquivo de estoque vira uma 'foto' so com venda (estoque sem dado)."""
    for (chave, fonte), lojas in sorted(so['dados'].items()):
        sem = hist['semanas'].setdefault(chave, {})
        snap = sem.get(fonte)
        if not snap:
            ano, n = chave.split('-S')
            fim = datetime.strptime(f'{ano}-W{int(n)}-7', '%G-W%V-%u').strftime('%Y-%m-%d')
            snap = sem[fonte] = {'dataRef': fim, 'arquivo': so['arquivo'], 'estoqueCD': 0, 'lojas': [], 'soVenda': True}
        snap['arquivoVenda'] = so['arquivo']
        por_cod = {str(l.get('codVarejo')): l for l in snap['lojas'] if l.get('codVarejo')}
        por_dash = {str(l['cod']): l for l in snap['lojas'] if l.get('cod')}
        # quem nao aparece no sell-out da semana vendeu zero
        for l in snap['lojas']:
            for p in l['produtos']:
                p['vendaPeriodo'] = 0
        for cod_v, info in lojas.items():
            cod_dash = por_cnpj.get(info['cnpj'])
            l = por_cod.get(cod_v) or (por_dash.get(cod_dash) if cod_dash else None)
            if not l:
                d = lojas_dash.get(cod_dash) if cod_dash else None
                l = {'cod': cod_dash or '', 'rede': d['rede'] if d else 'CARREFOUR',
                     'bandeira': d['bandeira'] if d else ("SAM'S CLUB" if fonte == "Sam's Club" else 'CARREFOUR'),
                     'nome': d['nome'] if d else info['nome'], 'nomeVarejo': info['nome'],
                     'uf': d['uf'] if d else info['uf'], 'cidade': d.get('cidade', '') if d else info['cidade'],
                     'fonte': fonte, 'dataRef': snap['dataRef'], 'codVarejo': cod_v, 'produtos': []}
                snap['lojas'].append(l)
                por_cod[cod_v] = l
            for ean, pi in info['produtos'].items():
                p = next((p for p in l['produtos'] if (p.get('ean') and p['ean'].lstrip('0') == ean)
                          or sem_acento(p['produto']) == sem_acento(pi['produto'])), None)
                if not p:
                    p = {'produto': pi['produto'], 'categoria': categoria_por_nome(pi['produto']), 'estoque': None,
                         'vendaDia': None, 'coberturaDias': None, 'diasSemVenda': None, 'diasSemAbast': None, 'parado': False, 'ean': ean}
                    l['produtos'].append(p)
                p['vendaPeriodo'] = (p.get('vendaPeriodo') or 0) + pi['qtd']
    return sorted({k[0] for k in so['dados']})


def compactar(hist, ns2_base):
    """Formato enxuto pro dashboard (o doc do Firestore tem limite de 1 MB):
    lojas (dimensao), produtos (dimensao), semanas e fatos = [semana, loja, produto, estoque, vendaDia, vendaPeriodo,
    diasSemVenda, diasSemAbast, preco]. Semana sem arquivo novo de uma rede repete o da semana anterior (1 semana no maximo),
    porque as redes mandam em dias diferentes (Sam's na quinta, Carrefour na segunda...)."""
    chaves = sorted(hist['semanas'])
    lojas, idx_loja, produtos, idx_prod, fatos, semanas = [], {}, [], {}, [], []
    for wi, chave in enumerate(chaves):
        visao, ant = dict(hist['semanas'][chave]), hist['semanas'].get(chaves[wi - 1]) if wi else None
        for fonte, snap in (ant or {}).items():
            visao.setdefault(fonte, snap)
        ano, num = chave.split('-S')
        segunda = datetime.strptime(f'{ano}-W{int(num)}-1', '%G-W%V-%u')
        semanas.append({'chave': chave, 'ano': int(ano), 'semana': int(num), 'inicio': segunda.strftime('%Y-%m-%d'),
                        'fim': (segunda + timedelta(days=6)).strftime('%Y-%m-%d'),
                        'fontes': {f: {'dataRef': s['dataRef'], 'arquivo': s['arquivo'], 'estoqueCD': s.get('estoqueCD', 0),
                                       'repetido': f not in hist['semanas'][chave], 'soVenda': bool(s.get('soVenda')),
                                       'arquivoVenda': s.get('arquivoVenda')} for f, s in visao.items()},
                        'vivoNacional': sum(hist.get('vivoNacional', {}).get(chave, {}).values()) or None})
        for fonte, snap in visao.items():
            for l in snap['lojas']:
                k = f"{fonte}|{l.get('cod') or l.get('codVarejo') or l['nomeVarejo']}"
                if k not in idx_loja:
                    idx_loja[k] = len(lojas)
                    lojas.append({c: l.get(c) for c in ('cod', 'rede', 'bandeira', 'nome', 'nomeVarejo', 'uf', 'cidade', 'fonte')})
                else:  # dados de cadastro mais recentes valem
                    lojas[idx_loja[k]].update({c: l.get(c) for c in ('cod', 'rede', 'bandeira', 'nome', 'nomeVarejo', 'uf', 'cidade') if l.get(c)})
                li = idx_loja[k]
                for p in l['produtos']:
                    kp = (p['produto'], p['categoria'])
                    if kp not in idx_prod:
                        idx_prod[kp] = len(produtos)
                        produtos.append({'nome': p['produto'], 'categoria': p['categoria']})
                    fatos.append([wi, li, idx_prod[kp], p['estoque'], p.get('vendaDia'), p.get('vendaPeriodo'),
                                  p.get('diasSemVenda'), p.get('diasSemAbast'), p.get('preco')])
    for l in lojas:
        l['ns2Base'] = ns2_base.get(str(l['cod'])) if l['cod'] else None
    vivo = {k: sum(v.values()) for k, v in sorted(hist.get('vivoNacional', {}).items())}
    return {'semanas': semanas, 'lojas': lojas, 'produtos': produtos, 'fatos': fatos, 'vivoNacional': vivo}


def main():
    if not os.path.isdir(PASTA):
        print('pasta dados\\estoque_varejo nao existe - nada a importar'); return 2
    html = open(HTML, encoding='utf-8').read()
    m = re.search(r'/\*__DADOS_START__\*/\s*const DADOS=(.*?);\s*/\*__DADOS_END__\*/', html, re.S)
    lojas_dash = {str(l['cod']): l for l in json.loads(m.group(1))['lojas']} if m else {}
    base = json.load(open(CRUZAMENTO, encoding='utf-8'))['base'] if os.path.exists(CRUZAMENTO) else []

    arquivos = [os.path.join(PASTA, f) for f in os.listdir(PASTA) if f.lower().endswith('.xlsx') and not f.startswith('~$')]
    nome_arq = lambda a: sem_acento(os.path.basename(a)).replace("'", '')
    por_cnpj, por_lasa = mapa_cnpj(lojas_dash), mapa_lasa(lojas_dash)
    # sell-out Carrefour/Sam's: alem da venda, liga o codigo de loja da rede ao CNPJ (e dai ao cadastro)
    sellout = None
    cands = [a for a in arquivos if eh_sellout(nome_arq(a))]
    if cands:
        sellout = ler_sellout_crf_sams(max(cands, key=os.path.getmtime))
    cod_rede_dash = {}
    for (_, fonte), lojas in (sellout or {'dados': {}})['dados'].items():
        for cod_v, info in lojas.items():
            if por_cnpj.get(info['cnpj']):
                cod_rede_dash[(fonte, cod_v)] = por_cnpj[info['cnpj']]
    fontes_out, lojas_out, sem_data = [], [], []
    for fonte, eh_da_rede, leitor, rede_base, rede_dash, bandeira in FONTES:
        candidatos = [a for a in arquivos if eh_da_rede(nome_arq(a))]
        if not candidatos:
            continue
        arq = max(candidatos, key=lambda a: (data_ref(a), os.path.getmtime(a)))
        itens = leitor(arq)
        if not itens:
            continue
        # data do arquivo: semana escrita dentro dele > data no nome > (sem data: semana dos outros arquivos)
        ref = max((i['dataRef'] for i in itens if i.get('dataRef')), default=None) or data_no_nome(arq)
        if not ref:
            sem_data.append(fonte)
            ref = data_ref(arq)
        base_rede = indexar_base(base, rede_base, fonte)
        por_loja = collections.OrderedDict()
        estoque_cd = 0
        for it in itens:
            if it['cd']:
                estoque_cd += it['estoque']
                continue
            por_loja.setdefault(it['nomeLoja'], []).append(it)

        casadas = 0
        # 1o passo pela base do roteiro (sigla/numero/nome); o que sobrar tenta direto no cadastro
        pares = []
        for nome_loja, its in por_loja.items():
            b = casar_loja(fonte, its[0], base_rede)
            # codigo da loja: Carrefour/Sam's pelo CNPJ do sell-out, Americanas pelo 'Lnnn' do cadastro
            cod = cod_rede_dash.get((fonte, its[0]['codLoja'])) if fonte in ('Carrefour', "Sam's Club") else \
                por_lasa.get(codigo_lasa(its[0]['chave'])) if fonte == 'Americanas' else None
            dash = lojas_dash.get(cod) if cod else (lojas_dash.get(str(b['cod'])) if b and b.get('cod') else None)
            pares.append([nome_loja, its, b, dash])
        usadas = {str(p[3]['cod']) for p in pares if p[3]}
        # nomes mais especificos primeiro: "IGUATEMI CAMPINAS" escolhe antes de "IGUATEMI"
        for p in sorted(pares, key=lambda p: -len(palavras(p[0]))):
            if not p[3]:
                if fonte in CASA_POR_NOME:
                    p[3] = casar_por_nome(fonte, p[1][0], lojas_dash, usadas)
                else:
                    p[3] = casar_cadastro(fonte, p[1][0], p[2], lojas_dash, usadas)
                if p[3]:
                    usadas.add(str(p[3]['cod']))
        for nome_loja, its, b, dash in pares:
            casadas += 1 if dash else 0
            tem_venda = any(i['vendaDia'] is not None for i in its)
            # Livrarias Curitiba manda a venda do periodo do relatorio, sem dizer quantos dias sao
            tem_periodo = any(i.get('vendaPeriodo') is not None for i in its)
            produtos = []
            for i in its:
                cob = round(i['estoque'] / i['vendaDia'], 0) if i['vendaDia'] else None
                parado = i['estoque'] > 0 and (
                    (i['diasSemVenda'] is not None and i['diasSemVenda'] >= DIAS_SEM_VENDA_PARADO)
                    or (tem_venda and (i['vendaDia'] or 0) == 0)
                    or (tem_periodo and (i.get('vendaPeriodo') or 0) == 0)
                    or (cob is not None and cob > COBERTURA_PARADO))
                prod = {'produto': i['produto'], 'categoria': i['categoria'], 'estoque': i['estoque'], 'ean': i.get('ean'),
                        'vendaDia': round(i['vendaDia'], 3) if i['vendaDia'] is not None else None,
                        'coberturaDias': cob, 'diasSemVenda': i['diasSemVenda'], 'diasSemAbast': i['diasSemAbast'],
                        'parado': parado}
                if tem_periodo:
                    prod['vendaPeriodo'] = i.get('vendaPeriodo')
                if i.get('preco') is not None:
                    prod['preco'] = i['preco']
                produtos.append(prod)
            produtos.sort(key=lambda p: (p['categoria'] != 'CONSOLE', -p['estoque']))
            estoque = sum(p['estoque'] for p in produtos)
            estoque_console = sum(p['estoque'] for p in produtos if p['categoria'] == 'CONSOLE')
            venda_dia = sum(p['vendaDia'] or 0 for p in produtos) if tem_venda else None
            abast = [p['diasSemAbast'] for p in produtos if p['diasSemAbast'] is not None]
            cons_falta = (dash or {}).get('consoleRupturas') or 0
            lojas_out.append({
                'cod': str(dash['cod']) if dash else '',
                'rede': dash['rede'] if dash else rede_dash,
                'bandeira': dash['bandeira'] if dash else bandeira,
                'nome': dash['nome'] if dash else (b['loja'] if b else nome_loja),
                'nomeVarejo': nome_loja, 'codVarejo': its[0].get('codLoja') or '',
                'uf': dash['uf'] if dash else (b['uf'] if b else its[0]['uf']),
                'cidade': dash.get('cidade', '') if dash else (b['cidade'] if b else ''),
                'fonte': fonte, 'dataRef': ref, 'temVenda': tem_venda,
                'estoque': estoque, 'estoqueConsole': estoque_console,
                'vendaDia': round(venda_dia, 2) if venda_dia is not None else None,
                'venda': round(venda_dia * 30, 1) if venda_dia is not None else None,
                'vendaPeriodo': sum(p.get('vendaPeriodo') or 0 for p in produtos) if tem_periodo else None,
                'giro': round(venda_dia * 30 / estoque, 3) if venda_dia is not None and estoque else None,
                'coberturaDias': round(estoque / venda_dia, 0) if venda_dia else None,
                'diasSemAbastMax': max(abast) if abast else None,
                'produtosParados': sum(1 for p in produtos if p['parado']),
                'excesso': any(p['parado'] for p in produtos),
                'wow': None,
                'atendida': bool(dash and dash.get('atendida')),
                'taxaRuptura': dash.get('taxaRuptura') if dash and dash.get('itensNoMix') else None,
                'consoleNoMix': (dash or {}).get('consoleNoMix') or 0,
                'consoleRupturas': cons_falta,
                'ultimaVisita': (dash or {}).get('ultimaVisita') or '',
                # console em falta na gondola (campo) com estoque de console no sistema do varejo;
                # so vale se a visita e o arquivo da rede forem de datas proximas
                'consoleSemExposicao': bool(dash and cons_falta > 0 and estoque_console > 0
                                            and visita_proxima((dash or {}).get('ultimaVisita'), ref)),
                'produtos': produtos,
            })
        info = {'rede': fonte, 'arquivo': os.path.basename(arq), 'dataRef': ref,
                'lojas': len(por_loja), 'lojasCasadas': casadas, 'estoqueCD': estoque_cd,
                'temVenda': any(i['vendaDia'] is not None for i in itens),
                'temVendaPeriodo': any(i.get('vendaPeriodo') is not None for i in itens)}
        if fonte == 'Vivo':
            info['vendaNacional'] = vivo_venda_nacional(arq)
        fontes_out.append(info)
        print(f"  {fonte}: {os.path.basename(arq)} ({ref}) | {len(por_loja)} lojas, {casadas} casadas com o dashboard"
              f" | estoque lojas {sum(i['estoque'] for i in itens if not i['cd']):.0f} un, CDs {estoque_cd:.0f} un")

    if not fontes_out:
        print('nenhum arquivo de estoque do varejo reconhecido em dados\\estoque_varejo'); return 2

    # arquivo sem data nenhuma (ex.: estoque Lasa): vale a semana mais recente dos outros arquivos
    if sem_data:
        refs = [f['dataRef'] for f in fontes_out if f['rede'] not in sem_data]
        if refs:
            for f in fontes_out:
                if f['rede'] in sem_data:
                    f['dataRef'] = max(refs)
            for l in lojas_out:
                if l['fonte'] in sem_data:
                    l['dataRef'] = max(refs)
            print(f"  sem data no arquivo: {', '.join(sem_data)} -> semana de {max(refs)}")
    hist = atualizar_historico(lojas_out, fontes_out)
    if sellout:
        sems = aplicar_sellout(hist, sellout, lojas_dash, por_cnpj)
        print(f"  venda loja a loja Carrefour/Sam's: {sellout['arquivo']} | semanas {sems[0]} a {sems[-1]} "
              f"| {sum(len(v) for v in sellout['dados'].values())} loja-semanas")
    salvar_historico(hist)
    ns2_base = {}
    if os.path.exists(CRUZAMENTO):
        ns2_base = {str(c['cod']): bool(c.get('ns2')) for c in json.load(open(CRUZAMENTO, encoding='utf-8'))['casadas']}
    dados = {'geradoEm': datetime.now().strftime('%Y-%m-%d %H:%M'), 'exemplo': False, 'fontes': fontes_out,
             'modulacao': modulacao_americanas(arquivos, lojas_dash), **compactar(hist, ns2_base)}
    cands = [a for a in arquivos if eh_casas_bahia(nome_arq(a))]
    if cands:
        dados['casasBahia'] = ler_casas_bahia(max(cands, key=os.path.getmtime))
        cb = dados['casasBahia']
        print(f"  Casas Bahia (rede toda): {cb['arquivo']} | {len(cb['semanas'])} semanas, de {cb['semanas'][0]['chave']} a {cb['semanas'][-1]['chave']}")
    print(f"  historico: {len(dados['semanas'])} semana(s) {[s['chave'] for s in dados['semanas']]} | {len(dados['fatos'])} linhas loja x produto")
    bloco = '/*__VENDAS_START__*/\nconst DADOS_VENDAS=' + json.dumps(dados, ensure_ascii=False, separators=(',', ':')) + ';\n/*__VENDAS_END__*/'
    novo, n = re.subn(r'/\*__VENDAS_START__\*/.*?/\*__VENDAS_END__\*/', lambda _: bloco, html, flags=re.S)
    if not n:
        raise SystemExit('marcadores /*__VENDAS_START__*/ nao encontrados no dashboard.html')
    open(HTML, 'w', encoding='utf-8', newline='').write(novo)
    sem = sum(1 for l in lojas_out if l['consoleSemExposicao'])
    print(f"OK -> dashboard.html ({len(bloco) // 1024} KB de vendas) | {len(lojas_out)} lojas com estoque do varejo | console sem exposicao: {sem} lojas")
    return 0


if __name__ == '__main__':
    sys.exit(main())
