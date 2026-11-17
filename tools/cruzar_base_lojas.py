"""Uso: python cruzar_base_lojas.py <planilha.xlsx> <lojas_banco.json> <saida.json> [coordenadas.xlsx]"""
import json, re, sys, unicodedata, collections
from datetime import datetime
import openpyxl

ABA = 'LOJAS COM NINTENDO GERAL'
CORTE_NOME = 0.34
REDE_BASE = {'CASAS BAHIA': 'VIA VAREJO', "SAM'S CLUB": 'SAMS'}
# nome da rede como aparece no dashboard (pro filtro de rede pegar as lojas da base)
REDE_DASH = {'CASAS BAHIA': 'VIA VAREJO', "SAM'S CLUB": 'CARREFOUR'}
PARAR = set('DE DA DO DAS DOS E LOJA SHOPPING SHC SHOP CENTER LASA CARREFOUR HAVAN VIVO '
            'CASAS BAHIA SAMS SAM S CLUB HIPER RUA AVENIDA AV TIM'.split())


def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', str(s or '')) if unicodedata.category(c) != 'Mn').upper().strip()


def texto(v):
    s = str(v or '').strip()
    if 'Ã' in s:  # UTF-8 lido como latin-1 na origem da planilha ("GALVÃ£O")
        try:
            s = s.encode('latin-1').decode('utf-8')
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return s


def cnpj14(c):
    d = re.sub(r'\D', '', str(c or ''))
    return d.zfill(14) if d else ''


def palavras(s):
    return {t for t in re.split(r'[^A-Z0-9]+', sem_acento(s)) if len(t) > 1 and t not in PARAR}


def grupo_banco(l):
    txt = sem_acento(f"{l.get('rede')} {l.get('bandeira')} {l.get('nome')}")
    return 'SAMS' if 'SAM' in txt and 'CLUB' in txt else sem_acento(l.get('rede'))


def ler_base(caminho):
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb[ABA] if ABA in wb.sheetnames else next(
        w for w in wb.worksheets if {'CNPJ', 'LOJA'} <= {str(c).strip() for c in next(w.iter_rows(values_only=True)) if c})
    linhas = ws.iter_rows(values_only=True)
    cab = [str(c).strip() if c else '' for c in next(linhas)]
    col = {nome: i for i, nome in enumerate(cab)}
    pega = lambda r, nome: r[col[nome]] if nome in col and col[nome] < len(r) else None
    base = []
    for r in linhas:
        if not pega(r, 'LOJA'):
            continue
        base.append({
            'rede': texto(pega(r, 'REDE')),
            'loja': texto(pega(r, 'LOJA')),
            'cod': texto(pega(r, 'SIGLA/CÓD. LOJA')),
            'cnpj': cnpj14(pega(r, 'CNPJ')),
            'cidade': texto(pega(r, 'CIDADE')),
            'uf': texto(pega(r, 'UF')).upper(),
            'ns2': str(pega(r, 'LOJA COM NS2') or '').strip().upper() == 'SIM',
            'online': 'ONLINE' in sem_acento(pega(r, 'OBS')),
        })
    return base


# 1 linha da base <-> 1 loja do banco, na ordem de confianca: CNPJ (o Excel perde zeros a
# esquerda), rede+cidade+nome parecido (o codigo "L227" conta como palavra), unica da rede na cidade.
def cruzar(base, lojas):
    por_cnpj = collections.defaultdict(list)
    por_cidade = collections.defaultdict(list)
    for l in lojas:
        if cnpj14(l.get('cnpj')):
            por_cnpj[cnpj14(l['cnpj'])].append(l)
        por_cidade[(grupo_banco(l), sem_acento(l.get('cidade')))].append(l)
    base_por_cidade = collections.Counter(
        (REDE_BASE.get(b['rede'], sem_acento(b['rede'])), sem_acento(b['cidade'])) for b in base)

    pares = []  # (score, i_base, cod_banco, via)
    for i, b in enumerate(base):
        chave = (REDE_BASE.get(b['rede'], sem_acento(b['rede'])), sem_acento(b['cidade']))
        for l in por_cnpj.get(b['cnpj'], []) if b['cnpj'] else []:
            pares.append((2.0, i, l['cod'], 'cnpj'))
        cands = por_cidade.get(chave, [])
        pb = palavras(b['loja']) | ({sem_acento(b['cod'])} if len(b['cod']) > 1 else set())
        for l in cands:
            pl = palavras(l.get('nome'))
            if pb and pl:
                j = len(pb & pl) / len(pb | pl)
                if j >= CORTE_NOME:
                    pares.append((1.0 + j, i, l['cod'], 'nome'))
        if len(cands) == 1 and base_por_cidade[chave] == 1:
            pares.append((1.0, i, cands[0]['cod'], 'unica na cidade'))

    pares.sort(key=lambda p: -p[0])
    base_usada, banco_usado, casadas = set(), set(), []
    for score, i, cod, via in pares:
        if i in base_usada or cod in banco_usado:
            continue
        base_usada.add(i)
        banco_usado.add(cod)
        base[i]['codCadastro'] = str(cod)
        casadas.append({'cod': str(cod), 'via': via, 'ns2': base[i]['ns2'], 'online': base[i]['online']})
    sem_cadastro = [{'rede': REDE_DASH.get(b['rede'], b['rede']), 'bandeira': b['rede'],
                     **{k: b[k] for k in ('loja', 'cidade', 'uf', 'ns2', 'online')},
                     'lat': b.get('lat'), 'lng': b.get('lng'), 'aprox': b.get('aprox', False)}
                    for i, b in enumerate(base) if i not in base_usada]
    return casadas, sem_cadastro


def chave_loja(rede, loja, cidade):
    return tuple(re.sub(r'[^A-Z0-9]', '', sem_acento(texto(v))) for v in (rede, loja, cidade))


# planilha geocodificada (mesmas lojas da base + LATITUDE/LONGITUDE). Linha repetida na base
# (mesma rede/loja/cidade) pega as coordenadas na mesma ordem.
def aplicar_coordenadas(base, caminho):
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    linhas = wb.worksheets[0].iter_rows(values_only=True)
    cab = [str(c).strip() if c else '' for c in next(linhas)]
    col = {nome: i for i, nome in enumerate(cab)}
    coords = collections.defaultdict(list)
    for r in linhas:
        lat, lng = r[col['LATITUDE']], r[col['LONGITUDE']]
        if isinstance(lat, (int, float)) and isinstance(lng, (int, float)):
            # FONTE_COORDENADA=CIDADE: loja sem endereco, ponto no centro da cidade (o mapa mostra como aproximada)
            aprox = str(r[col['FONTE_COORDENADA']] or '').strip().upper() == 'CIDADE' if 'FONTE_COORDENADA' in col else False
            coords[chave_loja(r[col['REDE']], r[col['LOJA']], r[col['CIDADE']])].append((lat, lng, aprox))
        else:
            coords[chave_loja(r[col['REDE']], r[col['LOJA']], r[col['CIDADE']])].append(None)
    achadas = 0
    for b in base:
        fila = coords.get(chave_loja(b['rede'], b['loja'], b['cidade']))
        par = fila.pop(0) if fila else None
        if par:
            b['lat'], b['lng'], b['aprox'] = par
            achadas += 1
    return achadas


def main():
    caminho, lojas_json, saida = sys.argv[1:4]
    base = ler_base(caminho)
    com_coord = aplicar_coordenadas(base, sys.argv[4]) if len(sys.argv) > 4 else 0
    with open(lojas_json, encoding='utf-8-sig') as f:
        lojas = json.load(f)
    casadas, sem_cadastro = cruzar(base, lojas)
    vias = collections.Counter(c['via'] for c in casadas)
    with open(saida, 'w', encoding='utf-8') as f:
        # "base" completa (linha da planilha -> loja do cadastro) e o de-para de lojas que os
        # importadores de estoque do varejo usam (sigla Carrefour, numero Havan...)
        json.dump({'arquivo': caminho, 'geradoEm': datetime.now().strftime('%Y-%m-%d %H:%M'), 'linhasBase': len(base),
                   'casadas': casadas, 'semCadastro': sem_cadastro,
                   'base': [{'rede': b['rede'], 'loja': b['loja'], 'codBase': b['cod'], 'cidade': b['cidade'], 'uf': b['uf'],
                             'cod': b.get('codCadastro')} for b in base]}, f, ensure_ascii=False)
    print(f"base: {len(base)} linhas | casadas com o cadastro: {len(casadas)} {dict(vias)} | sem cadastro: {len(sem_cadastro)}"
          f" | com coordenada: {com_coord} ({sum(1 for s in sem_cadastro if s['lat'] is not None)} das sem cadastro)")


if __name__ == '__main__':
    main()
