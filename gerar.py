#!/usr/bin/env python3
"""Lê a aba de receitas da planilha online, calcula os indicadores e gera site/index.html.

O painel sai criptografado com a senha do secret PAINEL_SENHA. Se qualquer etapa falhar,
o script termina com erro antes de gravar o site, e o painel publicado continua o anterior.

Uso na automação:   python gerar.py
Teste local:        python gerar.py --xlsx copia.xlsx --hoje 2026-08-04
"""

import argparse
import base64
import calendar
import collections
import datetime as dt
import html
import json
import os
import re
import sys
import traceback
import unicodedata
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
BRT = dt.timezone(dt.timedelta(hours=-3))  # Brasília (sem horário de verão desde 2019)
EPOCA = dt.date(1899, 12, 30)  # dia zero das datas do Google Sheets e do Excel
MN = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
ITERACOES = 600_000
SAL = b'painel-receitas-v1'
PLANOS = ('TF', 'Pilates')
FISIO = ('Atendimento', 'FC')
NAO_TRATAMENTO = {'Consulta', 'Produtos', 'Avaliações', 'Outros'}
SEM_CORE = {'Consulta', 'Produtos', 'Avaliações'}
OBRIGATORIAS = ('empresa', 'cliente', 'pagamento', 'inicio', 'tratamento', 'fim', 'valor')
D30, D90 = dt.timedelta(days=30), dt.timedelta(days=90)


class ErroPainel(Exception):
    """Erro esperado. A mensagem vai para o log público, então nunca leva dados da planilha."""


# ---------------------------------------------------------------- utilitários

def norm(s):
    s = unicodedata.normalize('NFKD', str(s or ''))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', s).strip().lower()


def data(v):
    if v is None or v == '':
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return EPOCA + dt.timedelta(days=int(v)) if 1 < v < 100000 else None
    s = str(v).strip()
    for f in ('%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d'):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def numero(v):
    if v is None or v == '' or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r'[R$\s]', '', str(v))
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def mes(d):
    return f'{d.year:04d}-{d.month:02d}'


def ml(k):
    return f'{MN[int(k[5:7]) - 1]}/{k[2:4]}'


def fim_do_mes(k):
    y, m = int(k[:4]), int(k[5:7])
    return dt.date(y, m, calendar.monthrange(y, m)[1])


def meses_entre(a, b):
    y, m, out = int(a[:4]), int(a[5:7]), []
    while f'{y:04d}-{m:02d}' <= b:
        out.append(f'{y:04d}-{m:02d}')
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out


def ano_antes(d):
    try:
        return d.replace(year=d.year - 1)
    except ValueError:  # 29/02
        return d.replace(year=d.year - 1, day=28)


def nf(v, d=0):
    return f'{v:,.{d}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def brl(v):
    return 'R$ ' + nf(round(v))


def brlK(v):
    a, s = abs(v), '−' if v < 0 else ''
    if a >= 1e6:
        return f'{s}R$ {nf(a / 1e6, 2)} mi'
    if a >= 1e4:
        return f'{s}R$ {nf(a / 1e3, 1)} mil'
    return f'{s}R$ {nf(a)}'


def pct(v, d=0):
    return nf(v * 100, d) + '%'


def dpct(v, d=0):
    return ('+' if v > 0 else '−' if v < 0 else '') + nf(abs(v) * 100, d) + '%'


def var(a, b):
    return (b / a - 1) if a else None


def dm(d):
    return d.strftime('%d/%m/%y')


def dma(d):
    return d.strftime('%d/%m/%Y')


def rotulo_periodo(a, b):
    """'jan a set/26' quando cobre meses inteiros; senão '01/01 a 15/10/26'."""
    if a.day == 1 and b == fim_do_mes(mes(b)):
        if a.year == b.year:
            return ml(mes(b)) if a.month == b.month else f'{MN[a.month - 1]} a {ml(mes(b))}'
        return f'{ml(mes(a))} a {ml(mes(b))}'
    if a.year == b.year:
        return f'{a:%d/%m} a {b:%d/%m/%y}'
    return f'{a:%d/%m/%y} a {b:%d/%m/%y}'


def lista_pt(itens):
    itens = list(itens)
    return itens[0] if len(itens) == 1 else ', '.join(itens[:-1]) + ' e ' + itens[-1]


# ---------------------------------------------------------------- leitura

def id_planilha(ref):
    m = re.search(r'/d/([A-Za-z0-9_-]+)', ref)
    return m.group(1) if m else ref.strip()


def escolher_aba(nome, abas, titulo):
    for a in abas:
        if titulo(a) == nome:
            return a
    for a in abas:
        if norm(titulo(a)) == norm(nome):
            return a
    raise ErroPainel(f'aba "{nome}" não encontrada na planilha')


def ultima_coluna(intervalo):
    m = re.search(r':([A-Z]+)', intervalo.upper())
    n = 0
    for c in (m.group(1) if m else 'K'):
        n = n * 26 + ord(c) - 64
    return n


def ler_google(cfg):
    import gspread
    from gspread.utils import DateTimeOption, ValueRenderOption

    bruto = os.environ.get('GOOGLE_SA_JSON', '').strip()
    if not bruto:
        raise ErroPainel('o secret GOOGLE_SA_JSON não está cadastrado no GitHub')
    try:
        cred = json.loads(bruto)
    except json.JSONDecodeError:
        raise ErroPainel('o secret GOOGLE_SA_JSON não é um JSON válido (cole o conteúdo inteiro do arquivo .json da chave)') from None
    try:
        gc = gspread.service_account_from_dict(cred)
        planilha = gc.open_by_key(id_planilha(cfg['planilha']))
        abas = planilha.worksheets()
    except Exception as e:  # noqa: BLE001 (a mensagem original não vai para o log)
        cod = getattr(getattr(e, 'response', None), 'status_code', None)
        texto = str(e)
        if 'SERVICE_DISABLED' in texto or 'has not been used' in texto:
            raise ErroPainel('a Google Sheets API não está ativada no projeto do Google Cloud') from None
        if cod in (403, 404) or isinstance(e, (PermissionError, gspread.exceptions.SpreadsheetNotFound)):
            raise ErroPainel(f'sem acesso à planilha: compartilhe-a como Leitor com {cred.get("client_email", "o e-mail da conta de serviço")}') from None
        raise ErroPainel(f'falha ao abrir a planilha ({type(e).__name__}, código {cod})') from None
    aba = escolher_aba(cfg['aba'], abas, lambda w: w.title)
    linhas = aba.get(cfg.get('intervalo', 'A1:K'),
                     value_render_option=ValueRenderOption.unformatted,
                     date_time_render_option=DateTimeOption.serial_number)
    return [list(r) for r in linhas]


def ler_xlsx(caminho, cfg):
    import openpyxl
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    aba = escolher_aba(cfg['aba'], wb.worksheets, lambda w: w.title)
    return [list(r) for r in aba.iter_rows(max_col=ultima_coluna(cfg.get('intervalo', 'A1:K')), values_only=True)]


# ---------------------------------------------------------------- limpeza

class Lanc:
    __slots__ = ('emp', 'nome', 'cli', 'pg', 'ini', 'fim', 'trat', 'tipo', 'val', 'dur', 'forma', 'col1')


def agrupador(cfg):
    regras = [(norm(p), g['tipo']) for g in cfg['grupos'] for p in g['comeca_com']]
    novos = {}

    def tipo(trat):
        n = norm(trat)
        if not n:
            return 'Outros'
        for p, g in regras:
            if n.startswith(p):
                return g
        # sem < > & ": o nome vai para textos HTML do painel
        novos.setdefault(n, re.sub(r'[<>&"]', '', str(trat)).strip() or 'Outros')
        return novos[n]
    return tipo, novos


def montar(linhas, cfg, hoje):
    if not linhas:
        raise ErroPainel('a aba está vazia')
    cab = [norm(c) for c in linhas[0]]
    idx = {}
    for chave, nome in cfg['colunas'].items():
        if norm(nome) in cab:
            idx[chave] = cab.index(norm(nome))
        elif chave in OBRIGATORIAS:
            raise ErroPainel(f'a coluna "{nome}" sumiu da aba (confira o cabeçalho na linha 1)')

    def cel(r, chave):
        i = idx.get(chave)
        return r[i] if i is not None and i < len(r) else None

    tipo, novos = agrupador(cfg)
    validos, futuros, sem_valor = [], 0, 0
    for r in linhas[1:]:
        pg, val = data(cel(r, 'pagamento')), numero(cel(r, 'valor'))
        if pg is None or val is None:
            if pg is not None or (cel(r, 'cliente') not in (None, '')):
                sem_valor += 1
            continue
        if pg > hoje:
            futuros += 1
            continue
        x = Lanc()
        x.emp = str(cel(r, 'empresa') or '').strip() or 'Sem empresa'
        x.nome = re.sub(r'\s+', ' ', str(cel(r, 'cliente') or '')).strip() or None
        x.cli = norm(x.nome).replace(' ', '') if x.nome else None
        x.pg, x.ini, x.fim = pg, data(cel(r, 'inicio')), data(cel(r, 'fim'))
        x.trat = str(cel(r, 'tratamento') or '').strip()
        x.tipo = tipo(x.trat)
        x.val = val
        x.dur = numero(cel(r, 'duracao'))
        x.forma = str(cel(r, 'forma') or '').strip()
        x.col1 = numero(cel(r, 'mes'))
        validos.append(x)

    minimo = int(cfg.get('linhas_minimas', 1))
    if len(validos) < minimo:
        raise ErroPainel(f'a aba tem só {len(validos)} lançamentos válidos (mínimo esperado: {minimo})')
    if sum(x.val for x in validos) <= 0:
        raise ErroPainel('a soma da coluna Valor é zero')
    validos.sort(key=lambda x: x.pg)
    return validos, {'futuros': futuros, 'semValor': sem_valor, 'tiposNovos': sorted(novos.values())}


# ---------------------------------------------------------------- cálculo

def dur_plano(x):
    if x.fim is None:
        return None
    return (x.fim - (x.ini or x.pg)).days


def eh_plano(x):
    return x.tipo in PLANOS and x.cli and x.fim is not None and (dur_plano(x) or 0) >= 7


class Ctx:
    """Tudo que é do grupo inteiro e serve para os recortes por empresa."""

    def __init__(self, lanc, cfg, hoje):
        self.lanc = lanc
        self.hoje = hoje
        self.ultimo = lanc[-1].pg
        self.M = meses_entre(cfg.get('inicio', mes(lanc[0].pg)), mes(self.ultimo))
        self.mi = {k: i for i, k in enumerate(self.M)}
        par = mes(self.ultimo) == mes(hoje) and self.ultimo != fim_do_mes(mes(self.ultimo))
        self.parcial = {'i': len(self.M) - 1, 'dia': self.ultimo.day} if par else None
        # renovação e conversão: só o que já teve 30 dias para acontecer, em meses inteiros
        ok = [k for k in self.M if fim_do_mes(k) + D30 <= self.ultimo]
        self.corte_mes = ok[-1] if ok else None
        self.corte = fim_do_mes(self.corte_mes) if ok else None
        self.corte90 = self.ultimo - D90
        self.compras = collections.defaultdict(list)
        for x in lanc:
            if x.cli:
                self.compras[x.cli].append(x)
        self.nome = {}
        for c, xs in self.compras.items():
            self.nome[c] = collections.Counter(x.nome for x in xs).most_common(1)[0][0]

    def renovou(self, p):
        lim = p.fim + D30
        return any(y.tipo == p.tipo and p.pg < y.pg <= lim for y in self.compras[p.cli])

    def converteu(self, c):
        lim = c.pg + D30
        return any(y.tipo not in NAO_TRATAMENTO and c.pg <= y.pg <= lim for y in self.compras[c.cli])

    def trat30(self, c):
        lim = c.pg + D30
        return sum(y.val for y in self.compras[c.cli] if y.tipo not in NAO_TRATAMENTO and c.pg <= y.pg <= lim)


def r2(v):
    return round(v, 2)


def taxa_serie(ctx, itens, ok, mes_de):
    n, s = [0] * len(ctx.M), [0] * len(ctx.M)
    for x in itens:
        i = ctx.mi.get(mes_de(x))
        if i is None:
            continue
        n[i] += 1
        s[i] += ok(x)
    return {'taxa': [round(s[i] / n[i], 4) if n[i] >= 5 else None for i in range(len(n))], 'n': n}, sum(s), sum(n)


def bloco(ctx, L, tipos):
    M, mi, n = ctx.M, ctx.mi, len(ctx.M)
    rev = {t: [0.0] * n for t in tipos}
    qtd = {t: [0] * n for t in tipos}
    revAte = {t: [0.0] * n for t in tipos}
    qtdAte = {t: [0] * n for t in tipos}
    ativos = [set() for _ in M]
    sem_nome = [0] * n
    dia = ctx.parcial['dia'] if ctx.parcial else 31
    for x in L:
        i = mi.get(mes(x.pg))
        if i is None:
            continue
        rev[x.tipo][i] += x.val
        qtd[x.tipo][i] += 1
        if x.pg.day <= dia:
            revAte[x.tipo][i] += x.val
            qtdAte[x.tipo][i] += 1
        if x.cli:
            ativos[i].add(x.cli)
        else:
            sem_nome[i] += 1

    # "novo" = primeira compra do cliente neste recorte (no grupo, em "Todas"; na empresa, nas demais)
    primeira = {}
    for x in L:
        if x.cli:
            primeira.setdefault(x.cli, x.pg)
    novos = [0] * n
    for d in primeira.values():
        i = mi.get(mes(d))
        if i is not None:
            novos[i] += 1

    # uma compra = cliente + data: vários lançamentos do mesmo plano no mesmo dia contam uma vez
    planos = {}
    for x in L:
        if eh_plano(x) and ctx.corte and x.fim <= ctx.corte:
            k = (x.cli, x.tipo, x.pg)
            if k not in planos or x.fim > planos[k].fim:
                planos[k] = x
    planos = list(planos.values())
    renov, kpi_ren = {}, {}
    for t in PLANOS:
        serie, s, tot = taxa_serie(ctx, [p for p in planos if p.tipo == t], ctx.renovou, lambda p: mes(p.fim))
        renov[t] = serie
        kpi_ren[t] = (round(s / tot, 4) if tot else None, tot)
    consultas = {}
    for x in L:
        if x.tipo == 'Consulta' and x.cli and ctx.corte and x.pg <= ctx.corte:
            consultas.setdefault((x.cli, x.pg), x)
    conv, conv_s, conv_n = taxa_serie(ctx, consultas.values(), ctx.converteu, lambda c: mes(c.pg))

    carteira = []
    vig = [x for x in L if eh_plano(x)]
    for k in M:
        d = dt.date(int(k[:4]), int(k[5:7]), 15)
        d = min(d, ctx.ultimo)
        carteira.append(len({x.cli for x in vig if (x.ini or x.pg) <= d <= x.fim}))

    # clientes, recompra
    datas = collections.defaultdict(set)
    gasto = collections.defaultdict(float)
    tipos_cli = collections.defaultdict(set)
    for x in L:
        if x.cli:
            datas[x.cli].add(x.pg)
            gasto[x.cli] += x.val
            tipos_cli[x.cli].add(x.tipo)
    nc = len(datas)
    rec2n = sum(1 for d in datas.values() if len(d) >= 2)
    base90 = [sorted(d) for d in datas.values() if min(d) <= ctx.corte90]
    volta90 = sum(1 for d in base90 if any(d[0] < y <= d[0] + D90 for y in d[1:]))

    # venda cruzada
    dist = {k: {'cl': 0, 'gasto': 0.0, 'rev': 0.0} for k in ('1', '2', '3', '4')}
    for c, ts in tipos_cli.items():
        k = str(min(len(ts), 4))
        dist[k]['cl'] += 1
        dist[k]['rev'] += gasto[c]
    for d in dist.values():
        d['gasto'] = round(d['rev'] / d['cl']) if d['cl'] else 0
        d['rev'] = round(d['rev'], 2)
    multi = sum(1 for ts in tipos_cli.values() if len(ts) >= 2)
    multi_core = sum(1 for ts in tipos_cli.values() if len(ts - SEM_CORE) >= 2)
    rev_multi = sum(gasto[c] for c, ts in tipos_cli.items() if len(ts) >= 2)
    rev_cli = sum(gasto.values())

    cont = collections.Counter(t for ts in tipos_cli.values() for t in ts)
    mt = [t for t, _ in cont.most_common() if t not in ctx.matriz_excluir]
    m = [[0] * len(mt) for _ in mt]
    pos = {t: i for i, t in enumerate(mt)}
    for ts in tipos_cli.values():
        js = [pos[t] for t in ts if t in pos]
        for a in js:
            for b in js:
                m[a][b] += 1

    fisio = {c for c, ts in tipos_cli.items() if ts & set(FISIO)}
    mov = sum(1 for c in fisio if any(y.tipo in PLANOS for y in ctx.compras[c]))

    return {
        'rev': {t: [r2(v) for v in rev[t]] for t in tipos},
        'qtd': qtd,
        'revAte': {t: [r2(v) for v in revAte[t]] for t in tipos} if ctx.parcial else None,
        'qtdAte': qtdAte if ctx.parcial else None,
        'ativos': [len(a) for a in ativos],
        'novos': novos,
        'semNome': sem_nome,
        'renov': renov,
        'conv': conv,
        'carteira': carteira,
        'kpi': {
            'clientes': nc,
            'rec2': round(rec2n / nc, 4) if nc else None, 'rec2n': rec2n,
            'rec90': round(volta90 / len(base90), 4) if base90 else None, 'n90': len(base90),
            'renPilates': kpi_ren['Pilates'][0], 'renPilatesN': kpi_ren['Pilates'][1],
            'renTF': kpi_ren['TF'][0], 'renTFN': kpi_ren['TF'][1],
            'conv': round(conv_s / conv_n, 4) if conv_n else None, 'convN': conv_n,
        },
        'dist': dist,
        'cross': {
            'multi': round(multi / nc, 4) if nc else 0,
            'multiCore': round(multi_core / nc, 4) if nc else 0,
            'revMulti': round(rev_multi / rev_cli, 4) if rev_cli else 0,
        },
        'matrix': {'tipos': mt, 'm': m},
        'fisioMov': {'base': len(fisio), 'mov': mov},
    }


def oportunidades(ctx):
    R = ctx.ultimo
    linha = lambda x, v: {'c': ctx.nome[x.cli], 'e': x.emp, 't': x.tipo, 'pg': dm(x.pg),
                          'fim': dm(x.fim) if x.fim else '', 'v': r2(v), '_o': (x.fim or x.pg)}

    def pacote(x):
        return sum(y.val for y in ctx.compras[x.cli] if y.pg == x.pg and y.tipo == x.tipo)

    venc, jan, vindo, fis, cons = [], [], [], [], []
    for c, xs in ctx.compras.items():
        planos = [x for x in xs if eh_plano(x)]
        if planos:
            u = max(planos, key=lambda x: (x.fim, x.pg))
            if not any(y.tipo in PLANOS and y.pg > u.pg for y in xs):
                if R - D90 <= u.fim <= R - dt.timedelta(days=31):
                    venc.append(linha(u, pacote(u)))
                elif R - D30 <= u.fim < R:
                    jan.append(linha(u, pacote(u)))
                elif R <= u.fim <= R + D30:
                    vindo.append(linha(u, pacote(u)))
        if not any(x.tipo in PLANOS for x in xs):
            # pacote = Atendimento ou FC com data de fim e 7 dias ou mais (sessão avulsa não entra)
            pac = [x for x in xs if x.tipo in FISIO and x.fim and (x.fim - x.pg).days >= 7]
            if pac:
                u = max(pac, key=lambda x: (x.fim, x.pg))
                if R - D90 <= u.fim <= R and not any(y.pg > u.pg for y in xs):
                    fis.append(linha(u, pacote(u)))
        cs = [x for x in xs if x.tipo == 'Consulta' and R - D90 <= x.pg <= R - dt.timedelta(days=31)]
        if cs:
            u = cs[-1]
            if not any(y.tipo not in NAO_TRATAMENTO and y.pg >= u.pg for y in xs):
                cons.append(linha(u, u.val))

    out = {}
    for k, rows in (('vencidos', venc), ('janela', jan), ('vencendo', vindo), ('fisioSemMov', fis), ('consSemConv', cons)):
        rows.sort(key=lambda r: (r['_o'], r['v']), reverse=k != 'vencendo')
        for r in rows:
            r.pop('_o')
        out[k] = rows
    return out


def textos_opp(ctx, kpi):
    R, d31 = ctx.ultimo, dt.timedelta(days=31)
    ren = ''
    if kpi['renPilates'] is not None and kpi['renTF'] is not None:
        ren = f' Estão dentro da janela de 30 dias, quando a renovação histórica é de {pct(kpi["renPilates"])} no Pilates e {pct(kpi["renTF"])} no TF.'
    return [
        {'k': 'vencidos', 't': f'Reativar: TF ou Pilates vencido entre {dm(R - D90)} e {dm(R - d31)}',
         'p': f'O último plano venceu há mais de 30 dias e o cliente não comprou outro plano de TF ou Pilates até {dm(R)}.'},
        {'k': 'janela', 't': 'Renovar agora: vencidos nos últimos 30 dias',
         'p': f'Venceram entre {dm(R - D30)} e {dm(R - dt.timedelta(days=1))} e ainda não renovaram.{ren}'},
        {'k': 'vencendo', 't': 'Antecipar: vencem nos próximos 30 dias',
         'p': f'Planos de TF ou Pilates com fim entre {dm(R)} e {dm(R + D30)}. Abordar antes do vencimento.'},
        {'k': 'fisioSemMov', 't': 'Fisio → TF ou Pilates',
         'p': f'O pacote de Atendimento ou FC (7 dias ou mais) terminou entre {dm(R - D90)} e {dm(R)}, o paciente nunca comprou TF ou Pilates e não comprou mais nada desde então.'},
        {'k': 'consSemConv', 't': f'Consultas sem tratamento ({dm(R - D90)} a {dm(R - d31)})',
         'p': f'A consulta foi feita e não houve compra de tratamento depois dela até {dm(R)}.'},
    ]


# ---------------------------------------------------------------- textos automáticos

def periodos(ctx, inicio):
    """Período atual e o mesmo período do ano anterior. No 1º trimestre, usa 12 meses móveis."""
    u = ctx.ultimo
    ini_dados = dt.date(int(inicio[:4]), int(inicio[5:7]), 1)
    a1 = dt.date(u.year, 1, 1)
    if (u - a1).days < 90 and u - dt.timedelta(days=729) >= ini_dados:
        p1 = (u - dt.timedelta(days=364), u)
        p0 = (u - dt.timedelta(days=729), u - dt.timedelta(days=365))
        return p1, p0, 'nos últimos 12 meses', 'nos 12 meses anteriores', 'os últimos 12 meses com os 12 meses anteriores'
    p0 = (dt.date(u.year - 1, 1, 1), ano_antes(u))
    if p0[0] < ini_dados:
        return None
    r1 = rotulo_periodo(a1, u)
    return (a1, u), p0, f'de {r1}', f'no mesmo período de {u.year - 1}', f'{r1} com o mesmo período de {u.year - 1}'


def resumo(L, a, b):
    xs = [x for x in L if a <= x.pg <= b]
    por = collections.defaultdict(lambda: [0.0, 0])
    for x in xs:
        por[x.tipo][0] += x.val
        por[x.tipo][1] += 1
    return {'v': sum(x.val for x in xs), 'q': len(xs), 'cli': len({x.cli for x in xs if x.cli}), 'tipo': por, 'xs': xs}


def tendencias(ctx, D, inicio):
    per = periodos(ctx, inicio)
    if not per:
        D['info']['tendencia'] = ''
        return [], []
    (a1, b1), (a0, b0), lab1, lab0, D['info']['tendencia'] = per
    S1, S0 = resumo(ctx.lanc, a1, b1), resumo(ctx.lanc, a0, b0)
    ins, opp = [], []

    # 1. receita, volume e ticket
    v, q, c = var(S0['v'], S1['v']), var(S0['q'], S1['q']), var(S0['cli'], S1['cli'])
    tk0, tk1 = S0['v'] / S0['q'], S1['v'] / S1['q']
    vt = var(tk0, tk1)
    if v is not None and q is not None:
        if v >= 0 and q < 0:
            titulo = 'A receita cresce pelo ticket, não pelo volume'
        elif v >= 0:
            titulo = 'Receita e volume em alta'
        elif vt and vt > 0:
            titulo = 'Receita em queda, apesar do ticket maior'
        else:
            titulo = 'Receita abaixo do ano anterior'
        ins.append({'t': titulo, 'p': (
            f'{lab1.capitalize()}, a receita somou <b>{brlK(S1["v"])}</b>, contra {brlK(S0["v"])} {lab0} (<b>{dpct(v, 1)}</b>). '
            f'Os lançamentos foram de {nf(S0["q"])} para {nf(S1["q"])} ({dpct(q)}) e os clientes atendidos, de {nf(S0["cli"])} para {nf(S1["cli"])} ({dpct(c)}). '
            f'O ticket médio por lançamento foi de {brl(tk0)} para <b>{brl(tk1)}</b> ({dpct(vt)}).')})

    # 2. consultas
    def cons(S):
        cs = [x for x in S['xs'] if x.tipo == 'Consulta']
        el = {(x.cli, x.pg): x for x in cs if x.cli and ctx.corte and x.pg <= ctx.corte}
        return cs, list(el.values())
    c1, e1 = cons(S1)
    c0, e0 = cons(S0)
    tr = [ctx.trat30(x) for x in e0 + e1]
    media_trat = sum(tr) / len(tr) if tr else 0
    if c0 and c1:
        vc = var(len(c0), len(c1))
        titulo = ('Entrada de consultas em queda' if vc <= -0.1 else 'Entrada de consultas em alta' if vc >= 0.1 else 'Entrada de consultas estável')
        conv_txt = ''
        if len(e0) >= 20 and len(e1) >= 20:
            r0 = sum(map(ctx.converteu, e0)) / len(e0)
            r1 = sum(map(ctx.converteu, e1)) / len(e1)
            conv_txt = f' A conversão em tratamento foi de {pct(r0)} para {pct(r1)}.'
        ins.append({'t': titulo, 'p': (
            f'Foram {nf(len(c0))} consultas {lab0} e <b>{nf(len(c1))}</b> {lab1} (<b>{dpct(vc)}</b>). '
            f'O preço médio foi de {brl(sum(x.val for x in c0) / len(c0))} para {brl(sum(x.val for x in c1) / len(c1))}.{conv_txt}'
            + (f' Cada consulta gerou em média <b>{brl(media_trat)}</b> em tratamento nos 30 dias seguintes.' if tr else ''))})
        falta = len(c0) - len(c1)
        if falta > 0 and media_trat:
            opp.append({'t': 'Recuperar a entrada de consultas', 'p': (
                f'Faltaram {nf(falta)} consultas {lab1} em relação ao ano anterior. Com a média de {brl(media_trat)} em tratamento por consulta, '
                f'a diferença equivale a cerca de <b>{brlK(falta * media_trat)}</b>. É uma simulação ({nf(falta)} × {brl(media_trat)}) que supõe a mesma conversão e o mesmo ticket. '
                'A origem dos pacientes não está na planilha, então antes de agir é preciso saber de qual canal veio a queda.')})

    # 3. tipo que mais cresceu
    tipos = D['tipos']
    g = lambda S, t: S['tipo'][t][0] if t in S['tipo'] else 0.0
    qn = lambda S, t: S['tipo'][t][1] if t in S['tipo'] else 0
    difs = sorted(tipos, key=lambda t: g(S1, t) - g(S0, t), reverse=True)
    destaque = difs[0] if difs and g(S1, difs[0]) > g(S0, difs[0]) else None
    if destaque:
        t = destaque
        txt = (f'A receita de {t} passou de {brlK(g(S0, t))} para <b>{brlK(g(S1, t))}</b>'
               + (f' ({dpct(var(g(S0, t), g(S1, t)))})' if g(S0, t) else '') + f', com {nf(qn(S0, t))} → {nf(qn(S1, t))} lançamentos.')
        cart = D['emp']['Todas']['carteira']
        k = len(ctx.M) - (2 if ctx.parcial else 1)
        if k >= 12 and cart[k - 12]:
            txt += (f' A carteira ativa de TF e Pilates no dia 15 foi de {nf(cart[k - 12])} clientes em {ml(ctx.M[k - 12])} '
                    f'para <b>{nf(cart[k])}</b> em {ml(ctx.M[k])} ({dpct(var(cart[k - 12], cart[k]))}).')
        ins.append({'t': f'{t} é o tipo que mais cresceu', 'p': txt})

    # 4. volume e ticket dos principais tipos
    princ = [t for t in sorted(tipos, key=lambda t: -g(S1, t)) if t not in (destaque, 'Consulta') and qn(S0, t) >= 20 and qn(S1, t) >= 20][:3]
    if princ:
        partes = []
        for t in princ:
            r, qv = var(g(S0, t), g(S1, t)), var(qn(S0, t), qn(S1, t))
            t0, t1 = g(S0, t) / qn(S0, t), g(S1, t) / qn(S1, t)
            verbo = 'cresceu' if r >= 0 else 'caiu'
            partes.append(f'{t} {verbo} {nf(abs(r) * 100)}% ({brlK(g(S0, t))} → {brlK(g(S1, t))}), com volume {dpct(qv)} '
                          f'({nf(qn(S0, t))} → {nf(qn(S1, t))}) e ticket {dpct(var(t0, t1))} ({brl(t0)} → {brl(t1)}).')
        ins.append({'t': 'Volume e ticket dos principais tipos', 'p': ' '.join(partes)})

    # 5. linhas novas, em alta e em queda
    primeira, ultima = {}, {}
    for x in ctx.lanc:
        primeira.setdefault(x.tipo, x.pg)
        ultima[x.tipo] = x.pg
    frases = []
    for t in tipos:
        if g(S0, t) == 0 and g(S1, t) > 0 and primeira[t] >= a0:
            frases.append(f'<b>{t}</b> começou em {ml(mes(primeira[t]))} e já soma {brlK(g(S1, t))} {lab1}.')
    for t in tipos:
        if g(S0, t) > 0 and ultima[t] < ctx.ultimo - D90:
            frases.append(f'<b>{t}</b> não tem lançamentos desde {dm(ultima[t])}.')
    alta = [t for t in tipos if t not in (destaque, *princ) and g(S0, t) >= 10000 and var(g(S0, t), g(S1, t)) >= 0.25]
    for t in sorted(alta, key=lambda t: -var(g(S0, t), g(S1, t)))[:2]:
        frases.append(f'<b>{t}</b> subiu {nf(var(g(S0, t), g(S1, t)) * 100)}% ({brlK(g(S0, t))} → {brlK(g(S1, t))}).')
    queda = [t for t in tipos if g(S0, t) >= 10000 and g(S1, t) > 0 and var(g(S0, t), g(S1, t)) <= -0.2]
    queda = sorted(queda, key=lambda t: var(g(S0, t), g(S1, t)))[:3]
    if queda:
        frases.append(lista_pt(f'{t} ({dpct(var(g(S0, t), g(S1, t)))})' for t in queda) + (' recuaram.' if len(queda) > 1 else ' recuou.'))
    if frases:
        ins.append({'t': 'Linhas novas, em alta e em queda', 'p': ' '.join(frases)})

    # 6. melhores e piores meses
    tot = [sum(D['emp']['Todas']['rev'][t][i] for t in tipos) for i in range(len(ctx.M))]
    fechados = list(range(len(ctx.M) - (1 if ctx.parcial else 0)))
    if len(fechados) >= 6:
        lo = min(fechados, key=lambda i: tot[i])
        hi = max(fechados, key=lambda i: tot[i])
        media = sum(tot[i] for i in fechados) / len(fechados)
        ins.append({'t': 'Meses mais fortes e mais fracos', 'p': (
            f'O maior mês da série foi {ml(ctx.M[hi])} (<b>{brlK(tot[hi])}</b>) e o menor, {ml(ctx.M[lo])} (<b>{brlK(tot[lo])}</b>). '
            f'A média dos meses fechados é de {brlK(media)}.')})

    # 7. venda cruzada
    b = D['emp']['Todas']
    ds, cr = b['dist'], b['cross']
    if all(ds[k]['cl'] for k in ds):
        cresc = ds['1']['gasto'] < ds['2']['gasto'] < ds['3']['gasto'] < ds['4']['gasto']
        ins.append({'t': 'Quem compra mais de um tipo vale mais' if cresc else 'Clientes com mais de um tipo', 'p': (
            f'Clientes com 1 tipo gastaram em média {brl(ds["1"]["gasto"])} no período. Com 2 tipos, {brl(ds["2"]["gasto"])}; '
            f'com 3, {brl(ds["3"]["gasto"])}; com 4 ou mais, {brl(ds["4"]["gasto"])}. Quem comprou 2 tipos ou mais ({pct(cr["multi"])} dos clientes) '
            f'gerou <b>{pct(cr["revMulti"])} da receita</b>. Sem contar consulta, produtos e avaliações, {pct(cr["multiCore"])} compram dois serviços diferentes ou mais.')})

    # oportunidades: fisio → movimento, renovação, combinações
    f = b['fisioMov']
    if f['base'] >= 20:
        nlist = len(D['opp']['fisioSemMov'])
        opp.append({'t': 'Levar o paciente da fisio para o TF ou Pilates', 'p': (
            f'Dos {nf(f["base"])} pacientes de Atendimento ou FC, <b>{nf(f["mov"])} ({pct(f["mov"] / f["base"])})</b> compraram TF ou Pilates em algum momento. '
            f'A alta da fisio é o ponto natural de oferta. A lista “Fisio → TF ou Pilates” traz {nf(nlist)} pacientes que terminaram o pacote nos últimos 90 dias.')})
    k = b['kpi']
    if k['renPilates'] is not None and k['renTF'] is not None and min(k['renPilatesN'], k['renTFN']) >= 20:
        a, z = ('TF', 'Pilates') if k['renTF'] < k['renPilates'] else ('Pilates', 'TF')
        opp.append({'t': f'Renovação de {a} abaixo da de {z}', 'p': (
            f'Até 30 dias depois do vencimento, <b>{pct(k["renPilates"])}</b> dos planos de Pilates são renovados, contra <b>{pct(k["renTF"])}</b> dos de TF '
            f'(planos vencidos até {dm(ctx.corte)}). Acompanhar essa taxa todo mês é o indicador mais direto de retenção.')})
    mt = b['matrix']
    pares = []
    for i, ti in enumerate(mt['tipos']):
        for j in range(i + 1, len(mt['tipos'])):
            tj = mt['tipos'][j]
            if not ({ti, tj} & {'Consulta', 'Atendimento', 'FC'}):
                pares.append((mt['m'][i][j], ti, tj))
    frases = []
    if pares:
        nmax, ti, tj = max(pares)
        if nmax >= 10:
            frases.append(f'{ti} + {tj} é a combinação mais comum fora da fisio ({nf(nmax)} clientes).')
    sozinhos = collections.Counter()
    clientes_t = collections.Counter()
    tipos_cli = collections.defaultdict(set)
    for x in ctx.lanc:
        if x.cli:
            tipos_cli[x.cli].add(x.tipo)
    for ts in tipos_cli.values():
        for t in ts:
            clientes_t[t] += 1
            if len(ts) == 1:
                sozinhos[t] += 1
    cand = [t for t in clientes_t if clientes_t[t] >= 50 and t not in {'Consulta', 'Atendimento', 'FC', *NAO_TRATAMENTO}]
    if cand:
        t = max(cand, key=lambda t: sozinhos[t] / clientes_t[t])
        s = sozinhos[t] / clientes_t[t]
        if s >= 0.4:
            frases.append(f'No {t}, {pct(s)} dos {nf(clientes_t[t])} clientes não compraram nenhum outro tipo, o que faz dele uma porta de entrada pouco aproveitada para avaliação, fisio e outros serviços.')
    if frases:
        opp.append({'t': 'Combos e portas de entrada', 'p': ' '.join(frases)})
    return ins, opp


def calcular(lanc, extra, cfg, hoje, agora):
    ctx = Ctx(lanc, cfg, hoje)
    ctx.matriz_excluir = set(cfg.get('matriz_excluir', []))
    M = ctx.M
    tot_tipo = collections.Counter()
    for x in lanc:
        if mes(x.pg) in ctx.mi:
            tot_tipo[x.tipo] += x.val
    tipos = sorted({x.tipo for x in lanc}, key=lambda t: (t == 'Outros', -tot_tipo[t], t))
    emps = [e for e, _ in collections.Counter(x.emp for x in lanc).most_common()]

    D = {'meses': M, 'tipos': tipos, 'top': [t for t in tipos if t != 'Outros'][:5], 'emps': ['Todas'] + emps,
         'parcial': ctx.parcial, 'emp': {}}
    D['emp']['Todas'] = bloco(ctx, lanc, tipos)
    for e in emps:
        D['emp'][e] = bloco(ctx, [x for x in lanc if x.emp == e], tipos)

    comb = collections.Counter(' + '.join(sorted({x.emp for x in xs})) for xs in ctx.compras.values())
    D['empresasCli'] = dict(comb.most_common())
    D['opp'] = oportunidades(ctx)
    D['oppTxt'] = textos_opp(ctx, D['emp']['Todas']['kpi'])

    fora = [x for x in lanc if mes(x.pg) not in ctx.mi]
    sem = [x for x in lanc if not x.cli]
    D['qual'] = {
        'total': len(lanc), 'ultimo': dma(ctx.ultimo),
        'foraN': len(fora), 'foraV': r2(sum(x.val for x in fora)),
        'semNomeN': len(sem), 'semNomeV': r2(sum(x.val for x in sem)),
        'semNomeMes': sorted({mes(x.pg) for x in sem if mes(x.pg) in ctx.mi}),
        'nomesBrutos': len({x.nome for x in lanc if x.nome}), 'nomesUnif': len(ctx.compras),
        'durErr': sum(1 for x in lanc if x.dur is not None and (x.dur > 3660 or x.dur < 0)),
        'col1Err': sum(1 for x in lanc if x.col1 is not None and int(x.col1) != x.pg.month),
        'formaErr': sum(1 for x in lanc if x.forma and not re.fullmatch(r'(pix|dinheiro|debito|credito( \d+x)?|boleto|transferencia|cortesia)', norm(x.forma))),
        'futuros': extra['futuros'], 'semValor': extra['semValor'], 'tiposNovos': extra['tiposNovos'],
    }
    # grafias da planilha reunidas em cada tipo (vai criptografado, junto com os dados)
    grafias = collections.defaultdict(dict)
    for x in lanc:
        if x.trat:
            grafias[x.tipo].setdefault(norm(x.trat), re.sub(r'[<>&"]', '', x.trat))
    D['qual']['grupos'] = [[t, sorted(grafias[t].values(), key=norm)] for t in tipos
                           if len(grafias[t]) > 1 or (grafias[t] and norm(t) not in grafias[t])]
    D['info'] = {
        'aba': cfg['aba'], 'ultimo': dma(ctx.ultimo), 'gerado': agora.strftime('%d/%m/%Y %H:%M'),
        'primeira': dm(lanc[0].pg), 'corte': dm(ctx.corte) if ctx.corte else None,
        'corte90': dm(ctx.corte90), 'periodoTodo': f'{dm(lanc[0].pg)} a {dm(ctx.ultimo)}',
    }
    D['ins'], D['insOpp'] = tendencias(ctx, D, cfg.get('inicio', M[0]))
    return D


# ---------------------------------------------------------------- saída

def criptografar(texto, senha):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    chave = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=SAL, iterations=ITERACOES).derive(senha.encode('utf-8'))
    iv = os.urandom(12)
    ct = AESGCM(chave).encrypt(iv, texto.encode('utf-8'), None)
    b64 = lambda b: base64.b64encode(b).decode('ascii')
    return {'s': b64(SAL), 'i': b64(iv), 'c': b64(ct), 'n': ITERACOES}


def gravar(D, cfg, saida):
    senha = os.environ.get('PAINEL_SENHA', '')
    if len(senha) < 8:
        raise ErroPainel('o secret PAINEL_SENHA não está cadastrado ou tem menos de 8 caracteres')
    modelo = (RAIZ / 'modelo.html').read_text(encoding='utf-8')
    if '/*__CIFRA__*/null' not in modelo:
        raise ErroPainel('modelo.html sem o marcador /*__CIFRA__*/null')
    cifra = criptografar(json.dumps(D, ensure_ascii=False, separators=(',', ':')), senha)
    pagina = (modelo.replace('/*__CIFRA__*/null', json.dumps(cifra))
              .replace('__TITULO__', html.escape(cfg.get('titulo', 'Painel'))))
    saida.mkdir(parents=True, exist_ok=True)
    (saida / 'index.html').write_text(pagina, encoding='utf-8')
    (saida / 'robots.txt').write_text('User-agent: *\nDisallow: /\n', encoding='utf-8')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--xlsx', help='lê uma cópia local (.xlsx) em vez da planilha online (teste)')
    ap.add_argument('--hoje', help='data de referência AAAA-MM-DD (teste)')
    ap.add_argument('--saida', default=str(RAIZ / 'site'), help='pasta do site gerado')
    ap.add_argument('--json', help='grava também os dados calculados, sem criptografia, neste arquivo (só para teste local)')
    a = ap.parse_args()
    try:
        try:
            cfg = json.loads((RAIZ / 'config.json').read_text(encoding='utf-8-sig'))
        except json.JSONDecodeError as e:
            raise ErroPainel(f'config.json com erro de formato na linha {e.lineno} (confira vírgulas e aspas)') from None
        agora = dt.datetime.now(BRT)
        hoje = dt.date.fromisoformat(a.hoje) if a.hoje else agora.date()
        linhas = ler_xlsx(a.xlsx, cfg) if a.xlsx else ler_google(cfg)
        lanc, extra = montar(linhas, cfg, hoje)
        D = calcular(lanc, extra, cfg, hoje, agora)
        if a.json:
            Path(a.json).write_text(json.dumps(D, ensure_ascii=False, indent=1), encoding='utf-8')
        gravar(D, cfg, Path(a.saida))
    except ErroPainel as e:
        print(f'ERRO: {e}. O painel publicado não foi alterado.', file=sys.stderr)
        sys.exit(1)
    except Exception as e:  # noqa: BLE001
        # Sem a mensagem original: ela pode conter dados da planilha e o log do GitHub é público.
        print(f'ERRO inesperado ({type(e).__name__}). O painel publicado não foi alterado.', file=sys.stderr)
        for fr in traceback.extract_tb(e.__traceback__):
            print(f'  {Path(fr.filename).name}:{fr.lineno} em {fr.name}', file=sys.stderr)
        sys.exit(1)
    novos = extra['tiposNovos']
    print(f'Painel gerado: {len(lanc)} lançamentos, último pagamento em {D["info"]["ultimo"]}.')
    if novos:
        print(f'Aviso: {len(novos)} tratamento(s) fora dos grupos do config.json aparecem como tipo próprio.')


if __name__ == '__main__':
    main()
