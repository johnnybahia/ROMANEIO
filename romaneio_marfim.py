# -*- coding: utf-8 -*-
"""
Romaneio de Expedicao - Marfim
Le PDFs de DANFE de uma pasta, gera romaneio em PDF, envia por e-mail e imprime.
"""
import os, re, json, sys, glob, time, base64, subprocess, threading, smtplib, shutil
from datetime import datetime
from email.message import EmailMessage

import pdfplumber
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, CondPageBreak, HRFlowable)

from datetime import date, timedelta

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
CACHE_PATH = os.path.join(APP_DIR, "indice_cache.json")
# Senha do servidor das notas - fica fora do config.json de proposito.
CRED_PATH = os.path.join(APP_DIR, "credenciais_rede.json")

CONFIG_PADRAO = {
    "pasta_notas": "",
    "pasta_saida": "",
    "servidor": {
        "ativo": False,
        "host": "",
        "compartilhamento": "C$",
        "pasta_remota": "",
        "usuario": "",
        "dominio": "",
        "usar_credenciais_windows": False
    },
    "empresa": "MARFIM IND TEXTIL DO CEARA LTDA",
    "titulo": "ROMANEIO DE EXPEDICAO",
    "impressora": "",
    "arquivo_contador": "contador_romaneio.json",
    "dpi_impressao": 200,
    "email": {
        "modo_envio": "smtp",
        "smtp": "",
        "porta": 587,
        "seguranca": "starttls",
        "autenticar": True,
        "usuario": "",
        "senha": "",
        "remetente": "",
        "destinatarios": [],
        "assunto": "Romaneio de Expedicao - {romaneio}",
        "corpo": "Segue em anexo o romaneio de expedicao {romaneio}.\n\nNotas: {notas}\nVolumes: {volumes}\n"
    }
}


def carregar_config():
    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(CONFIG_PADRAO, f, indent=2, ensure_ascii=False)
        return dict(CONFIG_PADRAO)
    with open(CONFIG_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    base = dict(CONFIG_PADRAO)
    base.update(cfg)
    e = dict(CONFIG_PADRAO["email"])
    e.update(cfg.get("email", {}))
    base["email"] = e
    s = dict(CONFIG_PADRAO["servidor"])
    s.update(cfg.get("servidor", {}))
    base["servidor"] = s
    return base


# ----------------------------------------------------------------------------
# PASTA DAS NOTAS EM SERVIDOR REMOTO
# ----------------------------------------------------------------------------
# As DANFEs ficam num servidor da rede (ex.: 10.121.19.8, pasta
# C:\Users\marfimbh1\Desktop\NOTAS). O programa le direto de la pelo caminho
# UNC  \\10.121.19.8\C$\Users\marfimbh1\Desktop\NOTAS  autenticando com o
# usuario e a senha do servidor.
#
# A senha NAO fica no config.json. Ela e gravada em credenciais_rede.json,
# protegida pela conta do Windows (DPAPI) quando o pywin32 esta instalado.
# Tambem pode vir da variavel de ambiente ROMANEIO_SENHA_REDE.

SENHA_AUSENTE = '__SENHA_AUSENTE__'

ERROS_REDE = {
    5: 'Acesso negado. O usuario informado nao tem permissao nessa pasta.\n'
       'O compartilhamento administrativo (C$) so aceita contas que sejam '
       'administradoras do servidor. Se essa conta nao for administradora, '
       'compartilhe a pasta NOTAS no servidor (botao direito > Propriedades > '
       'Compartilhamento) e use o nome do compartilhamento aqui.',
    51: 'A rede esta fora do ar ou o servidor nao respondeu.',
    53: 'Servidor nao encontrado na rede. Confira o IP e se o servidor esta ligado.',
    59: 'Erro de rede inesperado ao falar com o servidor.',
    64: 'O servidor recusou a conexao (compartilhamento removido ou servico parado).',
    67: 'Compartilhamento nao encontrado no servidor. Confira o nome apos o IP.',
    71: 'O servidor atingiu o limite de conexoes simultaneas. Tente de novo em instantes.',
    86: 'A senha desta conta nao confere no servidor.',
    1219: 'Ja existe uma conexao aberta com esse servidor usando outro usuario.\n'
          'Feche as pastas de rede abertas e use Arquivo > Reconectar ao servidor.',
    1326: 'Usuario ou senha incorretos.',
    1327: 'Contas sem senha nao podem acessar pastas pela rede.',
    1331: 'A conta esta desativada no servidor.',
    1330: 'A senha desta conta expirou no servidor.',
    1907: 'A senha desta conta precisa ser trocada no servidor.',
    1909: 'A conta esta bloqueada no servidor.',
    1460: 'Tempo esgotado esperando o servidor responder.',
}


def normalizar_caminho(texto, windows=False):
    """Padroniza um caminho: tira aspas e espacos; usa \\ em caminhos do Windows.

    Um caminho de outro sistema (/mnt/notas) e devolvido como esta, para o
    programa continuar testavel fora do Windows.
    """
    p = (texto or '').strip().strip('"')
    if (windows or os.name == 'nt' or p.startswith('\\\\') or '\\' in p
            or re.match(r'^[A-Za-z]:[\\/]', p)):
        p = p.replace('/', '\\')
    return p if len(p) <= 3 else p.rstrip('\\/')


def montar_caminho_rede(srv):
    """Monta o caminho UNC completo da pasta de notas no servidor.

    Aceita a pasta remota como caminho local do servidor (C:\\Users\\...\\NOTAS),
    como caminho relativo ao compartilhamento (Users\\...\\NOTAS) ou ja como UNC.
    """
    pasta = normalizar_caminho(srv.get('pasta_remota'), windows=True)
    if pasta.startswith('\\\\'):
        return pasta
    host = normalizar_caminho(srv.get('host'), windows=True).lstrip('\\')
    if not host:
        return ''
    share = normalizar_caminho(srv.get('compartilhamento'), windows=True).strip('\\')
    letra = re.match(r'^([A-Za-z]):\\?(.*)$', pasta)
    if letra:
        share = f'{letra.group(1).upper()}$'
        pasta = letra.group(2)
    resto = pasta.strip('\\')
    return '\\\\' + host + '\\' + (share or 'C$') + (('\\' + resto) if resto else '')


def raiz_compartilhamento(unc):
    """Devolve \\\\servidor\\compartilhamento de um caminho UNC."""
    partes = [p for p in normalizar_caminho(unc, windows=True).split('\\') if p]
    if len(partes) < 2:
        return ''
    return '\\\\' + partes[0] + '\\' + partes[1]


def host_do_caminho(unc):
    partes = [p for p in normalizar_caminho(unc, windows=True).split('\\') if p]
    return partes[0] if partes else ''


def usuario_rede(srv, host=''):
    """Monta o login como o Windows espera: SERVIDOR\\usuario."""
    usuario = (srv.get('usuario') or '').strip()
    if not usuario or '\\' in usuario or '@' in usuario:
        return usuario
    dominio = (srv.get('dominio') or '').strip() or host or \
        normalizar_caminho(srv.get('host')).lstrip('\\')
    return f'{dominio}\\{usuario}' if dominio else usuario


# ------------------------------------------------------- senha do servidor
def _chave_credencial(host, usuario):
    return f"{(host or '').strip().lower()}|{(usuario or '').strip().lower()}"


def _proteger(texto):
    """Cifra a senha com a conta do Windows (DPAPI). Sem pywin32, so codifica."""
    try:
        import win32crypt
        dados = win32crypt.CryptProtectData(texto.encode('utf-8'), 'romaneio',
                                            None, None, None, 0)
        return 'dpapi:' + base64.b64encode(dados).decode('ascii')
    except Exception:
        return 'b64:' + base64.b64encode(texto.encode('utf-8')).decode('ascii')


def _desproteger(valor):
    if not valor:
        return ''
    if valor.startswith('dpapi:'):
        try:
            import win32crypt
            _rot, dados = win32crypt.CryptUnprotectData(
                base64.b64decode(valor[6:]), None, None, None, 0)
            return dados.decode('utf-8')
        except Exception:
            return ''
    if valor.startswith('b64:'):
        try:
            return base64.b64decode(valor[4:]).decode('utf-8')
        except Exception:
            return ''
    return valor


def ler_credenciais():
    if not os.path.exists(CRED_PATH):
        return {}
    try:
        with open(CRED_PATH, encoding='utf-8') as f:
            dados = json.load(f)
        return dados if isinstance(dados, dict) else {}
    except Exception:
        return {}


def salvar_senha_rede(host, usuario, senha):
    """Guarda (ou apaga) a senha do servidor no arquivo local de credenciais."""
    host = normalizar_caminho(host).lstrip('\\')
    dados = ler_credenciais()
    chave = _chave_credencial(host, usuario)
    if senha:
        dados[chave] = _proteger(senha)
    elif chave in dados:
        dados.pop(chave)
    else:
        return True
    try:
        with open(CRED_PATH, 'w', encoding='utf-8') as f:
            json.dump(dados, f, indent=2)
        if os.name != 'nt':
            os.chmod(CRED_PATH, 0o600)
        return True
    except Exception:
        return False


def senha_rede(srv):
    """Procura a senha: variavel de ambiente, arquivo local, config.json."""
    env = os.environ.get('ROMANEIO_SENHA_REDE')
    if env:
        return env
    host = normalizar_caminho(srv.get('host')).lstrip('\\')
    guardada = ler_credenciais().get(_chave_credencial(host, srv.get('usuario')))
    if guardada:
        return _desproteger(guardada)
    return srv.get('senha') or ''


def migrar_senha_config(cfg):
    """Tira a senha do config.json e passa para o arquivo de credenciais.

    O config.json vai junto com o programa (e pode ir para o repositorio),
    entao senha nenhuma deve ficar guardada nele.
    """
    srv = cfg.get('servidor') or {}
    senha = srv.pop('senha', '')
    if not senha:
        return False
    salvar_senha_rede(srv.get('host'), srv.get('usuario'), senha)
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass
    return True


# ------------------------------------------------------------ conexao SMB
def _texto_processo(dados):
    for cp in ('cp850', 'cp1252', 'utf-8'):
        try:
            return dados.decode(cp)
        except Exception:
            continue
    return dados.decode('utf-8', errors='replace')


def _rodar_oculto(cmd, timeout=40):
    """Roda um comando sem abrir janela preta. Devolve (codigo, saida)."""
    extras = {}
    if os.name == 'nt':
        extras['creationflags'] = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)
    try:
        p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                           timeout=timeout, **extras)
    except subprocess.TimeoutExpired:
        return 1460, ''
    except Exception as ex:
        return 1, str(ex)
    return p.returncode, _texto_processo(p.stdout or b'').strip()


def _desconectar_rede(raiz):
    """Derruba a conexao atual com o compartilhamento, se houver."""
    try:
        import win32wnet
        win32wnet.WNetCancelConnection2(raiz, 0, True)
        return
    except Exception:
        pass
    _rodar_oculto(['net', 'use', raiz, '/delete', '/y'], timeout=25)


def _abrir_conexao(raiz, usuario, senha):
    """Autentica no compartilhamento. Devolve (codigo do Windows, texto)."""
    try:
        import win32wnet
        import win32netcon
    except ImportError:
        win32wnet = None
    if win32wnet is not None:
        recurso = win32wnet.NETRESOURCE()
        recurso.dwType = win32netcon.RESOURCETYPE_DISK
        recurso.lpRemoteName = raiz
        try:
            win32wnet.WNetAddConnection2(recurso, senha or None, usuario or None, 0)
            return 0, ''
        except Exception as ex:
            cod = ex.args[0] if ex.args and isinstance(ex.args[0], int) else 1
            texto = ex.args[2] if len(ex.args) > 2 else str(ex)
            return cod, str(texto)
    cmd = ['net', 'use', raiz]
    if senha:
        cmd.append(senha)
    if usuario:
        cmd.append('/user:' + usuario)
    cmd.append('/persistent:no')
    cod, saida = _rodar_oculto(cmd)
    if cod:
        # "Erro do sistema 5 ocorreu." / "System error 1326 has occurred."
        m = re.search(r'err(?:o|or)[^\d]{0,30}(\d{1,5})', saida, re.I)
        if m:
            cod = int(m.group(1))
    return cod, saida


def mensagem_erro_rede(cod, texto, raiz, usuario):
    detalhe = ERROS_REDE.get(cod) or (texto or '').strip() or f'Erro {cod}.'
    return (f'Nao foi possivel acessar {raiz}'
            + (f' como {usuario}' if usuario else '')
            + f'.\n\n{detalhe}'
            + (f'\n\n(codigo {cod})' if cod else ''))


def conectar_rede(caminho, srv, log=None, reconectar=False):
    """Garante acesso a uma pasta de rede. Devolve (caminho, erro).

    erro None = pasta pronta para uso. erro SENHA_AUSENTE = falta a senha.
    """
    caminho = normalizar_caminho(caminho, windows=True)
    raiz = raiz_compartilhamento(caminho)
    if not raiz:
        return caminho, f'Caminho de rede invalido: {caminho}'
    if not reconectar and os.path.isdir(caminho):
        return caminho, None
    if os.name != 'nt':
        return caminho, (f'Sem acesso a {caminho}. Pastas compartilhadas do '
                         'Windows so podem ser montadas pelo proprio Windows.')
    usa_windows = bool(srv.get('usar_credenciais_windows'))
    usuario = '' if usa_windows else usuario_rede(srv, host_do_caminho(caminho))
    senha = '' if usa_windows else senha_rede(srv)
    if usuario and not senha:
        return caminho, SENHA_AUSENTE
    if reconectar:
        _desconectar_rede(raiz)
    if log:
        log(f'Conectando em {raiz}...')
    cod, texto = _abrir_conexao(raiz, usuario, senha)
    if cod == 1219:
        _desconectar_rede(raiz)
        cod, texto = _abrir_conexao(raiz, usuario, senha)
    if os.path.isdir(caminho):
        return caminho, None
    if cod:
        return caminho, mensagem_erro_rede(cod, texto, raiz, usuario)
    return caminho, (f'O servidor respondeu, mas a pasta nao existe:\n{caminho}\n\n'
                     'Confira o caminho dentro do servidor.')


def preparar_pasta_notas(cfg, pasta='', log=None, reconectar=False):
    """Resolve a pasta das notas (local ou no servidor) e garante o acesso."""
    srv = dict(cfg.get('servidor') or {})
    pasta = normalizar_caminho(pasta) or normalizar_caminho(cfg.get('pasta_notas'))
    if srv.get('ativo') and not pasta:
        pasta = montar_caminho_rede(srv)
    if not pasta.startswith('\\\\'):
        return pasta, None
    if not srv.get('ativo') or host_do_caminho(pasta).lower() != \
            normalizar_caminho(srv.get('host')).lstrip('\\').lower():
        srv = {'usar_credenciais_windows': True}
    return conectar_rede(pasta, srv, log=log, reconectar=reconectar)


# ----------------------------------------------------------------------------
# PARSER DANFE
# ----------------------------------------------------------------------------
LINHAS = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
# Alterar esta versão força o cache a reler as DANFEs após mudanças no parser.
PARSER_VERSAO = "2026-08-20-v2"


def _somar_produtos_do_texto(texto):
    """Extrai QTD e VALOR TOTAL de cada linha numérica da grade de produtos.

    Algumas DANFEs não desenham a linha horizontal inferior do último item.
    Nesse caso, pdfplumber.extract_table() pode ignorar justamente o último
    produto. Para o checksum, usamos também o texto bruto da grade.
    """
    if not texto:
        return 0.0, 0.0, 0

    bloco = texto
    partes = re.split(r'DADOS DO PRODUTO/SERVIÇO', bloco, maxsplit=1)
    if len(partes) == 2:
        bloco = partes[1]
    partes = re.split(r'DADOS ADICIONAIS', bloco, maxsplit=1)
    bloco = partes[0]

    # Ex.: 54074200 051 5101 PRS 360,0000 1,200000 432,00 ...
    rx = re.compile(
        r'\b\d{8}\s+\d{3}\s+\d{4}\s+\S+\s+'
        r'([\d.,]+)\s+[\d.,]+\s+([\d.,]+)\s+0[,.]00',
        re.MULTILINE)

    soma_q = 0.0
    soma_v = 0.0
    n = 0
    for m in rx.finditer(bloco):
        q = num(m.group(1))
        v = num(m.group(2))
        if q is None or v is None:
            continue
        soma_q += q
        soma_v += v
        n += 1
    return round(soma_q, 4), round(soma_v, 4), n


def num(s):
    if not s:
        return None
    
    # Captura a ultima sequencia que pareca um numero, aceitando ponto ou virgula
    achados = re.findall(r'-?\d+(?:[.,]\d+)*', s.replace('\n', ' '))
    if not achados:
        return None
    
    val = achados[-1]
    last_comma = val.rfind(',')
    last_dot = val.rfind('.')
    
    try:
        # Se a virgula for o ultimo separador (ex: 1.234,56 ou 12,50), e padrao BR
        if last_comma > last_dot:
            return float(val.replace('.', '').replace(',', '.'))
        # Se o ponto for o ultimo separador (ex: 1,234.56 ou 12.50), e padrao US
        else:
            return float(val.replace(',', ''))
    except ValueError:
        return None


def somar_celula(texto):
    """Soma multiplos valores, lidando com quebras de linha e detectando virgula ou ponto decimal dinamicamente."""
    if not texto:
        return None
    
    achados = re.findall(r'-?\d+(?:[.,]\d+)*', str(texto).replace('\n', ' '))
    if not achados:
        return None
        
    soma = 0
    achou = False
    for val in achados:
        last_comma = val.rfind(',')
        last_dot = val.rfind('.')
        
        try:
            if last_comma > last_dot:
                soma += float(val.replace('.', '').replace(',', '.'))
            else:
                soma += float(val.replace(',', ''))
            achou = True
        except ValueError:
            continue
            
    return round(soma, 4) if achou else None


def _seq(palavras, seq):
    n = len(seq)
    for i in range(len(palavras) - n + 1):
        g = palavras[i:i + n]
        if [w['text'] for w in g] == seq and len({round(w['top']) for w in g}) == 1:
            return g[0], g[-1]
    return None, None


def _campo(page, palavras, seq, x1=None, dy=12):
    a, b = _seq(palavras, seq)
    if not a:
        return None
    x1 = x1 if x1 else min(a['x0'] + 230, page.width - 2)
    if x1 <= a['x0']:
        x1 = min(a['x0'] + 230, page.width - 2)
    txt = page.crop((max(a['x0'] - 3, 0), b['bottom'] + 0.5, x1,
                     b['bottom'] + dy)).extract_text() or ''
    return re.sub(r'\s+', ' ', txt).strip()


def _grade_produtos(page):
    """Localiza a grade de itens pelas linhas verticais abaixo do cabecalho."""
    palavras = page.extract_words()
    ncm = [w for w in palavras if w['text'] == 'NCM/SH']
    if not ncm:
        return []
    cab = ncm[0]
    verticais = sorted([l for l in page.lines
                        if abs(l['x1'] - l['x0']) < 0.5 and l['top'] >= cab['bottom'] - 2],
                       key=lambda l: l['top'])
    if not verticais:
        return []
    topo = verticais[0]['top']
    base = verticais[0]['bottom']
    for l in verticais:
        if l['top'] <= base + 1.5:
            base = max(base, l['bottom'])
        else:
            break
    seg = [l for l in verticais if l['top'] <= base + 1.5]
    x0 = min(l['x0'] for l in seg)
    x1 = max(l['x0'] for l in seg)
    return page.crop((x0 - 0.5, topo - 0.5, x1 + 0.5, base + 0.5)).extract_table(LINHAS) or []


def ler_danfe(caminho):
    d = {'arquivo': os.path.basename(caminho), 'caminho': caminho, 'itens': [],
         'volumes': None, 'peso_bruto': None, 'tot_prod': None, 'qtd_total': None,
         'nota': None, 'serie': None, 'chave': None, 'cliente': None,
         'cnpj': None, 'emissao': None, 'transportadora': None, 'paginas': 0,
         '_soma_text_q': 0.0, '_soma_text_v': 0.0, '_soma_text_n': 0}
    with pdfplumber.open(caminho) as pdf:
        d['paginas'] = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            for linha in _grade_produtos(page):
                if len(linha) < 9:
                    continue
                cod = (linha[0] or '').strip()
                desc = re.sub(r'\s+', ' ', (linha[1] or '').strip())
                qtd = somar_celula(linha[6])
                vtot = somar_celula(linha[8])
                if qtd is None or vtot is None:
                    if desc and d['itens']:
                        d['itens'][-1]['desc'] += ' ' + desc
                    continue
                d['itens'].append({'cod': cod, 'desc': desc,
                                   'unid': (linha[5] or '').strip(),
                                   'qtd': qtd, 'vunit': num(linha[7]), 'vtot': vtot})
            texto = page.extract_text() or ''
            sq_txt, sv_txt, n_txt = _somar_produtos_do_texto(texto)
            d['_soma_text_q'] += sq_txt
            d['_soma_text_v'] += sv_txt
            d['_soma_text_n'] += n_txt
            m = re.search(r'[OQ]TD\.?\s*\n?TOTAL:\s*([\d\.,]+)', texto)
            if m:
                d['qtd_total'] = num(m.group(1))
            if i == 0:
                pl = page.extract_words()
                cnpjs = sorted([w for w in pl if w['text'] == 'CNPJ/CPF'], key=lambda w: w['top'])
                x_cnpj = cnpjs[0]['x0'] - 4 if cnpjs else None
                d['cliente'] = _campo(page, pl, ['NOME/RAZÃO', 'SOCIAL'], x1=x_cnpj)
                d['cnpj'] = _campo(page, pl, ['CNPJ/CPF'],
                                   x1=(cnpjs[0]['x0'] + 120 if cnpjs else None))
                d['emissao'] = _campo(page, pl, ['DATA', 'DA', 'EMISSÃO'], x1=page.width - 2)
                rz_a, _ = _seq(pl, ['RAZÃO', 'SOCIAL'])
                x_frete = None
                if rz_a:
                    cands = [w['x0'] for w in pl
                             if w['text'] == 'FRETE' and abs(w['top'] - rz_a['top']) < 6]
                    x_frete = min(cands) - 4 if cands else None
                d['transportadora'] = _campo(page, pl, ['RAZÃO', 'SOCIAL'], x1=x_frete)
                esp = [w for w in pl if w['text'] == 'ESPÉCIE']
                d['volumes'] = num(_campo(page, pl, ['QUANTIDADE'],
                                          x1=(esp[0]['x0'] - 4 if esp else None)))
                pesos = sorted([w for w in pl if w['text'] == 'PESO'], key=lambda w: w['x0'])
                x_pl = pesos[1]['x0'] - 4 if len(pesos) > 1 else None
                d['peso_bruto'] = num(_campo(page, pl, ['PESO', 'BRUTO'], x1=x_pl))
                d['tot_prod'] = num(_campo(page, pl, ['VALOR', 'TOTAL', 'DOS', 'PRODUTOS'],
                                           x1=page.width - 2))
                m = re.search(r'N[ºo]\s*(\d{3}\.\d{3}\.\d{3})', texto)
                d['nota'] = m.group(1).lstrip('0.') if m else None
                m = re.search(r'SÉRIE:\s*(\d+)', texto)
                d['serie'] = m.group(1) if m else None
                m = re.search(r'CHAVE DE ACESSO\s*\n([\d ]{50,})', texto)
                d['chave'] = re.sub(r'\D', '', m.group(1)) if m else None
    d['status'] = _validar(d)
    return d


def _validar(d):
    if not d['itens']:
        return 'ILEGIVEL'

    # Para a validação fiscal, prefira a soma obtida do texto bruto da grade.
    # Isso evita que um último item sem linha horizontal inferior seja perdido
    # pelo extract_table(), como acontece na NF 29357.
    if d.get('_soma_text_n'):
        soma_v = round(d['_soma_text_v'], 2)
        soma_q = round(d['_soma_text_q'], 2)
    else:
        soma_v = round(sum(i['vtot'] for i in d['itens']), 2)
        soma_q = round(sum(i['qtd'] for i in d['itens']), 2)

    d['_divergencia_valor'] = (
        round(soma_v - d['tot_prod'], 2)
        if d['tot_prod'] is not None else None
    )
    d['_divergencia_qtd'] = (
        round(soma_q - d['qtd_total'], 2)
        if d['qtd_total'] is not None else None
    )

    if d['tot_prod'] is not None and abs(d['_divergencia_valor']) > 0.02:
        return 'DIVERGENCIA'
    if d['qtd_total'] is not None and abs(d['_divergencia_qtd']) > 0.02:
        return 'DIVERGENCIA'
    if d['tot_prod'] is None and d['qtd_total'] is None:
        return 'SEM CHECKSUM'
    return 'OK'


# ----------------------------------------------------------------------------
# INDICE / CACHE
# ----------------------------------------------------------------------------
def varrer_pasta(pasta, log=None):
    cache = {}
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, encoding='utf-8') as f:
                cache = json.load(f)
        except Exception:
            cache = {}
    resultado, vistos, chaves = [], set(), set()
    arquivos = sorted(glob.glob(os.path.join(pasta, '*.pdf')) +
                      glob.glob(os.path.join(pasta, '*.PDF')))
    for caminho in arquivos:
        try:
            st = os.stat(caminho)
        except OSError:
            # Arquivo sumiu ou a rede oscilou no meio da varredura.
            continue
        assinatura = f"{PARSER_VERSAO}:{st.st_mtime_ns}:{st.st_size}"
        chave_cache = os.path.abspath(caminho)
        item = cache.get(chave_cache)
        if not item or item.get('_assin') != assinatura:
            if log:
                log(f"Lendo {os.path.basename(caminho)}...")
            try:
                item = ler_danfe(caminho)
            except Exception as e:
                item = {'arquivo': os.path.basename(caminho), 'caminho': caminho,
                        'itens': [], 'status': 'ERRO', 'erro': str(e), 'nota': None,
                        'serie': None, 'cliente': None, 'emissao': None,
                        'volumes': None, 'peso_bruto': None, 'chave': None,
                        'transportadora': None, 'cnpj': None, '_soma_text_q': 0.0,
                         '_soma_text_v': 0.0, '_soma_text_n': 0}
            item['_assin'] = assinatura
        item['caminho'] = caminho
        cache[chave_cache] = item
        ch = item.get('chave')
        if ch and ch in chaves:
            continue
        if ch:
            chaves.add(ch)
        vistos.add(chave_cache)
        resultado.append(item)
    for k in [k for k in cache if k not in vistos]:
        cache.pop(k, None)
    try:
        with open(CACHE_PATH, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception:
        pass
    resultado.sort(key=lambda d: (d.get('nota') or ''))
    return resultado


# ----------------------------------------------------------------------------
# NUMERO DO ROMANEIO
# ----------------------------------------------------------------------------
def proximo_romaneio(cfg):
    caminho = cfg.get('arquivo_contador') or 'contador_romaneio.json'
    if not os.path.isabs(caminho):
        caminho = os.path.join(APP_DIR, caminho)
    dados = {'ultimo': 0}
    if os.path.exists(caminho):
        try:
            with open(caminho, encoding='utf-8') as f:
                dados = json.load(f)
        except Exception:
            pass
    dados['ultimo'] = int(dados.get('ultimo', 0)) + 1
    with open(caminho, 'w', encoding='utf-8') as f:
        json.dump(dados, f)
    return dados['ultimo']


# ----------------------------------------------------------------------------
# GERACAO DO ROMANEIO EM PDF
# ----------------------------------------------------------------------------
def para_data(txt):
    """Converte 'dd/mm/aaaa ...' em date; devolve None se nao der."""
    if not txt:
        return None
    m = re.search(r'(\d{2})/(\d{2})/(\d{4})', str(txt))
    if not m:
        return None
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def fmt(v, casas=2):
    if v is None:
        return '-'
    s = f"{v:,.{casas}f}"
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


EST_ROTULO = ParagraphStyle('rot', fontName='Helvetica-Bold', fontSize=7.5, leading=9)
EST_VALOR = ParagraphStyle('val', fontName='Helvetica', fontSize=8, leading=9.5)
EST_CEL = ParagraphStyle('cel', fontName='Helvetica', fontSize=7, leading=8.2, alignment=TA_LEFT)
EST_CAB = ParagraphStyle('cab', fontName='Helvetica-Bold', fontSize=7, leading=8.2)
EST_TIT = ParagraphStyle('tit', fontName='Helvetica-Bold', fontSize=13, leading=15)
EST_SUB = ParagraphStyle('sub', fontName='Helvetica-Bold', fontSize=9, leading=11)
EST_SEC = ParagraphStyle('sec', fontName='Helvetica-Bold', fontSize=8, leading=10)


def _bloco_nota(d):
    fluxo = []
    fluxo.append(CondPageBreak(70 * mm))
    cab = [
        [Paragraph('CNPJ:', EST_ROTULO), Paragraph(d.get('cnpj') or '-', EST_VALOR),
         Paragraph('Nota Fiscal:', EST_ROTULO),
         Paragraph(f"{d.get('nota') or '-'} / {d.get('serie') or '-'}", EST_VALOR)],
        [Paragraph('Cliente:', EST_ROTULO), Paragraph(d.get('cliente') or '-', EST_VALOR),
         Paragraph('Data/Hora:', EST_ROTULO), Paragraph(d.get('emissao') or '-', EST_VALOR)],
        [Paragraph('Transportadora:', EST_ROTULO),
         Paragraph(d.get('transportadora') or '-', EST_VALOR),
         Paragraph('Volumes:', EST_ROTULO),
         Paragraph(fmt(d.get('volumes'), 0), EST_VALOR)],
    ]
    t = Table(cab, colWidths=[24 * mm, 78 * mm, 24 * mm, 54 * mm])
    t.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
    ]))
    fluxo.append(t)
    fluxo.append(Spacer(1, 3 * mm))
    fluxo.append(Paragraph('ITENS DA NOTA FISCAL', EST_SEC))
    fluxo.append(Spacer(1, 1 * mm))

    dados = [[Paragraph('ITEM', EST_CAB), Paragraph('CÓDIGO', EST_CAB),
              Paragraph('DESCRIÇÃO', EST_CAB), Paragraph('QTD', EST_CAB),
              Paragraph('UN', EST_CAB)]]
    for n, it in enumerate(d['itens'], 1):
        dados.append([Paragraph(f"{n:02d}", EST_CEL),
                      Paragraph(it['cod'], EST_CEL),
                      Paragraph(it['desc'], EST_CEL),
                      Paragraph(fmt(it['qtd']), EST_CEL),
                      Paragraph(it['unid'], EST_CEL)])
    ti = Table(dados, colWidths=[11 * mm, 24 * mm, 108 * mm, 24 * mm, 13 * mm],
               repeatRows=1)
    ti.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.4, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.88, 0.88, 0.88)),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (3, 1), (3, -1), 'RIGHT'),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (4, 0), (4, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    fluxo.append(ti)

    por_un = {}
    for it in d['itens']:
        por_un[it['unid']] = por_un.get(it['unid'], 0) + it['qtd']
    sep = '&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;'
    totais = sep.join(f"Total {u}: <b>{fmt(v)}</b>" for u, v in sorted(por_un.items()))
    resumo = sep.join([
        f"Total de itens: <b>{len(d['itens'])}</b>",
        totais,
        f"Peso Bruto NF: <b>{fmt(d.get('peso_bruto'), 3)} kg</b>",
        f"Volumes: <b>{fmt(d.get('volumes'), 0)}</b>"])
    fluxo.append(Spacer(1, 1.5 * mm))
    fluxo.append(Paragraph(resumo, EST_VALOR))
    fluxo.append(Spacer(1, 2 * mm))
    fluxo.append(HRFlowable(width='100%', thickness=1, color=colors.black))
    fluxo.append(Spacer(1, 3 * mm))
    return fluxo


def gerar_romaneio(notas, caminho_pdf, numero, cfg):
    doc = SimpleDocTemplate(caminho_pdf, pagesize=A4,
                            leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=14 * mm,
                            title=f"Romaneio {numero:06d}")
    fluxo = []
    topo = Table([[Paragraph(cfg.get('empresa', ''), EST_TIT),
                   Paragraph(f"Ro: <b>{numero:06d}</b>", EST_SUB)],
                  [Paragraph(cfg.get('titulo', 'ROMANEIO DE EXPEDICAO'), EST_SUB),
                   Paragraph(datetime.now().strftime('%d/%m/%Y %H:%M'), EST_VALOR)]],
                 colWidths=[140 * mm, 40 * mm])
    topo.setStyle(TableStyle([
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
    ]))
    fluxo.append(topo)
    fluxo.append(Spacer(1, 1 * mm))
    fluxo.append(HRFlowable(width='100%', thickness=1.2, color=colors.black))
    fluxo.append(Spacer(1, 3 * mm))

    for d in notas:
        fluxo.extend(_bloco_nota(d))

    tot_vol = sum((d.get('volumes') or 0) for d in notas)
    tot_peso = sum((d.get('peso_bruto') or 0) for d in notas)
    tot_itens = sum(len(d['itens']) for d in notas)
    fluxo.append(Paragraph(
        f"TOTAL GERAL: <b>{len(notas)}</b> nota(s)&nbsp;&nbsp;|&nbsp;&nbsp;"
        f"Itens: <b>{tot_itens}</b>&nbsp;&nbsp;|&nbsp;&nbsp;"
        f"Volumes: <b>{fmt(tot_vol, 0)}</b>&nbsp;&nbsp;|&nbsp;&nbsp;"
        f"Peso Bruto: <b>{fmt(tot_peso, 3)} kg</b>",
        EST_SUB))
    fluxo.append(Spacer(1, 12 * mm))
    ass = Table([['_' * 28, '_' * 28, '_' * 28],
                 ['Expedição', 'Conferência', 'Motorista']],
                colWidths=[60 * mm, 60 * mm, 60 * mm])
    ass.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 1), (-1, 1), 1),
    ]))
    fluxo.append(ass)

    def rodape(canvas, doc_):
        canvas.saveState()
        canvas.setFont('Helvetica', 6.5)
        canvas.drawString(12 * mm, 8 * mm,
                          f"Impresso em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - "
                          f"Romaneio {numero:06d}")
        canvas.drawRightString(A4[0] - 12 * mm, 8 * mm, f"Pagina {doc_.page}")
        canvas.restoreState()

    doc.build(fluxo, onFirstPage=rodape, onLaterPages=rodape)
    return caminho_pdf


# ----------------------------------------------------------------------------
# E-MAIL
# ----------------------------------------------------------------------------
def _contexto(notas, numero):
    return {'romaneio': f"{numero:06d}",
            'notas': ', '.join(str(d.get('nota') or '?') for d in notas),
            'volumes': fmt(sum((d.get('volumes') or 0) for d in notas), 0)}


def localizar_thunderbird(cfg=None):
    if cfg and cfg.get('thunderbird') and os.path.exists(cfg['thunderbird']):
        return cfg['thunderbird']
    achado = shutil.which('thunderbird.exe') or shutil.which('thunderbird')
    if achado:
        return achado
    try:
        import winreg
        for raiz in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                k = winreg.OpenKey(
                    raiz, r'SOFTWARE\Clients\Mail\Mozilla Thunderbird\shell\open\command')
                valor = winreg.QueryValueEx(k, '')[0].strip('"').split('"')[0]
                if os.path.exists(valor):
                    return valor
            except OSError:
                continue
    except Exception:
        pass
    for base in (os.environ.get('ProgramFiles'), os.environ.get('ProgramFiles(x86)')):
        if base:
            p = os.path.join(base, 'Mozilla Thunderbird', 'thunderbird.exe')
            if os.path.exists(p):
                return p
    return None


def _limpar_campo(txt):
    """O parametro -compose separa campos por virgula; virgula e aspa quebram."""
    return str(txt).replace(',', ';').replace("'", '')


def enviar_por_thunderbird(pdf, notas, numero, cfg):
    exe = localizar_thunderbird(cfg)
    if not exe:
        raise RuntimeError('Thunderbird nao encontrado. Informe o caminho do thunderbird.exe '
                           'em Configuracoes > Conta de envio.')
    e = cfg['email']
    dest = e.get('destinatarios') or []
    if not dest:
        raise RuntimeError('Nenhum destinatario cadastrado.')
    ctx = _contexto(notas, numero)
    campos = [f"to='{';'.join(dest)}'",
              f"subject='{_limpar_campo(e['assunto'].format(**ctx))}'",
              f"body='{_limpar_campo(e['corpo'].format(**ctx))}'",
              f"attachment='{os.path.abspath(pdf)}'"]
    subprocess.Popen([exe, '-compose', ','.join(campos)])
    return 'thunderbird'


def enviar_por_smtp(pdf, notas, numero, cfg):
    e = cfg['email']
    
    if not e.get('smtp'):
        raise RuntimeError('Servidor SMTP nao configurado.\nVa em Configuracoes > Conta de envio e preencha os dados.')
        
    if not e.get('destinatarios'):
        raise RuntimeError('Nenhum destinatario cadastrado.')
    ctx = _contexto(notas, numero)
    msg = EmailMessage()
    msg['Subject'] = e['assunto'].format(**ctx)
    msg['From'] = e.get('remetente') or e.get('usuario')
    msg['To'] = ', '.join(e['destinatarios'])
    msg.set_content(e['corpo'].format(**ctx))
    with open(pdf, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf',
                           filename=os.path.basename(pdf))
    seg = (e.get('seguranca') or 'starttls').lower()
    porta = int(e.get('porta') or 587)
    try:
        if seg == 'ssl':
            servidor = smtplib.SMTP_SSL(e['smtp'], porta, timeout=60)
        else:
            servidor = smtplib.SMTP(e['smtp'], porta, timeout=60)
        with servidor as s:
            if seg == 'starttls':
                s.starttls()
            if e.get('autenticar', True) and e.get('usuario'):
                s.login(e['usuario'], e['senha'])
            s.send_message(msg)
    except (TimeoutError, OSError) as ex:
        if 'timed out' in str(ex).lower() or isinstance(ex, TimeoutError):
            raise RuntimeError(
                f"Sem resposta de {e['smtp']}:{porta}.\n\n"
                'Isso e bloqueio de rede ou servidor/porta errados - nao e senha.\n'
                'Confira no Thunderbird: Configuracoes da conta > Servidor de saida (SMTP), '
                'e copie servidor, porta e metodo de seguranca para ca.') from ex
        raise
    return 'smtp'


def enviar_email(pdf, notas, numero, cfg):
    modo = (cfg.get('email', {}).get('modo_envio') or 'smtp').lower()
    if modo == 'thunderbird':
        return enviar_por_thunderbird(pdf, notas, numero, cfg)
    return enviar_por_smtp(pdf, notas, numero, cfg)


def listar_impressoras():
    try:
        import win32print
        return [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    except Exception:
        return []


def impressora_padrao():
    try:
        import win32print
        return win32print.GetDefaultPrinter()
    except Exception:
        return ''


def impressao_nativa_disponivel():
    try:
        import win32ui  # noqa: F401
        import win32con  # noqa: F401
        import pypdfium2  # noqa: F401
        from PIL import ImageWin  # noqa: F401
        return True
    except Exception:
        return False


def abrir_preferencias_impressora(nome):
    if not nome:
        return
    subprocess.Popen(['rundll32', 'printui.dll,PrintUIEntry', '/e', f'/n{nome}'])


def imprimir_gdi(arquivos, impressora, copias=1, dpi=200, progresso=None):
    """Renderiza cada pagina e desenha no contexto da impressora (GDI)."""
    import win32ui
    import win32con
    import pypdfium2 as pdfium
    from PIL import ImageWin

    copias = max(1, int(copias))
    escala = float(dpi) / 72.0
    for arq in arquivos:
        doc = pdfium.PdfDocument(arq)
        try:
            total = len(doc)
            for c in range(copias):
                dc = win32ui.CreateDC()
                dc.CreatePrinterDC(impressora)
                larg = dc.GetDeviceCaps(win32con.HORZRES)
                alt = dc.GetDeviceCaps(win32con.VERTRES)
                dc.StartDoc(os.path.basename(arq))
                try:
                    for i in range(total):
                        if progresso:
                            progresso(f'{os.path.basename(arq)} - pagina {i + 1}/{total}'
                                      f' (copia {c + 1}/{copias})')
                        dc.StartPage()
                        img = doc[i].render(scale=escala).to_pil().convert('RGB')
                        prop = min(larg / img.width, alt / img.height)
                        lw, lh = int(img.width * prop), int(img.height * prop)
                        x, y = (larg - lw) // 2, (alt - lh) // 2
                        ImageWin.Dib(img).draw(dc.GetHandleOutput(), (x, y, x + lw, y + lh))
                        dc.EndPage()
                        del img
                    dc.EndDoc()
                finally:
                    try:
                        dc.DeleteDC()
                    except Exception:
                        pass
        finally:
            doc.close()
    return 'nativo'


def imprimir_fallback(arquivos, copias=1):
    if not hasattr(os, 'startfile'):
        raise RuntimeError('Impressao nao suportada neste sistema.')
    for arq in arquivos:
        for _ in range(max(1, int(copias))):
            os.startfile(arq, 'print')
            time.sleep(1.5)
    return 'fallback'


# ----------------------------------------------------------------------------
# INTERFACE
# ----------------------------------------------------------------------------
class JanelaImpressao(tk.Toplevel):
    """Caixa de impressao: impressora, copias e qualidade."""

    def __init__(self, pai, arquivos, titulo='Imprimir'):
        super().__init__(pai)
        self.pai = pai
        self.arquivos = arquivos
        self.title(titulo)
        self.transient(pai)
        self.grab_set()
        self.resizable(False, False)
        f = ttk.Frame(self, padding=12)
        f.pack(fill='both', expand=True)

        n_pag = self._contar_paginas()
        ttk.Label(f, text=f'{len(arquivos)} arquivo(s) - {n_pag} pagina(s)',
                  font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, columnspan=3,
                                                     sticky='w', pady=(0, 8))

        ttk.Label(f, text='Impressora:').grid(row=1, column=0, sticky='w', pady=4)
        lista = listar_impressoras()
        padrao = pai.cfg.get('impressora') or impressora_padrao()
        self.var_imp = tk.StringVar(value=padrao if padrao else (lista[0] if lista else ''))
        cb = ttk.Combobox(f, textvariable=self.var_imp, values=lista, width=42)
        cb.grid(row=1, column=1, sticky='w', pady=4)
        ttk.Button(f, text='Propriedades',
                   command=lambda: abrir_preferencias_impressora(self.var_imp.get())
                   ).grid(row=1, column=2, padx=4)

        ttk.Label(f, text='Copias:').grid(row=2, column=0, sticky='w', pady=4)
        self.var_cop = tk.IntVar(value=int(pai.var_copias.get()))
        ttk.Spinbox(f, from_=1, to=50, width=6, textvariable=self.var_cop
                    ).grid(row=2, column=1, sticky='w', pady=4)

        ttk.Label(f, text='Qualidade:').grid(row=3, column=0, sticky='w', pady=4)
        self.var_dpi = tk.StringVar(value=str(pai.cfg.get('dpi_impressao', 200)))
        ttk.Combobox(f, textvariable=self.var_dpi, width=14, state='readonly',
                     values=['150 (rapida)', '200 (normal)', '300 (alta)']
                     ).grid(row=3, column=1, sticky='w', pady=4)
        self.var_dpi.set({'150': '150 (rapida)', '300': '300 (alta)'}.get(
            str(pai.cfg.get('dpi_impressao', 200)), '200 (normal)'))

        self.var_msg = tk.StringVar(value='')
        ttk.Label(f, textvariable=self.var_msg, foreground='#205020',
                  wraplength=430).grid(row=4, column=0, columnspan=3, sticky='w', pady=(8, 0))
        if not impressao_nativa_disponivel():
            self.var_msg.set('pywin32/pypdfium2 nao instalados - sera usado o leitor de PDF '
                             'padrao do Windows (a impressora escolhida acima e ignorada). '
                             'Instale com: pip install pywin32 pypdfium2')

        rod = ttk.Frame(f)
        rod.grid(row=5, column=0, columnspan=3, sticky='e', pady=(14, 0))
        ttk.Button(rod, text='Cancelar', command=self.destroy).pack(side='right')
        ttk.Button(rod, text='Imprimir', command=self.imprimir).pack(side='right', padx=6)

    def _contar_paginas(self):
        try:
            import pypdfium2 as pdfium
            t = 0
            for a in self.arquivos:
                d = pdfium.PdfDocument(a)
                t += len(d)
                d.close()
            return t
        except Exception:
            return '?'

    def imprimir(self):
        dpi = int(self.var_dpi.get().split()[0])
        copias = int(self.var_cop.get())
        imp = self.var_imp.get().strip()
        self.pai.cfg['impressora'] = imp
        self.pai.cfg['dpi_impressao'] = dpi
        self.pai.salvar_config()
        arquivos = list(self.arquivos)
        nativo = impressao_nativa_disponivel()
        self.var_msg.set('Enviando para a impressora...')
        self.update_idletasks()

        def tarefa():
            try:
                if nativo:
                    imprimir_gdi(arquivos, imp, copias, dpi,
                                 progresso=lambda m: self.pai.var_status.set(m))
                else:
                    imprimir_fallback(arquivos, copias)
                self.pai.after(0, lambda: self.pai.var_status.set(
                    f'Impressao concluida ({len(arquivos)} arquivo(s), {copias} copia(s)).'))
                self.pai.after(0, self.destroy)
            except Exception as ex:
                msg = str(ex)
                if 'CreatePrinterDC' in msg or 'Nome de impressora' in msg:
                    msg = (f'Nao foi possivel acessar a impressora "{imp}".\n'
                           'Verifique se ela esta ligada, conectada e com o nome correto.')
                self.pai.after(0, lambda: messagebox.showerror('Imprimir', msg, parent=self))
                self.pai.after(0, lambda: self.var_msg.set('Falha na impressao.'))

        threading.Thread(target=tarefa, daemon=True).start()


class JanelaRevisao(tk.Toplevel):
    """Confere o que vai ser enviado: lista das notas + previa do PDF."""

    def __init__(self, pai, pdf, notas, numero):
        super().__init__(pai)
        self.pai = pai
        self.pdf = pdf
        self.notas = notas
        self.numero = numero
        self.pagina = 0
        self.doc = None
        self.foto = None
        self.title(f'Conferir romaneio {numero:06d} antes de enviar')
        alt = min(720, self.winfo_screenheight() - 90)
        self.geometry(f'900x{alt}')
        self.minsize(760, 380)
        self.transient(pai)
        self.grab_set()

        # rodape primeiro e ancorado embaixo: nunca sai da tela
        rod = ttk.Frame(self, padding=8)
        rod.pack(side='bottom', fill='x')
        ttk.Separator(self, orient='horizontal').pack(side='bottom', fill='x')
        ttk.Button(rod, text='Cancelar', command=self.destroy).pack(side='right')
        self.bt_env = ttk.Button(rod, text='ENVIAR POR E-MAIL', command=self.enviar)
        self.bt_env.pack(side='right', padx=8)
        ttk.Button(rod, text='Imprimir romaneio',
                   command=self.imprimir_romaneio).pack(side='right', padx=4)
        ttk.Button(rod, text='Imprimir notas',
                   command=self.imprimir_notas).pack(side='right')
        ttk.Button(rod, text='Abrir PDF', command=self.abrir_externo).pack(side='left')
        ttk.Label(rod, text='  Confira antes de enviar.',
                  font=('Segoe UI', 9, 'bold')).pack(side='left')

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=8, pady=(8, 0))
        self._aba_resumo(nb)
        self._aba_previa(nb)

    # ------------------------------------------------------------ aba resumo
    def _aba_resumo(self, nb):
        f = ttk.Frame(nb, padding=10)
        nb.add(f, text='Notas do envio')

        ttk.Label(f, text=f'Romaneio {self.numero:06d}  -  '
                          f'{len(self.notas)} nota(s) neste envio',
                  font=('Segoe UI', 11, 'bold')).pack(anchor='w', pady=(0, 6))

        cx = ttk.Frame(f)
        cx.pack(fill='both', expand=True)
        cols = ('nota', 'cliente', 'emissao', 'itens', 'volumes', 'peso')
        tv = ttk.Treeview(cx, columns=cols, show='headings', height=13)
        larg = {'nota': 80, 'cliente': 300, 'emissao': 140, 'itens': 60,
                'volumes': 80, 'peso': 100}
        tit = {'nota': 'Nota', 'cliente': 'Cliente', 'emissao': 'Emissao',
               'itens': 'Itens', 'volumes': 'Volumes', 'peso': 'Peso Bruto'}
        for c in cols:
            tv.heading(c, text=tit[c])
            tv.column(c, width=larg[c], anchor='w' if c == 'cliente' else 'center')
        tv.tag_configure('semvol', background='#fff3cd')
        for d in self.notas:
            tv.insert('', 'end', tags=('semvol',) if not d.get('volumes') else (),
                      values=(d.get('nota') or '?', (d.get('cliente') or '-')[:55],
                              d.get('emissao') or '-', len(d.get('itens') or []),
                              fmt(d.get('volumes'), 0), fmt(d.get('peso_bruto'), 3)))
        tv.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(cx, orient='vertical', command=tv.yview)
        sb.pack(side='right', fill='y')
        tv.configure(yscrollcommand=sb.set)

        vol = sum((d.get('volumes') or 0) for d in self.notas)
        peso = sum((d.get('peso_bruto') or 0) for d in self.notas)
        itens = sum(len(d.get('itens') or []) for d in self.notas)
        ttk.Label(f, text=f'TOTAL:  {len(self.notas)} nota(s)   |   {itens} itens   |   '
                          f'{fmt(vol, 0)} volumes   |   {fmt(peso, 3)} kg',
                  font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(8, 4))

        sem_vol = [d for d in self.notas if not d.get('volumes')]
        if sem_vol:
            ttk.Label(f, foreground='#8a5a00', wraplength=820,
                      text='Atencao: nota(s) sem volumes informados - '
                           + ', '.join(str(d.get('nota')) for d in sem_vol)).pack(anchor='w')

        dest = self.pai.cfg['email'].get('destinatarios', [])
        ttk.Separator(f, orient='horizontal').pack(fill='x', pady=8)
        ttk.Label(f, text='Sera enviado para:', font=('Segoe UI', 9, 'bold')).pack(anchor='w')
        ttk.Label(f, text=(', '.join(dest) if dest else
                           'Nenhum destinatario cadastrado - abra Configuracoes > Destinatarios.'),
                  foreground=('#205020' if dest else '#a02020'),
                  wraplength=820).pack(anchor='w')

    # ------------------------------------------------------------ aba previa
    def _aba_previa(self, nb):
        f = ttk.Frame(nb, padding=6)
        nb.add(f, text='Previa do PDF')
        barra = ttk.Frame(f)
        barra.pack(fill='x', pady=(0, 4))
        ttk.Button(barra, text='<', width=3, command=lambda: self.ir(-1)).pack(side='left')
        self.var_pg = tk.StringVar(value='carregando...')
        ttk.Label(barra, textvariable=self.var_pg, width=18,
                  anchor='center').pack(side='left', padx=4)
        ttk.Button(barra, text='>', width=3, command=lambda: self.ir(1)).pack(side='left')
        quadro = ttk.Frame(f)
        quadro.pack(fill='both', expand=True)
        self.canvas = tk.Canvas(quadro, background='#8f8f8f', highlightthickness=0)
        self.canvas.pack(side='left', fill='both', expand=True)
        sv = ttk.Scrollbar(quadro, orient='vertical', command=self.canvas.yview)
        sv.pack(side='right', fill='y')
        sh = ttk.Scrollbar(f, orient='horizontal', command=self.canvas.xview)
        sh.pack(fill='x')
        self.canvas.configure(yscrollcommand=sv.set, xscrollcommand=sh.set)
        self.after(80, self.carregar)

    def carregar(self):
        try:
            import pypdfium2 as pdfium
            self.doc = pdfium.PdfDocument(self.pdf)
            self.mostrar()
        except Exception:
            self.canvas.create_text(16, 16, anchor='nw', fill='white',
                                    text='Previa indisponivel (pypdfium2 nao instalado).\n'
                                         'Use "Abrir no leitor de PDF".')
            self.var_pg.set('previa off')

    def mostrar(self):
        if not self.doc:
            return
        from PIL import ImageTk
        self.pagina = max(0, min(self.pagina, len(self.doc) - 1))
        img = self.doc[self.pagina].render(scale=105 / 72).to_pil().convert('RGB')
        self.foto = ImageTk.PhotoImage(img)
        self.canvas.delete('all')
        self.canvas.create_image(8, 8, anchor='nw', image=self.foto)
        self.canvas.configure(scrollregion=(0, 0, img.width + 16, img.height + 16))
        self.var_pg.set(f'Pagina {self.pagina + 1} de {len(self.doc)}')

    def ir(self, passo):
        self.pagina += passo
        self.mostrar()

    def abrir_externo(self):
        try:
            os.startfile(self.pdf)
        except Exception:
            messagebox.showinfo('Arquivo', self.pdf, parent=self)

    def imprimir_romaneio(self):
        JanelaImpressao(self.pai, [self.pdf], 'Imprimir romaneio')

    def imprimir_notas(self):
        arqs = [d['caminho'] for d in self.notas if os.path.exists(d.get('caminho') or '')]
        if not arqs:
            messagebox.showinfo('Imprimir', 'Nenhum PDF de nota localizado.', parent=self)
            return
        JanelaImpressao(self.pai, arqs, 'Imprimir notas do romaneio')

    # ------------------------------------------------------------ envio
    def enviar(self):
        dest = self.pai.cfg['email'].get('destinatarios', [])
        if not dest:
            messagebox.showwarning('E-mail', 'Nenhum destinatario cadastrado.\n'
                                             'Abra Configuracoes > Destinatarios.', parent=self)
            return
        vol = fmt(sum((d.get('volumes') or 0) for d in self.notas), 0)
        lista = ', '.join(str(d.get('nota') or '?') for d in self.notas)
        if not messagebox.askyesno(
                'Confirmar envio',
                f'Esta tudo ok, pode enviar?\n\n'
                f'Romaneio: {self.numero:06d}\n'
                f'Notas ({len(self.notas)}): {lista}\n'
                f'Volumes: {vol}\n\n'
                f'Para: {", ".join(dest)}', parent=self):
            return
        self.bt_env.configure(state='disabled')
        modo = (self.pai.cfg['email'].get('modo_envio') or 'smtp').lower()
        self.pai.var_status.set('Abrindo o Thunderbird...' if modo == 'thunderbird'
                                else 'Enviando e-mail...')

        def tarefa():
            try:
                usado = enviar_email(self.pdf, self.notas, self.numero, self.pai.cfg)
                if usado == 'thunderbird':
                    self.pai.after(0, lambda: self.pai.var_status.set(
                        'Mensagem aberta no Thunderbird - clique em Enviar por la.'))
                    self.pai.after(0, lambda: messagebox.showinfo(
                        'Thunderbird', 'A mensagem foi aberta no Thunderbird com o romaneio '
                                       'anexado.\nConfira e clique em Enviar.', parent=self.pai))
                else:
                    self.pai.after(0, lambda: self.pai.var_status.set(
                        f'Romaneio {self.numero:06d} enviado para {len(dest)} destinatario(s).'))
                    self.pai.after(0, lambda: messagebox.showinfo(
                        'E-mail', 'Romaneio enviado com sucesso.', parent=self.pai))
                self.pai.after(0, self.destroy)
            except smtplib.SMTPAuthenticationError:
                self.pai.after(0, lambda: messagebox.showerror(
                    'E-mail', 'Usuario ou senha recusados pelo servidor.\n'
                              'No Gmail use senha de app de 16 digitos.', parent=self))
                self.pai.after(0, lambda: self.bt_env.configure(state='normal'))
            except Exception as ex:
                # ex some ao sair do except: guardar no proprio lambda.
                self.pai.after(0, lambda ex=ex: messagebox.showerror(
                    'E-mail', str(ex), parent=self))
                self.pai.after(0, lambda: self.bt_env.configure(state='normal'))

        threading.Thread(target=tarefa, daemon=True).start()

    def destroy(self):
        try:
            if self.doc:
                self.doc.close()
        except Exception:
            pass
        super().destroy()


class JanelaConfig(tk.Toplevel):
    """Cadastro da conta de envio, destinatarios e impressao."""

    def __init__(self, pai):
        super().__init__(pai)
        self.pai = pai
        self.cfg = json.loads(json.dumps(pai.cfg))
        self.title('Configuracoes')
        alt = min(600, self.winfo_screenheight() - 90)
        self.geometry(f'660x{alt}')
        self.transient(pai)
        self.grab_set()

        rod = ttk.Frame(self, padding=8)
        rod.pack(side='bottom', fill='x')
        ttk.Separator(self, orient='horizontal').pack(side='bottom', fill='x')
        ttk.Button(rod, text='Testar envio', command=self.testar).pack(side='left')
        ttk.Button(rod, text='Cancelar', command=self.destroy).pack(side='right')
        ttk.Button(rod, text='Salvar', command=self.salvar).pack(side='right', padx=6)

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=8, pady=8)
        self._aba_conta(nb)
        self._aba_destinatarios(nb)
        self._aba_servidor(nb)
        self._aba_impressao(nb)

    # ---------------------------------------------------------------- abas
    def _aba_conta(self, nb):
        f = ttk.Frame(nb, padding=10)
        nb.add(f, text='Conta de envio')
        e = self.cfg['email']
        self.v = {}

        ttk.Label(f, text='Como enviar:', font=('Segoe UI', 9, 'bold')).grid(
            row=0, column=0, sticky='w', pady=(0, 4))
        self.var_modo = tk.StringVar(value=(e.get('modo_envio') or 'smtp'))
        lm = ttk.Frame(f)
        lm.grid(row=0, column=1, sticky='w')
        ttk.Radiobutton(lm, text='SMTP direto (envia sozinho)', value='smtp',
                        variable=self.var_modo, command=self._alternar_modo).pack(side='left')
        ttk.Radiobutton(lm, text='Abrir no Thunderbird', value='thunderbird',
                        variable=self.var_modo,
                        command=self._alternar_modo).pack(side='left', padx=12)

        tb = localizar_thunderbird(self.cfg)
        ttk.Label(f, text=(f'Thunderbird: {tb}' if tb else
                           'Thunderbird nao localizado - informe o caminho abaixo.'),
                  foreground=('#206020' if tb else '#a02020'),
                  wraplength=430).grid(row=1, column=1, sticky='w', pady=(2, 6))
        ttk.Label(f, text='thunderbird.exe:').grid(row=2, column=0, sticky='w', pady=3)
        self.var_tb = tk.StringVar(value=self.cfg.get('thunderbird', '') or (tb or ''))
        ttk.Entry(f, textvariable=self.var_tb, width=40).grid(row=2, column=1, sticky='w')
        ttk.Button(f, text='...', width=3, command=self.buscar_tb).grid(row=2, column=2)

        ttk.Separator(f, orient='horizontal').grid(row=3, column=0, columnspan=3,
                                                   sticky='ew', pady=8)

        self.campos_smtp = []
        linha = 4
        for rot, chave, larg in [('Servidor SMTP:', 'smtp', 34), ('Porta:', 'porta', 8),
                                 ('Usuario (login):', 'usuario', 34), ('Senha:', 'senha', 34),
                                 ('Remetente:', 'remetente', 34)]:
            lb = ttk.Label(f, text=rot)
            lb.grid(row=linha, column=0, sticky='w', pady=3)
            var = tk.StringVar(value=str(e.get(chave, '')))
            en = ttk.Entry(f, textvariable=var, width=larg,
                           show='*' if chave == 'senha' else '')
            en.grid(row=linha, column=1, sticky='w', pady=3)
            self.v[chave] = var
            self.campos_smtp += [lb, en]
            linha += 1

        lb = ttk.Label(f, text='Seguranca:')
        lb.grid(row=linha, column=0, sticky='w', pady=3)
        self.var_seg = tk.StringVar(value=(e.get('seguranca') or 'starttls'))
        cb = ttk.Combobox(f, textvariable=self.var_seg, width=22, state='readonly',
                          values=['starttls', 'ssl', 'nenhuma'])
        cb.grid(row=linha, column=1, sticky='w', pady=3)
        self.campos_smtp += [lb, cb]
        linha += 1

        self.var_auth = tk.BooleanVar(value=bool(e.get('autenticar', True)))
        ck = ttk.Checkbutton(f, text='Servidor exige usuario e senha', variable=self.var_auth)
        ck.grid(row=linha, column=1, sticky='w', pady=3)
        self.campos_smtp.append(ck)
        linha += 1

        ttk.Label(f, text='Copie servidor, porta e seguranca do Thunderbird: '
                          'Configuracoes da conta > Servidor de saida (SMTP).',
                  foreground='#8a5a00', wraplength=430).grid(row=linha, column=1, sticky='w',
                                                             pady=(4, 8))
        linha += 1

        ttk.Label(f, text='Assunto:').grid(row=linha, column=0, sticky='w', pady=3)
        var = tk.StringVar(value=str(e.get('assunto', '')))
        ttk.Entry(f, textvariable=var, width=44).grid(row=linha, column=1, sticky='w')
        self.v['assunto'] = var
        linha += 1
        ttk.Label(f, text='Corpo:').grid(row=linha, column=0, sticky='nw', pady=3)
        self.txt_corpo = tk.Text(f, width=48, height=4, wrap='word')
        self.txt_corpo.grid(row=linha, column=1, sticky='w', pady=3)
        self.txt_corpo.insert('1.0', e.get('corpo', ''))
        linha += 1
        ttk.Label(f, text='Variaveis: {romaneio}  {notas}  {volumes}',
                  foreground='#555').grid(row=linha, column=1, sticky='w')
        self._alternar_modo()

    def _alternar_modo(self):
        estado = 'disabled' if self.var_modo.get() == 'thunderbird' else 'normal'
        for w in getattr(self, 'campos_smtp', []):
            try:
                w.configure(state=estado if not isinstance(w, ttk.Label) else 'normal')
            except Exception:
                pass

    def buscar_tb(self):
        p = filedialog.askopenfilename(parent=self, title='Selecione o thunderbird.exe',
                                       filetypes=[('Executavel', '*.exe'), ('Todos', '*.*')])
        if p:
            self.var_tb.set(p)

    def _aba_destinatarios(self, nb):
        f = ttk.Frame(nb, padding=10)
        nb.add(f, text='Destinatarios')
        ttk.Label(f, text='E-mails que recebem o romaneio:').pack(anchor='w')
        cx = ttk.Frame(f)
        cx.pack(fill='both', expand=True, pady=4)
        self.lst = tk.Listbox(cx, height=12)
        self.lst.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(cx, orient='vertical', command=self.lst.yview)
        sb.pack(side='right', fill='y')
        self.lst.configure(yscrollcommand=sb.set)
        for d in self.cfg['email'].get('destinatarios', []):
            self.lst.insert('end', d)
        ln = ttk.Frame(f)
        ln.pack(fill='x', pady=4)
        self.var_novo = tk.StringVar()
        ent = ttk.Entry(ln, textvariable=self.var_novo, width=40)
        ent.pack(side='left')
        ent.bind('<Return>', lambda ev: self.add_email())
        ttk.Button(ln, text='Adicionar', command=self.add_email).pack(side='left', padx=4)
        ttk.Button(ln, text='Remover', command=self.rem_email).pack(side='left')
        self.lst.bind('<Double-1>', self.edit_email)
        ttk.Label(f, text='Duplo clique edita. Enter adiciona.',
                  foreground='#555').pack(anchor='w')

    def _aba_impressao(self, nb):
        f = ttk.Frame(nb, padding=10)
        nb.add(f, text='Impressao e pastas')
        ok = impressao_nativa_disponivel()
        ttk.Label(f, text=('Impressao nativa ativa (pywin32 + pypdfium2).' if ok else
                           'Impressao nativa indisponivel. Instale com:  '
                           'pip install pywin32 pypdfium2'),
                  foreground=('#206020' if ok else '#a02020'),
                  wraplength=430).grid(row=0, column=1, columnspan=2, sticky='w', pady=(0, 8))

        ttk.Label(f, text='Impressora padrao:').grid(row=2, column=0, sticky='w', pady=4)
        self.var_impressora = tk.StringVar(
            value=self.cfg.get('impressora', '') or impressora_padrao())
        cb = ttk.Combobox(f, textvariable=self.var_impressora, width=44,
                          values=listar_impressoras())
        cb.grid(row=2, column=1, sticky='w')
        ttk.Button(f, text='Propriedades',
                   command=lambda: abrir_preferencias_impressora(self.var_impressora.get())
                   ).grid(row=2, column=2, padx=4)

        ttk.Label(f, text='Pasta das notas:').grid(row=3, column=0, sticky='w', pady=4)
        self.var_pasta_notas = tk.StringVar(value=self.cfg.get('pasta_notas', ''))
        ttk.Entry(f, textvariable=self.var_pasta_notas, width=46).grid(row=3, column=1,
                                                                       sticky='w')
        ttk.Button(f, text='...', width=3,
                   command=lambda: self._pasta(self.var_pasta_notas)).grid(row=3, column=2)

        ttk.Label(f, text='Pasta dos romaneios:').grid(row=4, column=0, sticky='w', pady=4)
        self.var_pasta_saida = tk.StringVar(value=self.cfg.get('pasta_saida', ''))
        ttk.Entry(f, textvariable=self.var_pasta_saida, width=46).grid(row=4, column=1,
                                                                       sticky='w')
        ttk.Button(f, text='...', width=3,
                   command=lambda: self._pasta(self.var_pasta_saida)).grid(row=4, column=2)

        ttk.Label(f, text='Contador do romaneio:').grid(row=5, column=0, sticky='w', pady=4)
        self.var_contador = tk.StringVar(value=self.cfg.get('arquivo_contador', ''))
        ttk.Entry(f, textvariable=self.var_contador, width=46).grid(row=5, column=1, sticky='w')
        ttk.Label(f, text='Se mais de um PC usar o programa, aponte para uma pasta de rede.',
                  foreground='#8a5a00', wraplength=430).grid(row=6, column=1, columnspan=2,
                                                             sticky='w', pady=(2, 0))
        if (self.cfg.get('servidor') or {}).get('ativo'):
            ttk.Label(f, text='A pasta das notas esta sendo preenchida pela aba '
                              '"Servidor das notas".',
                      foreground='#205080', wraplength=430).grid(
                row=7, column=1, columnspan=2, sticky='w', pady=(10, 0))

    def _aba_servidor(self, nb):
        f = ttk.Frame(nb, padding=10)
        nb.add(f, text='Servidor das notas')
        s = self.cfg.get('servidor') or {}

        ttk.Label(f, text='De onde o programa le os PDFs das notas fiscais',
                  font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, columnspan=3,
                                                     sticky='w', pady=(0, 6))
        self.var_srv_ativo = tk.BooleanVar(value=bool(s.get('ativo')))
        ttk.Checkbutton(f, text='Ler as notas de um servidor da rede',
                        variable=self.var_srv_ativo,
                        command=self._alternar_servidor).grid(
            row=1, column=0, columnspan=3, sticky='w', pady=(0, 6))

        self.campos_srv = []
        self.var_srv = {}
        linha = 2
        for rot, chave, larg in [('Servidor (IP ou nome):', 'host', 26),
                                 ('Pasta no servidor:', 'pasta_remota', 42),
                                 ('Compartilhamento:', 'compartilhamento', 12),
                                 ('Usuario:', 'usuario', 26),
                                 ('Dominio (opcional):', 'dominio', 26)]:
            lb = ttk.Label(f, text=rot)
            lb.grid(row=linha, column=0, sticky='w', pady=3)
            var = tk.StringVar(value=str(s.get(chave, '')))
            en = ttk.Entry(f, textvariable=var, width=larg)
            en.grid(row=linha, column=1, columnspan=2, sticky='w', pady=3)
            var.trace_add('write', lambda *a: self._mostrar_caminho_rede())
            self.var_srv[chave] = var
            self.campos_srv += [lb, en]
            linha += 1

        lb = ttk.Label(f, text='Senha:')
        lb.grid(row=linha, column=0, sticky='w', pady=3)
        self.var_srv_senha = tk.StringVar(value=senha_rede(s))
        en = ttk.Entry(f, textvariable=self.var_srv_senha, width=26, show='*')
        en.grid(row=linha, column=1, sticky='w', pady=3)
        bt = ttk.Button(f, text='Esquecer', width=10, command=self.esquecer_senha_rede)
        bt.grid(row=linha, column=2, sticky='w', padx=4)
        self.campos_srv += [lb, en, bt]
        linha += 1

        self.var_srv_win = tk.BooleanVar(value=bool(s.get('usar_credenciais_windows')))
        ck = ttk.Checkbutton(f, text='Entrar com o usuario do Windows deste PC '
                                     '(nao usar o usuario e a senha acima)',
                             variable=self.var_srv_win)
        ck.grid(row=linha, column=0, columnspan=3, sticky='w', pady=3)
        self.campos_srv.append(ck)
        linha += 1

        self.var_srv_caminho = tk.StringVar()
        ttk.Label(f, textvariable=self.var_srv_caminho, foreground='#205080',
                  wraplength=470).grid(row=linha, column=0, columnspan=3,
                                       sticky='w', pady=(8, 2))
        linha += 1
        bt = ttk.Button(f, text='Testar conexao', command=self.testar_servidor)
        bt.grid(row=linha, column=0, sticky='w', pady=6)
        self.campos_srv.append(bt)
        linha += 1

        ttk.Label(f, text='A senha fica gravada so neste computador, no arquivo '
                          'credenciais_rede.json, cifrada pela conta do Windows '
                          '(quando o pywin32 esta instalado). '
                          'Ela nunca e escrita no config.json.\n\n'
                          'A pasta pode ser informada como o servidor a enxerga '
                          '(C:\\Users\\fulano\\Desktop\\NOTAS): o programa converte '
                          'para \\\\servidor\\C$\\Users\\fulano\\Desktop\\NOTAS.\n'
                          'O compartilhamento C$ so aceita conta administradora do '
                          'servidor. Se a conta nao for administradora, compartilhe a '
                          'pasta NOTAS no servidor, escreva o nome do compartilhamento '
                          'no campo Compartilhamento e deixe a pasta em branco.',
                  foreground='#8a5a00', wraplength=470,
                  justify='left').grid(row=linha, column=0, columnspan=3, sticky='w')
        self._mostrar_caminho_rede()
        self._alternar_servidor()

    def _alternar_servidor(self):
        estado = 'normal' if self.var_srv_ativo.get() else 'disabled'
        for w in getattr(self, 'campos_srv', []):
            try:
                w.configure(state=estado)
            except Exception:
                pass

    def _mostrar_caminho_rede(self):
        alvo = montar_caminho_rede({k: v.get() for k, v in self.var_srv.items()})
        self.var_srv_caminho.set(f'Pasta que sera lida:   {alvo}' if alvo else
                                 'Preencha o servidor e a pasta.')

    def _coletar_servidor(self):
        s = dict(self.cfg.get('servidor') or {})
        s.update({k: v.get().strip() for k, v in self.var_srv.items()})
        s['ativo'] = bool(self.var_srv_ativo.get())
        s['usar_credenciais_windows'] = bool(self.var_srv_win.get())
        s.pop('senha', None)
        return s

    def esquecer_senha_rede(self):
        salvar_senha_rede(self.var_srv['host'].get(), self.var_srv['usuario'].get(), '')
        self.var_srv_senha.set('')
        messagebox.showinfo('Servidor das notas',
                            'Senha apagada deste computador.', parent=self)

    def testar_servidor(self):
        srv = self._coletar_servidor()
        alvo = montar_caminho_rede(srv)
        if not alvo:
            messagebox.showwarning('Servidor das notas',
                                   'Informe o servidor e a pasta.', parent=self)
            return
        salvar_senha_rede(srv.get('host'), srv.get('usuario'), self.var_srv_senha.get())
        self.config(cursor='watch')
        self.update_idletasks()
        erro = None
        try:
            _destino, erro = conectar_rede(alvo, srv, reconectar=True)
        except Exception as ex:
            erro = str(ex)
        finally:
            self.config(cursor='')
        if erro == SENHA_AUSENTE:
            erro = 'Informe a senha do usuario do servidor.'
        if erro:
            messagebox.showerror('Servidor das notas', erro, parent=self)
            return
        pdfs = glob.glob(os.path.join(alvo, '*.pdf')) + glob.glob(os.path.join(alvo, '*.PDF'))
        messagebox.showinfo('Servidor das notas',
                            f'Conexao OK.\n\n{alvo}\n\n{len(pdfs)} PDF(s) na pasta.',
                            parent=self)

    # ------------------------------------------------------------- apoio
    def _pasta(self, var):
        p = filedialog.askdirectory(parent=self)
        if p:
            var.set(p)

    def add_email(self):
        v = self.var_novo.get().strip()
        if not v:
            return
        if '@' not in v or '.' not in v.split('@')[-1]:
            messagebox.showwarning('E-mail', 'Endereco invalido.', parent=self)
            return
        if v in self.lst.get(0, 'end'):
            messagebox.showinfo('E-mail', 'Ja cadastrado.', parent=self)
            return
        self.lst.insert('end', v)
        self.var_novo.set('')

    def rem_email(self):
        for i in reversed(self.lst.curselection()):
            self.lst.delete(i)

    def edit_email(self, _ev=None):
        sel = self.lst.curselection()
        if not sel:
            return
        atual = self.lst.get(sel[0])
        novo = simpledialog.askstring('Editar', 'E-mail:', initialvalue=atual, parent=self)
        if novo:
            self.lst.delete(sel[0])
            self.lst.insert(sel[0], novo.strip())

    def _coletar(self):
        c = self.cfg
        c['impressora'] = self.var_impressora.get().strip()
        c['pasta_notas'] = self.var_pasta_notas.get().strip()
        c['servidor'] = srv = self._coletar_servidor()
        if srv.get('ativo'):
            c['pasta_notas'] = montar_caminho_rede(srv) or c['pasta_notas']
        c['pasta_saida'] = self.var_pasta_saida.get().strip()
        c['arquivo_contador'] = self.var_contador.get().strip() or 'contador_romaneio.json'
        c['thunderbird'] = self.var_tb.get().strip()
        e = c['email']
        for k, var in self.v.items():
            e[k] = var.get().strip()
        e['modo_envio'] = self.var_modo.get()
        e['seguranca'] = self.var_seg.get()
        e['autenticar'] = bool(self.var_auth.get())
        try:
            e['porta'] = int(e['porta'])
        except ValueError:
            e['porta'] = 587
        e['corpo'] = self.txt_corpo.get('1.0', 'end').rstrip('\n')
        e['destinatarios'] = list(self.lst.get(0, 'end'))
        return c

    def testar(self):
        cfg = self._coletar()
        e = cfg['email']
        if e.get('modo_envio') == 'thunderbird':
            if not localizar_thunderbird(cfg):
                messagebox.showerror('Teste', 'Thunderbird nao encontrado.', parent=self)
                return
            messagebox.showinfo('Teste', 'Modo Thunderbird: a mensagem sera aberta com o '
                                         'romaneio anexado e voce clica em Enviar por la.\n'
                                         'Nada a testar aqui.', parent=self)
            return
        if not e.get('smtp'):
            messagebox.showwarning('Teste', 'Informe o servidor SMTP.', parent=self)
            return
        if e.get('autenticar', True) and not (e.get('usuario') and e.get('senha')):
            messagebox.showwarning('Teste', 'Preencha usuario e senha.', parent=self)
            return
        destino = e.get('remetente') or e['usuario']
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            msg = EmailMessage()
            msg['Subject'] = 'Teste - Romaneio de Expedicao'
            msg['From'] = destino
            msg['To'] = destino
            msg.set_content('Teste de configuracao do programa de romaneio. '
                            'Se voce recebeu isto, o envio esta funcionando.')
            seg = (e.get('seguranca') or 'starttls').lower()
            porta = int(e.get('porta') or 587)
            conexao = (smtplib.SMTP_SSL(e['smtp'], porta, timeout=45) if seg == 'ssl'
                       else smtplib.SMTP(e['smtp'], porta, timeout=45))
            with conexao as s:
                if seg == 'starttls':
                    s.starttls()
                if e.get('autenticar', True) and e.get('usuario'):
                    s.login(e['usuario'], e['senha'])
                s.send_message(msg)
            messagebox.showinfo('Teste', f'Enviado para {destino}.\nConfira a caixa de entrada.',
                                parent=self)
        except smtplib.SMTPAuthenticationError:
            messagebox.showerror('Teste', 'Usuario ou senha recusados pelo servidor.\n'
                                          'No Gmail, use senha de app de 16 digitos.',
                                 parent=self)
        except (TimeoutError, OSError) as ex:
            messagebox.showerror(
                'Teste',
                f"Sem resposta de {e.get('smtp')}:{e.get('porta')}.\n\n"
                'Nao e senha - e rede ou servidor/porta errados.\n\n'
                'Abra o Thunderbird > Configuracoes da conta > Servidor de saida (SMTP) '
                'e copie servidor, porta e metodo de seguranca exatamente como estao la.\n\n'
                f'Detalhe: {ex}', parent=self)
        except Exception as ex:
            messagebox.showerror('Teste', f'Falha: {ex}', parent=self)
        finally:
            self.config(cursor='')

    def salvar(self):
        cfg = self._coletar()
        srv = cfg.get('servidor') or {}
        if not salvar_senha_rede(srv.get('host'), srv.get('usuario'),
                                 self.var_srv_senha.get()):
            messagebox.showwarning('Servidor das notas',
                                   'Nao consegui gravar a senha em '
                                   f'{CRED_PATH}.\nEla sera pedida na proxima leitura.',
                                   parent=self)
        try:
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
        except Exception as ex:
            messagebox.showerror('Salvar', str(ex), parent=self)
            return
        self.pai.cfg = cfg
        self.pai.var_pasta.set(cfg.get('pasta_notas', ''))
        self.pai.atualizar_aviso_impressao()
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Romaneio de Expedicao - Marfim')
        self.geometry(f'1150x{min(680, self.winfo_screenheight() - 80)}')
        self.minsize(900, 420)
        self.cfg = carregar_config()
        migrar_senha_config(self.cfg)
        self.notas = []
        self.marcadas = set()
        self.ultimo_romaneio = None
        self.ultimo_numero = None
        self.ultimas_notas = []
        self.img_off = self._icone(False)
        self.img_on = self._icone(True)
        self._montar()
        self.after(150, self.atualizar_aviso_impressao)
        if self.pasta_configurada():
            self.after(300, self.atualizar)

    @staticmethod
    def _icone(marcado):
        """Desenha um checkbox 16x16 em runtime (sem arquivo externo)."""
        img = tk.PhotoImage(width=16, height=16)
        img.blank()
        borda = '#4a4a4a'
        img.put('#ffffff', to=(3, 3, 14, 14))
        img.put(borda, to=(2, 2, 15, 3))
        img.put(borda, to=(2, 13, 15, 14))
        img.put(borda, to=(2, 2, 3, 14))
        img.put(borda, to=(14, 2, 15, 14))
        if marcado:
            for x, y in ((4, 8), (5, 9), (6, 10), (7, 9), (8, 8),
                         (9, 7), (10, 6), (11, 5)):
                img.put('#1a6fd4', to=(x, y, x + 2, y + 2))
        return img

    def _montar(self):
        topo = ttk.Frame(self, padding=6)
        topo.pack(fill='x')
        ttk.Label(topo, text='Pasta das notas:').pack(side='left')
        self.var_pasta = tk.StringVar(value=self.pasta_configurada())
        ttk.Entry(topo, textvariable=self.var_pasta, width=70).pack(side='left', padx=4)
        ttk.Button(topo, text='...', width=3, command=self.escolher_pasta).pack(side='left')
        ttk.Button(topo, text='Atualizar lista', command=self.atualizar).pack(side='left', padx=6)
        ttk.Label(topo, text='Filtro:').pack(side='left', padx=(16, 2))
        self.var_filtro = tk.StringVar()
        self.var_filtro.trace_add('write', lambda *a: self.preencher())
        ttk.Entry(topo, textvariable=self.var_filtro, width=18).pack(side='left')

        datas = ttk.Frame(self, padding=(6, 0, 6, 4))
        datas.pack(fill='x')
        hoje = date.today().strftime('%d/%m/%Y')
        ttk.Label(datas, text='Emissao  De:').pack(side='left')
        self.var_de = tk.StringVar(value=hoje)
        self.ent_de = ttk.Entry(datas, textvariable=self.var_de, width=12)
        self.ent_de.pack(side='left', padx=3)
        ttk.Label(datas, text='Ate:').pack(side='left')
        self.var_ate = tk.StringVar(value=hoje)
        self.ent_ate = ttk.Entry(datas, textvariable=self.var_ate, width=12)
        self.ent_ate.pack(side='left', padx=3)
        self.var_de.trace_add('write', lambda *a: self.preencher())
        self.var_ate.trace_add('write', lambda *a: self.preencher())
        ttk.Button(datas, text='Hoje', width=7,
                   command=lambda: self.periodo(0)).pack(side='left', padx=(8, 2))
        ttk.Button(datas, text='7 dias', width=7,
                   command=lambda: self.periodo(6)).pack(side='left', padx=2)
        ttk.Button(datas, text='Mes', width=7,
                   command=self.periodo_mes).pack(side='left', padx=2)
        ttk.Button(datas, text='Limpar', width=7,
                   command=self.periodo_limpar).pack(side='left', padx=2)
        self.var_contagem = tk.StringVar(value='')
        ttk.Label(datas, textvariable=self.var_contagem,
                  foreground='#333').pack(side='left', padx=12)

        meio = ttk.Frame(self, padding=(6, 0))
        meio.pack(fill='both', expand=True)
        cols = ('nota', 'serie', 'cliente', 'emissao', 'itens', 'volumes', 'peso', 'status')
        self.tree = ttk.Treeview(meio, columns=cols, show='tree headings',
                                 selectmode='browse')
        self.tree.heading('#0', text='')
        self.tree.column('#0', width=38, minwidth=38, stretch=False, anchor='center')
        self.tree.configure(style='Rom.Treeview')
        ttk.Style(self).configure('Rom.Treeview', rowheight=22)
        larguras = {'nota': 70, 'serie': 45, 'cliente': 300, 'emissao': 130,
                    'itens': 50, 'volumes': 65, 'peso': 80, 'status': 110}
        titulos = {'nota': 'Nota', 'serie': 'Serie', 'cliente': 'Cliente',
                   'emissao': 'Emissao', 'itens': 'Itens', 'volumes': 'Volumes',
                   'peso': 'Peso Bruto', 'status': 'Status'}
        for c in cols:
            self.tree.heading(c, text=titulos[c])
            self.tree.column(c, width=larguras[c],
                             anchor='center' if c != 'cliente' else 'w')
        self.tree.tag_configure('ruim', background='#ffd6d6')
        self.tree.tag_configure('alerta', background='#fff3cd')
        self.tree.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(meio, orient='vertical', command=self.tree.yview)
        sb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind('<Button-1>', self.clique)
        self.tree.bind('<Double-1>', self.duplo_clique)
        self.tree.bind('<space>', self.tecla_espaco)

        base = ttk.Frame(self, padding=6)
        base.pack(side='bottom', fill='x')
        ttk.Button(base, text='Marcar todas', command=lambda: self.marcar(True)).pack(side='left')
        ttk.Button(base, text='Desmarcar', command=lambda: self.marcar(False)).pack(side='left', padx=4)
        ttk.Button(base, text='GERAR ROMANEIO', command=self.gerar).pack(side='left', padx=16)
        ttk.Button(base, text='Enviar por e-mail', command=self.email).pack(side='left')
        ttk.Label(base, text='Copias:').pack(side='left', padx=(18, 2))
        self.var_copias = tk.IntVar(value=1)
        ttk.Spinbox(base, from_=1, to=20, width=4, textvariable=self.var_copias).pack(side='left')
        ttk.Button(base, text='Imprimir notas',
                   command=lambda: self.imprimir('notas')).pack(side='left', padx=4)
        ttk.Button(base, text='Imprimir romaneio',
                   command=lambda: self.imprimir('romaneio')).pack(side='left')
        ttk.Button(base, text='Configuracoes...',
                   command=self.abrir_config).pack(side='right')

        menu = tk.Menu(self)
        m_arq = tk.Menu(menu, tearoff=0)
        m_arq.add_command(label='Configuracoes...', command=self.abrir_config)
        m_arq.add_command(label='Reconectar ao servidor',
                          command=lambda: self.atualizar(reconectar=True))
        m_arq.add_separator()
        m_arq.add_command(label='Sair', command=self.destroy)
        menu.add_cascade(label='Arquivo', menu=m_arq)
        self.config(menu=menu)

        self.var_status = tk.StringVar(value='Pronto.')
        ttk.Label(self, textvariable=self.var_status, relief='sunken',
                  anchor='w', padding=4).pack(side='bottom', fill='x')

    # -------------------------------------------------- acoes
    def periodo(self, dias_atras):
        hoje = date.today()
        self.var_de.set((hoje - timedelta(days=dias_atras)).strftime('%d/%m/%Y'))
        self.var_ate.set(hoje.strftime('%d/%m/%Y'))

    def periodo_mes(self):
        hoje = date.today()
        self.var_de.set(hoje.replace(day=1).strftime('%d/%m/%Y'))
        self.var_ate.set(hoje.strftime('%d/%m/%Y'))

    def periodo_limpar(self):
        self.var_de.set('')
        self.var_ate.set('')

    def _intervalo(self):
        """Devolve (de, ate, erro). Campo vazio = sem limite."""
        d1 = self.var_de.get().strip()
        d2 = self.var_ate.get().strip()
        de = para_data(d1) if d1 else None
        ate = para_data(d2) if d2 else None
        erro = (d1 and not de) or (d2 and not ate)
        self.ent_de.configure(foreground='#a00000' if (d1 and not de) else 'black')
        self.ent_ate.configure(foreground='#a00000' if (d2 and not ate) else 'black')
        if de and ate and de > ate:
            return de, ate, 'invertido'
        return de, ate, ('data invalida' if erro else None)

    def salvar_config(self):
        try:
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(self.cfg, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def abrir_config(self):
        JanelaConfig(self)

    def atualizar_aviso_impressao(self):
        if impressao_nativa_disponivel():
            return
        self.var_status.set('Impressao nativa indisponivel - instale com: '
                            'pip install pywin32 pypdfium2')

    def escolher_pasta(self):
        p = filedialog.askdirectory()
        if not p:
            return
        p = normalizar_caminho(p)
        srv = self.cfg.setdefault('servidor', {})
        if srv.get('ativo') and not p.startswith('\\\\'):
            if not messagebox.askyesno(
                    'Servidor das notas',
                    f"As notas estao sendo lidas do servidor {srv.get('host')}.\n\n"
                    f'Passar a ler desta pasta do computador?\n{p}'):
                return
            srv['ativo'] = False
            self.salvar_config()
        self.var_pasta.set(p)

    def pasta_configurada(self):
        """Pasta das notas em uso: a do servidor quando ele esta ligado."""
        srv = self.cfg.get('servidor') or {}
        if srv.get('ativo'):
            return montar_caminho_rede(srv) or (self.cfg.get('pasta_notas') or '')
        return self.cfg.get('pasta_notas') or ''

    def pedir_senha_servidor(self, pasta):
        """Pede a senha do servidor uma unica vez e guarda no computador."""
        srv = self.cfg.get('servidor') or {}
        if (not srv.get('ativo') or srv.get('usar_credenciais_windows')
                or not srv.get('usuario')):
            return True
        host = normalizar_caminho(srv.get('host')).lstrip('\\')
        if not host or host_do_caminho(pasta).lower() != host.lower():
            return True
        if senha_rede(srv):
            return True
        s = simpledialog.askstring(
            'Servidor das notas',
            f"Senha de {srv.get('usuario')} em {host}:", show='*', parent=self)
        if not s:
            self.var_status.set('Senha do servidor nao informada.')
            return False
        if not salvar_senha_rede(host, srv.get('usuario'), s):
            messagebox.showwarning('Servidor das notas',
                                   'Nao consegui guardar a senha neste computador. '
                                   'Ela sera pedida de novo na proxima vez.')
        return True

    def atualizar(self, reconectar=False):
        pasta = self.var_pasta.get().strip() or self.pasta_configurada()
        if not pasta:
            messagebox.showwarning('Pasta', 'Informe a pasta das notas em '
                                            'Configuracoes > Servidor das notas.')
            return
        self.var_pasta.set(pasta)
        na_rede = pasta.startswith('\\\\')
        if na_rede and not self.pedir_senha_servidor(pasta):
            return
        self.var_status.set('Conectando ao servidor...' if na_rede else 'Lendo notas...')
        self.update_idletasks()

        def aviso(titulo, texto, status):
            self.after(0, lambda: self.var_status.set(status))
            self.after(0, lambda: messagebox.showwarning(titulo, texto))

        def tarefa():
            try:
                destino, erro = preparar_pasta_notas(
                    self.cfg, pasta, log=lambda m: self.var_status.set(m),
                    reconectar=reconectar)
            except Exception as ex:
                destino, erro = pasta, str(ex)
            if erro == SENHA_AUSENTE:
                erro = ('Falta a senha do servidor.\n\n'
                        'Preencha em Configuracoes > Servidor das notas.')
            if erro:
                aviso('Servidor das notas', erro, 'Sem acesso a pasta das notas.')
                return
            if not os.path.isdir(destino):
                aviso('Pasta', f'Pasta invalida ou fora do ar:\n{destino}',
                      'Pasta das notas indisponivel.')
                return
            self.cfg['pasta_notas'] = destino
            self.after(0, lambda: self.var_pasta.set(destino))
            self.after(0, self.salvar_config)
            try:
                dados = varrer_pasta(destino, log=lambda m: self.var_status.set(m))
            except Exception as ex:
                aviso('Erro', str(ex), 'Falha ao ler as notas.')
                return
            self.notas = dados
            self.marcadas.clear()
            self.after(0, self.preencher)
            self.after(0, lambda: self.var_status.set(f'{len(dados)} nota(s) na pasta.'))

        threading.Thread(target=tarefa, daemon=True).start()

    def preencher(self):
        filtro = self.var_filtro.get().lower().strip()
        de, ate, erro = self._intervalo()
        self.tree.delete(*self.tree.get_children())
        visiveis = 0
        for i, d in enumerate(self.notas):
            alvo = f"{d.get('nota') or ''} {d.get('cliente') or ''} {d.get('arquivo')}".lower()
            if filtro and filtro not in alvo:
                continue
            dt = para_data(d.get('emissao'))
            if dt is not None and not erro:
                if de and dt < de:
                    continue
                if ate and dt > ate:
                    continue
            visiveis += 1
            tag = ''
            if d.get('status') in ('DIVERGENCIA', 'ILEGIVEL', 'ERRO'):
                tag = 'ruim'
            elif d.get('status') == 'SEM CHECKSUM' or d.get('volumes') is None:
                tag = 'alerta'
            self.tree.insert('', 'end', iid=str(i), tags=(tag,),
                             image=(self.img_on if i in self.marcadas else self.img_off),
                             values=(
                d.get('nota') or '?', d.get('serie') or '-',
                (d.get('cliente') or '-')[:60], d.get('emissao') or '-',
                len(d.get('itens') or []),
                fmt(d.get('volumes'), 0), fmt(d.get('peso_bruto'), 3),
                d.get('status', '?')))
        aviso = ''
        if erro == 'invertido':
            aviso = '  (data inicial maior que a final)'
        elif erro:
            aviso = '  (data invalida - use dd/mm/aaaa)'
        self.var_contagem.set(f'{visiveis} de {len(self.notas)} nota(s) no periodo{aviso}')

    def clique(self, ev):
        if self.tree.identify_region(ev.x, ev.y) not in ('tree', 'image'):
            return
        iid = self.tree.identify_row(ev.y)
        if iid:
            self.alternar(iid)

    def tecla_espaco(self, _ev=None):
        for iid in self.tree.selection():
            self.alternar(iid)
        return 'break'

    def alternar(self, iid):
        i = int(iid)
        if i in self.marcadas:
            self.marcadas.discard(i)
        else:
            self.marcadas.add(i)
        self.tree.item(iid, image=(self.img_on if i in self.marcadas else self.img_off))
        self.var_status.set(f'{len(self.marcadas)} nota(s) marcada(s).')

    def duplo_clique(self, ev):
        iid = self.tree.identify_row(ev.y)
        col = self.tree.identify_column(ev.x)
        if not iid:
            return
        i = int(iid)
        if col == '#6':
            v = simpledialog.askinteger('Volumes',
                                        f"Volumes da nota {self.notas[i].get('nota')}:",
                                        initialvalue=int(self.notas[i].get('volumes') or 0),
                                        minvalue=0, parent=self)
            if v is not None:
                self.notas[i]['volumes'] = float(v)
                self.tree.set(iid, 'volumes', fmt(v, 0))
        else:
            try:
                os.startfile(self.notas[i]['caminho'])
            except Exception:
                pass

    def marcar(self, ligar):
        vis = [int(x) for x in self.tree.get_children()]
        if ligar:
            self.marcadas.update(vis)
        else:
            self.marcadas.difference_update(vis)
        self.preencher()

    def selecionadas(self):
        return [self.notas[i] for i in sorted(self.marcadas)]

    def gerar(self):
        sel = self.selecionadas()
        if not sel:
            messagebox.showinfo('Romaneio', 'Nenhuma nota marcada.')
            return
        ruins = [d for d in sel if d.get('status') in ('DIVERGENCIA', 'ILEGIVEL', 'ERRO')]
        if ruins:
            messagebox.showerror(
                'Conferencia bloqueada',
                'Estas notas nao passaram na conferencia e nao podem entrar no romaneio:\n\n'
                + '\n'.join(f"{d.get('nota') or d['arquivo']} - {d['status']}" for d in ruins))
            return
        visiveis = {int(x) for x in self.tree.get_children()}
        escondidas = [self.notas[i] for i in sorted(self.marcadas) if i not in visiveis]
        if escondidas:
            if not messagebox.askyesno(
                    'Notas fora do filtro',
                    f'{len(escondidas)} nota(s) marcada(s) nao aparecem no filtro atual:\n'
                    + ', '.join(str(d.get('nota')) for d in escondidas[:12])
                    + '\n\nIncluir no romaneio mesmo assim?'):
                return
        sem_vol = [d for d in sel if not d.get('volumes')]
        if sem_vol:
            if not messagebox.askyesno(
                    'Volumes em branco',
                    'Nota(s) sem volumes: '
                    + ', '.join(str(d.get('nota')) for d in sem_vol)
                    + '\n\nDe duplo clique na coluna Volumes para corrigir.\nGerar assim mesmo?'):
                return
        saida = self.cfg.get('pasta_saida') or self.var_pasta.get()
        os.makedirs(saida, exist_ok=True)
        numero = proximo_romaneio(self.cfg)
        caminho = os.path.join(saida, f"ROMANEIO_{numero:06d}.pdf")
        try:
            gerar_romaneio(sel, caminho, numero, self.cfg)
        except Exception as e:
            messagebox.showerror('Erro ao gerar', str(e))
            return
        self.ultimo_romaneio = caminho
        self.ultimo_numero = numero
        self.ultimas_notas = list(sel)
        self.var_status.set(f'Romaneio {numero:06d} gerado: {caminho}')
        JanelaRevisao(self, caminho, sel, numero)

    def email(self):
        if not self.ultimo_romaneio or not os.path.exists(self.ultimo_romaneio):
            messagebox.showinfo('E-mail', 'Gere o romaneio primeiro.')
            return
        JanelaRevisao(self, self.ultimo_romaneio,
                      getattr(self, 'ultimas_notas', self.selecionadas()),
                      self.ultimo_numero)

    def imprimir(self, alvo):
        if alvo == 'romaneio':
            if not self.ultimo_romaneio:
                messagebox.showinfo('Imprimir', 'Gere o romaneio primeiro.')
                return
            arquivos = [self.ultimo_romaneio]
        else:
            sel = self.selecionadas()
            if not sel:
                messagebox.showinfo('Imprimir', 'Nenhuma nota marcada.')
                return
            arquivos = [d['caminho'] for d in sel]
        JanelaImpressao(self, arquivos,
                        'Imprimir romaneio' if alvo == 'romaneio' else 'Imprimir notas')


def registrar_erro(tipo, valor, tb):
    """Sem console (pythonw), erro sem log e erro invisivel."""
    import traceback
    destinos = [os.path.join(APP_DIR, 'erro_romaneio.log'),
                os.path.join(os.path.expanduser('~'), 'erro_romaneio.log')]
    gravado = None
    for caminho in destinos:
        try:
            with open(caminho, 'a', encoding='utf-8') as f:
                f.write(f"\n===== {datetime.now():%d/%m/%Y %H:%M:%S} =====\n")
                traceback.print_exception(tipo, valor, tb, file=f)
            gravado = caminho
            break
        except Exception:
            continue
    try:
        messagebox.showerror('Erro no programa',
                             f'{tipo.__name__}: {valor}\n\n'
                             + (f'Detalhes gravados em:\n{gravado}' if gravado else ''))
    except Exception:
        pass


if __name__ == '__main__':
    sys.excepthook = registrar_erro
    tk.Tk.report_callback_exception = staticmethod(registrar_erro)
    try:
        App().mainloop()
    except Exception:
        registrar_erro(*sys.exc_info())