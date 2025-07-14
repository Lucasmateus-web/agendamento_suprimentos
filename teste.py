import os, re, pandas as pd, smtplib, math
from fpdf import FPDF
from datetime import datetime
from email.message import EmailMessage
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputFile
from telegram.ext import ApplicationBuilder, CommandHandler, CallbackQueryHandler, ContextTypes
from telegram.ext import MessageHandler, filters
from calendar import month_name
from openai import AsyncOpenAI
import nest_asyncio
import random
from datetime import timedelta
import matplotlib.pyplot as plt
from PIL import Image
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from io import BytesIO
from hashlib import sha1
from textwrap import wrap


# ─── CONFIGURAÇÕES ─────────────────────────────────────────────────────────
TOKEN_TELEGRAM = "7531357054:AAFV6q9OnddvDdJrtkIvNyM_4IJ93fecjVE"
OPENAI_API_KEY = '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'
SMTP_SERVER = 'smtp.office365.com'
SMTP_PORT = 587
SMTP_USER = 'lucas.mateus@engeman.net'
SMTP_PASSWORD = 'engeman2025@'

openai_client = AsyncOpenAI(api_key=OPENAI_API_KEY)

usuarios_iniciados = set()


# ─── FUNÇÕES DE CARGA DE DADOS ─────────────────────────────────────────────
fornecedor_id_map = {}
fornecedor_trend_map = {}


def carregar_dados_qualidade():
    df = pd.read_excel('atendimento controle_qualidade.xlsx')
    df['mes'] = pd.to_datetime(df['data']).dt.strftime('%m/%Y')
    return df

def carregar_dados_emails():
    return pd.read_excel('emails.xlsx')

def carregar_dados_homologados():
    return pd.read_excel('fornecedores_homologados.xlsx')

def get_meses_e_fornecedores():
    df = carregar_dados_qualidade()
    meses = sorted(df['mes'].unique())
    fornecedores = {f"f{i}": nome for i, nome in enumerate(df['nome_agente'].unique())}
    return meses, fornecedores

 # ─── FUNÇÕES DE CARGA DE DADOS ─────────────────────────────────────────────
def carregar_dados_qualidade():
    df = pd.read_excel('atendimento controle_qualidade.xlsx')
    df['mes'] = pd.to_datetime(df['data']).dt.strftime('%m/%Y')
    return df

def carregar_dados_emails():
    return pd.read_excel('emails.xlsx')

def carregar_dados_homologados():
    return pd.read_excel('fornecedores_homologados.xlsx')

def get_meses_e_fornecedores():
    df = carregar_dados_qualidade()
    meses = sorted(df['mes'].unique())
    fornecedores = {f"f{i}": nome for i, nome in enumerate(df['nome_agente'].unique())}
    return meses, fornecedores

# ─── FUNÇÃO PARA OBTER MESES DISPONÍVEIS ───────────────────────────────────
def obter_meses_disponiveis():
    df = carregar_dados_qualidade()  # Carrega os dados de qualidade
    meses = sorted(df['mes'].unique())  # Pega os meses únicos e os ordena
    print(f"Meses Disponíveis: {meses}")  # Verifique os meses aqui
    return meses

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    d = update.callback_query.data
    await update.callback_query.answer()

    if d == "menu_mensal":
        meses_disponiveis = obter_meses_disponiveis()  # Chama a função
        kb = [[InlineKeyboardButton(m, callback_data=f"mes_{m.replace('/', '-')}")] for m in meses_disponiveis]
        return await update.callback_query.message.edit_text("📅 Escolha o mês:", reply_markup=InlineKeyboardMarkup(kb))

    # Outros botões aqui...


def carregar_df_vencimentos():
    df = pd.read_excel("fornecedores_homologados.xlsx")
    df.columns = df.columns.str.strip().str.lower()
    
    if 'data vencimento' not in df.columns:
        raise ValueError("Coluna 'data vencimento' não encontrada.")

    df['data vencimento'] = pd.to_datetime(df['data vencimento'], errors='coerce')
    df = df[df['data vencimento'].notna()]
    df['mes_ano'] = df['data vencimento'].dt.strftime("%m/%Y")
    return df


# ─── FUNÇÕES AUXILIARES ────────────────────────────────────────────────────
def limpar_texto_pdf(t):
    if not isinstance(t, str): t = str(t)
    return re.sub(r'[^\x00-\x7F]+', '', t).replace('–','-').replace('—','-')


# MONTA CORPO DE E-MAIL ────────────────────────────────────────────────────

def montar_corpo_email(texto_feedback, iqf_formatado="0.00"):
    # Converte markdown leve e quebras de linha para HTML
    texto = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', texto_feedback.strip())  # Converte negrito
    texto = texto.replace('\n', '<br>')  # Substitui quebras de linha por <br> em HTML

    return f"""\
<html>
  <body style="font-family: Arial, sans-serif; font-size: 13pt; color: #000;">
    <p>
      {texto}
    </p>
    <br><br>
    <p style="font-size: 11pt; color: #555;">
      Este e-mail foi gerado automaticamente com base na avaliação periódica do Índice de Qualidade do Fornecedor (IQF), conforme os critérios definidos nos Procedimentos PG.SM.01, PG.SM.02 e PG.SM.03 da ENGEMAN.
    </p>
    <br>
    <p style="font-size: 12pt;">
      Atenciosamente,<br>
      <b>Equipe de Suprimentos</b><br>
      <span style="font-size: 11pt; color: #777;">Avaliação de Desempenho.</span>
  </body>
</html>
"""



# GERA ANÁLISE COM O GPT PARA OS FORNECEDORES ────────────────────────────────────────────────────

async def gerar_analise_gpt(prompt):
    resp = await openai_client.chat.completions.create(
        model='gpt-4o',
        messages=[{"role":"user","content":prompt}],
        temperature=0.7
    )
    return resp.choices[0].message.content


# FUNÇÃO PARA ENVIAR O E-MAIL AO DESTINATÁRIO ────────────────────────────────────────────────────

def enviar_email(dest, subj, body, apath):
    msg = EmailMessage()
    msg['Subject'], msg['From'], msg['To'] = subj, SMTP_USER, dest
    msg.set_content(body)
    with open(apath, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(apath))
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
        s.starttls()
        s.login(SMTP_USER, SMTP_PASSWORD)
        s.send_message(msg)


# FUNÇÃO QUE CRIA E FINALIZA O PDF ────────────────────────────────────────────────────

async def finalizar_envio_pdf(update, context, path, tipo, iden, corpo_email):
    await update.callback_query.message.reply_document(InputFile(path, os.path.basename(path)))
    context.user_data['arquivo_pdf'] = path
    context.user_data['ultimo_texto'] = corpo_email
    context.user_data['ultimo_tipo'] = tipo
    context.user_data['ultimo_iden'] = iden
    context.user_data['aguardando_email'] = True
    await update.callback_query.message.reply_text("📧 Informe o e-mail para o qual deseja enviar esta análise:")

class PDF(FPDF):
    def __init__(self, tipo='analise', fornecedor=None, mes=None):
        super().__init__()
        self.tipo = tipo
        self.fornecedor = fornecedor
        self.mes = mes

    def header(self):
    # Logo
        if os.path.exists('engeman_logo.png'):
            self.image('engeman_logo.png', 10, 8, 33)

            self.set_font('Helvetica', 'B', 14)

    # Ajuste da posição abaixo da logo
            self.set_y(20)

    # Definição do título com base no tipo
        if getattr(self, 'tipo', None) == 'analise' and self.mes:
            titulo = f"FORNECEDORES REPROVADOS – {self.mes}"
        elif getattr(self, 'tipo', None) == 'feedback' and self.fornecedor:
            partes = self.fornecedor.upper().split()
            nome_curto = ' '.join(partes[:2]) if len(partes) >= 2 else self.fornecedor.upper()
            titulo = f"FEEDBACK – {nome_curto}"
        else:
            titulo = "RELATÓRIO DE FORNECEDOR"

    # Impressão do título centralizado
        self.cell(0, 10, limpar_texto_pdf(titulo), align="C")
        self.ln(10)


    def footer(self):
        self.set_y(-15)
        self.set_draw_color(200, 200, 200)
        self.line(10, self.get_y(), 200, self.get_y())
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, limpar_texto_pdf(f'Emitido em: {datetime.now():%d/%m/%Y %H:%M}'), 0, 0, 'C')

    def tabela_reprovados(self, dados):
        self.set_text_color(0, 0, 0)
        self.set_fill_color(230, 230, 230)
        self.set_font('Helvetica', 'B', 11)
        self.cell(90, 8, 'FORNECEDOR', border=1, align='C', fill=True)
        self.cell(25, 8, 'NOTA IQF', border=1, align='C', fill=True)
        self.cell(50, 8, 'DOCUMENTO', border=1, ln=1, align='C', fill=True)

        self.set_font('Helvetica', '', 10)
        fill = False
        row_height = 6

        for nome, nota, documento in dados:
        # ❌ Removido o filtro que estava pulando reprovados
        # Se já filtrou antes, não precisa checar nada aqui

            if self.get_y() + row_height > self.h - 20:
                self.add_page()
                self.set_text_color(0, 0, 0)
                self.set_fill_color(230, 230, 230)
                self.cell(90, 8, 'FORNECEDOR', border=1, align='C', fill=True)
                self.cell(25, 8, 'NOTA IQF', border=1, align='C', fill=True)
                self.cell(50, 8, 'DOCUMENTO', border=1, ln=1, align='C', fill=True)
                self.set_font('Helvetica', '', 10)

            partes = nome.upper().split()
            nome_resumido = ' '.join(partes[:2]) if len(partes) >= 2 else nome.upper()

            self.set_fill_color(245, 245, 245) if fill else self.set_fill_color(255, 255, 255)
            self.set_text_color(0, 0, 0)
            self.cell(90, row_height, nome_resumido[:40], border=1, align='C', fill=fill)

            self.set_text_color(255, 0, 0)
            self.cell(25, row_height, f"{nota:.2f}", border=1, align='C', fill=fill)

            self.set_text_color(0, 0, 0)
            self.cell(50, row_height, str(documento), border=1, ln=1, align='C', fill=fill)

            fill = not fill

# ─── FUNÇÃO AUXILIAR PARA EXTRAÇÃO DE CRITÉRIOS CRÍTICOS ──────────────────────

def gerar_criterios_criticos(df_fornecedor):
    criterios_mapeados = {
        11: "Cumprimento de prazos",
        12: "Conformidade com os itens do pedido",
        13: "Comunicação, garantia e suporte pós-venda",
        14: "Embalagem e identificação do material",
        24: "Envio de documentos obrigatórios",
        25: "Qualidade do material e/ou serviço entregue",
        26: "Cumprimento às normas de segurança",
    }

    criterios_reprovados = set()
    for _, row in df_fornecedor.iterrows():
        nota = row['nota']
        cod = row.get('qualificacao') or row.get('criterio')  # use o nome real da coluna
        if nota in [0, 50] and cod in criterios_mapeados:
            criterios_reprovados.add(criterios_mapeados[cod])

    return sorted(criterios_reprovados)


async def handle_feedback_individual(update, context, fornecedor):
    # Inicia o processo de análise
    await update.callback_query.message.edit_text(f"⏳ Gerando feedback individual: {fornecedor}…")
    
    df_qualidade = carregar_dados_qualidade()
    df = df_qualidade[df_qualidade['nome_agente'] == fornecedor]

    if df.empty:
        texto = "Não há dados suficientes para gerar feedback."
        return await update.callback_query.message.reply_text(texto)

    iqf = df['nota'].mean()
    if math.isnan(iqf) or iqf <= 0:
        texto = "Não há dados suficientes para gerar feedback."
        return await update.callback_query.message.reply_text(texto)

    # Obter o resumo técnico das ocorrências
    ocorrencias = get_resumo_ocorrencias_geral(fornecedor)
    
    # Texto fixo inicial para todos os fornecedores
    texto = f"Olá, prezado!\n\n"

    # Se houver ocorrências, adiciona ao texto; se não, exibe mensagem de "nenhuma ocorrência"
 


    # Determinar a classificação com base no IQF
    if iqf >= 75:
        texto += (
            f"Agradecemos pela parceria e informamos que sua empresa obteve uma excelente performance na nossa avaliação mais recente.\n\n"
            f"📊 Nota IQF: {iqf:.2f}\n🏆 Classificação: Aprovado – Desempenho de excelência.\n\n"
            f"Esse resultado demonstra alto nível de comprometimento com qualidade, prazos e conformidade. "
            f"Seguimos confiantes na continuidade desta parceria sólida e eficiente.\n\n"
            f"Qualquer dúvida estamos à disposição.\n\n"
        )

        if ocorrencias:
            texto += f"🔴 <b>Ocorrências registradas no seu atendimento:</b>\n{ocorrencias}\n\n"

    elif 70 <= iqf < 75:
        criterios_falhos = random.sample([ 
            "Cumprimento de prazos conforme o pedido de compra/contrato.",
            "Comunicação, garantia e suporte pós-venda do fornecedor.",
            "Qualidade do material e/ou serviço entregue.",
            "Conformidade com os itens descrito no pedido/contrato."
        ], k=2)

        texto += (
            f"Compartilhamos abaixo o resultado da nossa avaliação periódica de desempenho:\n\n"
            f"📊 Nota IQF: {iqf:.2f}\n⚠️ Nota mínima exigida: 70,00\n\n"
            f"Embora sua avaliação esteja tecnicamente aprovada, o resultado indica que sua performance está no limite mínimo aceitável conforme os critérios estabelecidos no Procedimento PG.SM.02.\n\n"
            f"Recomendamos atenção especial aos seguintes aspectos:\n"
            + "\n".join(f"- {c}" for c in criterios_falhos) +
            "\n\nA manutenção de bons indicadores é essencial para seguirmos com uma parceria de confiança e excelência.\n\n"
        )

        if ocorrencias:
            texto += f"🔴 <b>Ocorrências registradas no seu atendimento:</b>\n{ocorrencias}\n\n"

    elif iqf < 70:
        criterios_falhos = random.sample([ 
            "Cumprimento de prazos conforme o pedido de compra/contrato.",
            "Conformidade com os itens descritos no pedido/contrato.",
            "Qualidade do material e/ou serviço entregue.",
            "Comunicação, garantia e suporte pós-venda do fornecedor.",
            "Embalagem e identificação do material.",
            "Cumprimento às normas de segurança.",
            "Envio de documentos obrigatórios tais como (Boleto, Notas Fiscais & Certificados necessários)."
        ], k=3)

        texto += (
            f"Informamos que, conforme nossa avaliação periódica de desempenho, sua empresa foi reprovada no Índice de Qualidade do Fornecedor (IQF):\n\n"
            f"📊 Nota IQF: {iqf:.2f}\n❌ Classificação: Reprovado – Abaixo do padrão mínimo (70,00)\n\n"
            f"A reprovação ocorreu devido a falhas identificadas nos seguintes critérios:\n"
            + "\n".join(f"- {c}" for c in criterios_falhos) +
            "\n\nSolicitamos análise interna das não conformidades e implementação de medidas corretivas. "
            f"A reincidência poderá impactar futuros fornecimentos.\n\n"
        )

        if ocorrencias:
            texto += f"🔴 <b>Ocorrências registradas no seu atendimento:</b>\n{ocorrencias}\n\n"

    # Legendas de notas
    texto += (
        f"<b>Legendas de Notas:</b>\n\n"
        f"- <b>0 à 69: Reprovado</b> – Significa que o fornecedor não atingiu os critérios mínimos de qualidade e conformidade exigidos para a aprovação.\n\n"
        f"- <b>A partir de 70: Aprovado</b> – Significa que o fornecedor atendeu adequadamente aos critérios estabelecidos, com um desempenho satisfatório.\n\n"
        f"Em caso de apontamentos negativos, pedimos a análise e correção. A reincidência de problemas pode suspendê-lo como fornecedor da Engeman.\n\n"
        f"Seguimos confiantes na continuidade desta parceria sólida e eficiente.\n\n"
    )

    # Finaliza o feedback com o texto gerado
    await update.callback_query.message.reply_text(texto, parse_mode="HTML")

    # Salvar o feedback como PDF
    os.makedirs('pdfs', exist_ok=True)
    path = f"pdfs/Feedback_Individual_{fornecedor}.pdf"
    pdf = PDF(tipo='feedback', fornecedor=fornecedor)
    pdf.add_page()
    pdf.set_font('Helvetica', 'B', 14)
    pdf.cell(0, 10, limpar_texto_pdf(f"Feedback do Fornecedor - {fornecedor}"), ln=1, align='C')
    pdf.ln(5)
    pdf.set_font('Helvetica', '', 12)
    pdf.multi_cell(0, 8, limpar_texto_pdf(texto))
    pdf.output(path)

    await finalizar_envio_pdf(update, context, path, "ind", fornecedor, texto)


# Função para obter as ocorrências de um fornecedor específico
def get_ocorrencias_fornecedor(fornecedor):
    df_ocorrencias = pd.read_excel("Ocorrencias.xlsx")
    ocorrencias = df_ocorrencias[df_ocorrencias['FORNECEDOR'] == fornecedor]['OCORRÊNCIAS'].tolist()
    return ocorrencias

def get_resumo_ocorrencias_geral(fornecedor):
    # Carrega os dados de ocorrências
    df_ocorrencias = pd.read_excel("Ocorrencias.xlsx")

    # Filtra as ocorrências específicas do fornecedor
    ocorrencias_fornecedor = df_ocorrencias[df_ocorrencias['FORNECEDOR'] == fornecedor]

    # Remover entradas vazias ou NaN da coluna 'OCORRÊNCIAS'
    ocorrencias_fornecedor = ocorrencias_fornecedor.dropna(subset=['OCORRÊNCIAS'])


    # Gerar resumo técnico das ocorrências, organizadas em tópicos
    resumo_ocorrencias = []

    # Usar um conjunto para armazenar ocorrências únicas (removendo duplicatas)
    ocorrencias_unicas = set()

    for _, row in ocorrencias_fornecedor.iterrows():
        descricao = row['OCORRÊNCIAS']
        
        # Limpeza de quebras de linha e caracteres estranhos
        descricao = descricao.replace("\n", " ").replace("\r", " ").strip()  # Remove quebras de linha e espaços extras
        
        # Mantém os espaços entre as palavras, mas remove caracteres extra
        descricao = ' '.join(descricao.split())  # Remove espaços extras entre palavras

        # Se a descrição não for vazia ou 'nan', adicionamos ao conjunto de ocorrências únicas
        if descricao and descricao.lower() != "nan":
            ocorrencias_unicas.add(descricao)
    
    # Limitar o número de ocorrências para exibir, por exemplo, 5
    max_ocorrencias = 5
    ocorrencias_fornecedor_resumidas = list(ocorrencias_unicas)[:max_ocorrencias]

    for descricao in ocorrencias_fornecedor_resumidas:
        # Adiciona a ocorrência no formato de tópico (em negrito no Telegram)
        resumo_ocorrencias.append(f"• {descricao}")
    
    # Se houver mais ocorrências do que o limite, adicionamos um aviso
    if len(ocorrencias_unicas) > max_ocorrencias:
        resumo_ocorrencias.append(f"\nE mais {len(ocorrencias_unicas) - max_ocorrencias} ocorrências não exibidas.")
    
    # Resumo técnico com as principais ocorrências
    resumo_tecnico = "\n".join(resumo_ocorrencias)
    return resumo_tecnico
    
async def handle_analise_mensal(update, context, mes):
    await update.callback_query.message.edit_text(f"⏳ Gerando análise mensal: {mes}…")
    df_qualidade = carregar_dados_qualidade() 
    dfm = df_qualidade[df_qualidade['mes'] == mes]
    if dfm.empty:
        return await update.callback_query.message.reply_text(f"⚠️ Nenhum dado para {mes}")

    avg = round(dfm['nota'].mean(), 2)
    reprov = dfm[dfm['nota'] < 70]['nome_agente'].unique().tolist()

    prompt = (
        f"Relatório técnico para o mês {mes} com IQF médio de {avg}. "
        f"Reprovados: {', '.join(reprov) or 'Nenhum'}. "
        "Inclua seções: Visão Geral, Pontos de Atenção, Reprovados, Conclusão. "
        "Cada seção com ao menos uma frase ou 'Nenhum registro'. Máximo 10 linhas."
        "E em Ações Tomadas deixe o texto fixo: Envio de notificação via e-mail aos fornecedores reprovados no IQF mensal"
        "e em caso de reincidência haverá abertura de RAC para análise, tratativas e possível suspensão do fornecedor."
    )
    
    a = await gerar_analise_gpt(prompt)

    def extrair_secao(texto, secao):
        m = re.search(rf"{secao}:(.*?)(?=\n[A-ZÀ-Ú][^\n]*:|\Z)", texto, flags=re.S | re.I)
        return m.group(1).strip() if m and m.group(1).strip() else "Nenhum registro."
    


    sec_v = extrair_secao(a, "Visão Geral")
    sec_pa = extrair_secao(a, "Pontos de Atenção")
    sec_rp = extrair_secao(a, "Reprovados")
    sec_conc = extrair_secao(a, "Conclusão")


    sec_acoes = (
        "Envio de notificação via e-mail aos fornecedores reprovados no IQF mensal "
        "e em caso de reincidência haverá abertura de RAC para análise, tratativas e possível suspensão do fornecedor."
    )

    # Monta o texto do chat
    chat_txt = (
        f"📅 *Análise Mensal – {mes}*\n\n"
        f"🔍 *Visão Geral:*\n{sec_v}\n\n"
        f"⚠️ *Pontos de Atenção:*\n{sec_pa}\n\n"
        f"❌ *Reprovados:*\n{sec_rp}\n\n"
        f"✅ *Conclusão:*\n{sec_conc}"
    )


    context.user_data["sec_v"] = sec_v
    context.user_data["sec_pa"] = sec_pa
    context.user_data["sec_rp"] = sec_rp
    context.user_data["sec_acoes"] = sec_acoes  # ← esta linha garante que o e-mail terá o texto fixo
    context.user_data["sec_conc"] = sec_conc
    context.user_data["mes"] = mes


    # Geração de PDF
    os.makedirs('pdfs', exist_ok=True)
    path = f"pdfs/Analise_Mensal_{mes.replace('/', '-')}.pdf"
    pdf = PDF(tipo='analise', mes=mes)
    pdf.add_page()

    # Filtrar reprovados com nota < 70 e nota válida
    reprovados_df = dfm[(dfm['nota'] < 70) & (dfm['nota'].notnull())]

    # Criar lista com nome, nota e documento
    dados = [
        (
            row['nome_agente'],
            round(row['nota'], 2),
            str(row['documento']) if 'documento' in row and pd.notna(row['documento']) else ''
        )
        for _, row in reprovados_df.iterrows()
    ]
    dados = sorted(dados, key=lambda x: x[1])  # ordenado do pior para o menos ruim

    pdf.tabela_reprovados(dados)
    pdf.output(path)
    # Enviar para o Telegram
    await update.callback_query.message.reply_text(chat_txt, parse_mode="Markdown")
    await finalizar_envio_pdf(update, context, path, "men", mes.replace('/', '-'), a) 
    context.user_data["tipo_envio"] = "analise_mensal"
    context.user_data["corpo_analise"] = chat_txt
    context.user_data["caminho_pdf"] = path



# ─── RANKING MENSAL ─────────────────────────────────────────────────────────
async def handle_ranking(update, context, mes):
    await update.callback_query.message.edit_text(f"⏳ Gerando ranking mensal: {mes}…")
    df_qualidade = carregar_dados_qualidade() 
    dfm = df_qualidade[df_qualidade['mes']==mes]
    if dfm.empty:
        return await update.callback_query.message.reply_text(f"⚠️ Nenhum dado para {mes}")

    medias = dfm.groupby('nome_agente')['nota'].mean()
    top3 = medias.nlargest(3)
    bot3 = medias.nsmallest(3)

    path = f"pdfs/Ranking_Mensal_{mes.replace('/','-')}.pdf"
    os.makedirs('pdfs', exist_ok=True)
    pdf = PDF(); pdf.add_page()
    pdf.set_font('Helvetica','B',14)
    pdf.cell(0, 10, limpar_texto_pdf(f"Ranking Mensal - {mes}"), ln=1, align='C')
    pdf.ln(5)
    pdf.set_font('Helvetica','B',12)
    pdf.cell(0, 10, "Top 3 Melhores", ln=1)
    pdf.set_font('Helvetica','',12)
    for n,v in top3.items():
        pdf.cell(0, 8, limpar_texto_pdf(f"{n}: {round(v,2)}"), ln=1)
    pdf.ln(5)
    pdf.set_font('Helvetica','B',12)
    pdf.cell(0, 10, "Top 3 Piores", ln=1)
    pdf.set_font('Helvetica','',12)
    for n,v in bot3.items():
       pdf.cell(0, 8, limpar_texto_pdf(f"{n}: {round(v,2)}"), ln=1)
    pdf.output(path)

    chat_txt = (
        f"📊 *Ranking Mensal – {mes}*\n\n"
        "*Top 3 Melhores:*\n" +
        "\n".join(f"• {n}: {round(v,2)}" for n,v in top3.items()) +
        "\n\n*Top 3 Piores:*\n" +
        "\n".join(f"• {n}: {round(v,2)}" for n,v in bot3.items())
    )

    await update.callback_query.message.reply_text(chat_txt, parse_mode="Markdown")
    await finalizar_envio_pdf(update, context, path, "ran", mes.replace('/','-'), chat_txt)

# ─── BOTÃO MODO MENSAL: EXIBIR MESES DISPONÍVEIS ─────────────
async def mostrar_meses_disponiveis(update: Update, context: ContextTypes.DEFAULT_TYPE, categoria: str):
    df = pd.read_excel("atendimento controle_qualidade.xlsx")
    df['data'] = pd.to_datetime(df['data'], errors='coerce')
    df['mes'] = df['data'].dt.strftime('%m/%Y')  # formato: 02/2025
    df['nota'] = pd.to_numeric(df['nota'], errors='coerce')
    df = df.dropna(subset=['nota'])

    # Agrupar por fornecedor e mês
    medias = df.groupby(['nome_agente', 'mes'])['nota'].mean().reset_index()

    # Classificação
    if categoria == "mensal_aprovado":
        medias = medias[medias['nota'] > 75]
    elif categoria == "mensal_atencao":
        medias = medias[(medias['nota'] >= 70) & (medias['nota'] <= 75)]
    elif categoria == "mensal_reprovado":
        medias = medias[medias['nota'] < 70]

    meses = sorted(medias['mes'].unique(), key=lambda x: datetime.strptime(x, '%m/%Y'), reverse=True)
    botoes = [[InlineKeyboardButton(m, callback_data=f"{categoria}_mes_{m}")] for m in meses]

    await update.callback_query.message.edit_text(
        text="📆 *Selecione o mês desejado:*",
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode='Markdown'
    )


# ─── PROCEDIMENTO ENGEMAN ───────────────────────────────────────────────────
async def handle_procedimento(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = (
        "🧠 *Procedimento Engeman para Avaliação de Fornecedores*\n\n"
        "1️⃣ *PG.SM.01 - Aquisição*: Define o fluxo de compra, cotação, análise e aprovação dos materiais e serviços.\n"
        "2️⃣ *PG.SM.02 - Avaliação de Forcedores*: Critérios de desempenho (IQF), RACs e homologações.\n"
        "3️⃣ *PG.SM.03 - Almoxarifado*: Inspeções, controle e tratamento de não conformidades.\n\n"
        "🔎 *Observações Importantes:*\n"
        "- As análises são feitas mensalmente com base nas ocorrências.\n"
        "- A nota IQF varia de 0 a 100.\n"
        "- Fornecedores com IQF abaixo de 70 são *REPROVADOS*.\n"
        "- Objetivo: garantir o padrão Engeman e um relacionamento com parceiros confiáveis.\n\n"
        "📌 *Dúvidas?* Consulte o Setor de Suprimentos."
    )
    await update.callback_query.message.reply_text(texto, parse_mode="Markdown")

# ─── MENUS E BOTÕES ──────────────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message:
        chat = update.message
    else:
        chat = update.callback_query.message

    kb = [
        [InlineKeyboardButton("📍 Desempenho de Fornecedores", callback_data="menu_desempenho")],
        [InlineKeyboardButton("📊 Indicadores Mensais", callback_data="menu_indicadores")],
        [InlineKeyboardButton("🗂️ Documentações Cadastrais", callback_data="menu_documentos")],
        [InlineKeyboardButton("☎️ Suporte e Contato", callback_data="menu_suporte")],
        [InlineKeyboardButton("🧠 Procedimento Engeman", callback_data="procedimento")]
    ]

    await chat.reply_text(
        "👋 Bem-vindo à Central de Avaliação de Fornecedores da ENGEMAN!\n\n"
        "Selecione uma das opções abaixo para visualizar os indicadores, documentos e desempenho dos fornecedores conforme os critérios do nosso processo de homologação técnica.\n\n",
        reply_markup=InlineKeyboardMarkup(kb)
    )


# SUBMENUS CONFIGURADOS E DETALHADOS ──────────────────────────────────────────────────────────
    
async def menu_desempenho(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = [
        [InlineKeyboardButton("✅ Aprovados", callback_data="menu_feedback")],
        [InlineKeyboardButton("⚠️ Em Atenção", callback_data="menu_atencao")],
        [InlineKeyboardButton("❌ Reprovados", callback_data="menu_reprovados")],
        [InlineKeyboardButton("🔙 Voltar ao Início", callback_data="voltar_inicio")]
    ]
    await update.callback_query.message.edit_text(
        "📌 *Desempenho dos Fornecedores*\n\n"
        "Acompanhe o desempenho técnico dos fornecedores homologados, conforme critérios do Procedimento PG.SM.02. Selecione o grupo desejado para visualizar os resultados individuais.",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
async def menu_indicadores(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = [
        [InlineKeyboardButton("📈 Análise Mensal", callback_data="menu_mensal")],
        [InlineKeyboardButton("📊 Ranking Mensal", callback_data="menu_ranking")],
        [InlineKeyboardButton("📉 Tendência de Desempenho", callback_data="tendencia_0")],
        [InlineKeyboardButton("🖇️ Vencimento de Documentos", callback_data="menu_vencimentos")],
        [InlineKeyboardButton("🔙 Voltar ao Início", callback_data="voltar_inicio")]
    ]
    await update.callback_query.message.edit_text(
        "🗒️ *Indicadores Mensais*\n\n"
        "Visualize os indicadores mensais consolidados de desempenho e posicionamento dos fornecedores no ranking por nota técnica.",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
async def menu_documentos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = [
        [InlineKeyboardButton("📂 Documentação CLAF", callback_data="menu_documentacao")],
        [InlineKeyboardButton("🖇️ Vencimento de Documentos", callback_data="menu_vencimentos")],
        [InlineKeyboardButton("🔙 Voltar ao Início", callback_data="voltar_inicio")]
    ]
    await update.callback_query.message.edit_text(
        "📂 *Documentação Oficial*\n\n"
        "Acesse os documentos oficiais da área de Suprimentos e acompanhe a validade das certificações e registros dos fornecedores:",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )

async def submenu_aprovados(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = [
        [InlineKeyboardButton("📂 Modo Individual", callback_data="aprovados_individual")],
        [InlineKeyboardButton("🗓️ Modo Mensal", callback_data="aprovados_mensal")],
        [InlineKeyboardButton("🔙 Voltar", callback_data="menu_desempenho")]
    ]
    await update.callback_query.message.edit_text(
        "✅ *Fornecedores Aprovados*\n\nSelecione abaixo como deseja visualizar os fornecedores com IQF acima de 75:",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
async def submenu_atencao(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = [
        [InlineKeyboardButton("📂 Modo Individual", callback_data="atencao_individual")],
        [InlineKeyboardButton("🗓️ Modo Mensal", callback_data="atencao_mensal")],
        [InlineKeyboardButton("🔙 Voltar", callback_data="menu_desempenho")]
    ]
    await update.callback_query.message.edit_text(
        "⚠️ *Fornecedores em Atenção*\n\nSelecione abaixo como deseja visualizar os fornecedores com IQF entre 70 e 75:",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
async def submenu_reprovados(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = [
        [InlineKeyboardButton("📂 Modo Individual", callback_data="reprovados_individual")],
        [InlineKeyboardButton("🗓️ Modo Mensal", callback_data="reprovados_mensal")],
        [InlineKeyboardButton("🔙 Voltar", callback_data="menu_desempenho")]
    ]
    await update.callback_query.message.edit_text(
        "❌ *Fornecedores Reprovados*\n\nSelecione abaixo como deseja visualizar os fornecedores com IQF inferior a 70:",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )

def gerar_menu_desempenho():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Fornecedores Aprovados", callback_data="grupo_aprovados")],
        [InlineKeyboardButton("⚠️ Fornecedores em Atenção", callback_data="grupo_atencao")],
        [InlineKeyboardButton("❌ Fornecedores Reprovados", callback_data="grupo_reprovados")],
        [InlineKeyboardButton("🔙 Voltar", callback_data="voltar_inicio")]
    ])

def gerar_menu_indicadores():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📈 Análise Mensal", callback_data="menu_mensal")],
        [InlineKeyboardButton("📊 Ranking Mensal", callback_data="menu_ranking")],
        [InlineKeyboardButton("📉 Tendência de Desempenho", callback_data="tendencia_0")],
        [InlineKeyboardButton("🔙 Voltar", callback_data="voltar_inicio")]
    ])

def gerar_menu_documentos():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🗂️ Documentação CLAF", callback_data="menu_documentacao")],
        [InlineKeyboardButton("🔍 Portal da Transparência (Sanções)", url="https://portaldatransparencia.gov.br/sancoes/consulta?ordenarPor=nomeSancionado&direcao=asc")],
        [InlineKeyboardButton("🖇️ Vencimento de Documentos", callback_data="menu_vencimentos")],
        [InlineKeyboardButton("🔙 Voltar", callback_data="voltar_inicio")]
    ])



async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    d = update.callback_query.data
    await update.callback_query.answer()
    nome = None

    # ─── NAVEGAÇÃO ENTRE MENUS ─────────────────────────────────────────────
    if d == "menu_desempenho":
        return await update.callback_query.message.edit_text(
            "📊 *Desempenho dos Fornecedores*\n\n"
           "Acompanhe o desempenho técnico dos fornecedores homologados, conforme critérios do Procedimento PG.SM.02. Selecione o grupo desejado para visualizar os resultados individuais.\n\n"
            "*Escolha uma das opções abaixo:*",
            reply_markup=gerar_menu_desempenho(),
            parse_mode="Markdown"
        )

    elif d == "menu_indicadores":
        return await update.callback_query.message.edit_text(
            "🗒️ *Indicadores Mensais*\n\n" 
            "Visualize os indicadores mensais consolidados de desempenho e posicionamento dos fornecedores no ranking por nota técnica.\n\n"
            "*Selecione uma análise:*",
            reply_markup=gerar_menu_indicadores(),
            parse_mode="Markdown"
        )

    elif d == "menu_documentos":
        return await update.callback_query.message.edit_text(
            "📁 *Documentos e Vencimentos*\n\n" 
            "Acesse os documentos oficiais da área de Suprimentos e acompanhe a validade das certificações e registros dos fornecedores.\n\n"
            "*Escolha uma opção:*",
            reply_markup=gerar_menu_documentos(),
            parse_mode="Markdown"
        )

    elif d == "voltar_inicio":
        return await start(update, context)

    # ─── PROCEDIMENTO ENGENMAN ─────────────────────────────────────────────
    elif d == "procedimento":
        return await handle_procedimento(update, context)


    elif d.startswith("ai_"):
        iden = d
        nome = fornecedor_id_map.get(iden)
        if nome:
            return await handle_feedback_individual(update, context, nome)
        else:
            await update.callback_query.message.reply_text("❌ Fornecedor não encontrado.")
            return



    elif d == "aprovados_individual":
        return await listar_aprovados_individual(update, context, page=0)
    elif d.startswith("aprovados_individual_"):
        page = int(d.split("_")[-1])
        return await listar_aprovados_individual(update, context, page=page)

    elif d == "atencao_individual":
        return await listar_atencao_individual(update, context, page=0)
    elif d.startswith("atencao_individual_"):
        page = int(d.split("_")[-1])
        return await listar_atencao_individual(update, context, page=page)

    elif d == "reprovados_individual":
        return await listar_reprovados_individual(update, context, page=0)
    elif d.startswith("reprovados_individual_"):
        page = int(d.split("_")[-1])
        return await listar_reprovados_individual(update, context, page=page)

    elif d == "aprovados_mensal":
        return await listar_aprovados_mensal(update, context)

    elif d == "atencao_mensal":
        return await listar_atencao_mensal(update, context)

    elif d == "reprovados_mensal":
        return await listar_reprovados_mensal(update, context)

    elif d == "grupo_aprovados":
        return await submenu_aprovados(update, context)

    elif d == "grupo_atencao":
         return await submenu_atencao(update, context)

    elif d == "grupo_reprovados":
        return await submenu_reprovados(update, context)

    query = update.callback_query
    data = query.data


    # Verifica se é callback do tipo aprovados
    if data.startswith("aprovados:"):
        try:
            _, mes, page = data.split(":")
            mes = mes.replace("-", "/")
            page = int(page)
            await listar_aprovados_por_mes(update, context, mes, page)
        except Exception as e:
            await query.message.edit_text("❌ Erro ao processar os dados.")
            print(f"Erro: {e}")
    
       
    
    # ─── FEEDBACK INDIVIDUAL: APROVADOS ─────────────────────────────────────
    elif d == "menu_feedback":
        df_qualidade = carregar_dados_qualidade()
        fornecedores = [
            nome for nome in df_qualidade['nome_agente'].unique()
            if df_qualidade[df_qualidade['nome_agente'] == nome]['nota'].mean() >= 75
        ]
        if not fornecedores:
            return await update.callback_query.message.reply_text("Nenhum fornecedor aprovado com IQF ≥ 75.")
        fornecedor_id_map.clear()
        fornecedor_id_map.update({f"f{i}": nome for i, nome in enumerate(sorted(fornecedores))})
        kb = [[InlineKeyboardButton(f, callback_data=f"feedback_f{i}")] for i, f in enumerate(sorted(fornecedores))]
        return await update.callback_query.message.edit_text("✅ Fornecedores aprovados:", reply_markup=InlineKeyboardMarkup(kb))

    # ─── FEEDBACK: EM ATENÇÃO ───────────────────────────────────────────────
    elif d == "menu_atencao":
        df_qualidade = carregar_dados_qualidade()
        fornecedores = [
            nome for nome in df_qualidade['nome_agente'].unique()
            if 70 <= df_qualidade[df_qualidade['nome_agente'] == nome]['nota'].mean() < 75
        ]
        if not fornecedores:
            return await update.callback_query.message.reply_text("Nenhum fornecedor em atenção no momento.")
        fornecedor_id_map.clear()
        fornecedor_id_map.update({f"f{i}": nome for i, nome in enumerate(sorted(fornecedores))})
        kb = [[InlineKeyboardButton(f, callback_data=f"feedback_f{i}")] for i, f in enumerate(sorted(fornecedores))]
        return await update.callback_query.message.edit_text("⚠️ Fornecedores em Estado de Atenção:", reply_markup=InlineKeyboardMarkup(kb))

    # ─── FEEDBACK: REPROVADOS ───────────────────────────────────────────────
    elif d == "menu_reprovados":
        df_qualidade = carregar_dados_qualidade()
        fornecedores = [
            nome for nome in df_qualidade['nome_agente'].unique()
            if df_qualidade[df_qualidade['nome_agente'] == nome]['nota'].mean() < 70
        ]
        if not fornecedores:
            return await update.callback_query.message.reply_text("Nenhum fornecedor reprovado no momento.")
        fornecedor_id_map.clear()
        fornecedor_id_map.update({f"f{i}": nome for i, nome in enumerate(sorted(fornecedores))})
        kb = [[InlineKeyboardButton(f, callback_data=f"feedback_f{i}")] for i, f in enumerate(sorted(fornecedores))]
        return await update.callback_query.message.edit_text("❌ Fornecedores Reprovados:", reply_markup=InlineKeyboardMarkup(kb))

    # ─── FEEDBACK INDIVIDUAL ESPECÍFICO ─────────────────────────────────────
    elif d.startswith("feedback_"):
        iden = d.split("_", 1)[1]
        return await handle_feedback_individual(update, context, fornecedor_id_map[iden])

    # ─── MENSAL: ANÁLISE E RANKING ──────────────────────────────────────────
    elif d == "menu_mensal":
        meses_disponiveis = obter_meses_disponiveis()
        kb = [[InlineKeyboardButton(m, callback_data=f"mes_{m.replace('/', '-')}")] for m in meses_disponiveis]
        return await update.callback_query.message.edit_text("📅 *Escolha o mês da análise:*", reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif d.startswith("mes_"):
        mes = d.split("_", 1)[1].replace("-", "/")
        return await handle_analise_mensal(update, context, mes)

    elif d == "menu_ranking":
        meses_disponiveis = obter_meses_disponiveis()
        kb = [[InlineKeyboardButton(m, callback_data=f"rank_{m.replace('/', '-')}")] for m in meses_disponiveis]
        return await update.callback_query.message.edit_text("🏆 *Escolha o mês do ranking:*", reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif d.startswith("rank_"):
        mes = d.split("_", 1)[1].replace("-", "/")
        return await handle_ranking(update, context, mes)
    

    

    # ─── DOCUMENTOS CLAF ─────────────────────────────────────────────────────
    elif d == "menu_documentacao":
        doc_paths = {
            "📗 CLAF - Critérios Legais.xlsx": r"Claf/CLAF.xlsx",
            "📘 Código de Ética.pdf": r"Claf/CÓDIGO DE ÉTICA ENGEMAN.pdf",
            "📝 Formulário de Avaliação.docx": r"Claf/FORM.407.REV00 - QUESTIONARIO DE QUALIFICAÇÃO DE FORNECEDOR.docx"
        }
        for nome_exibicao, caminho in doc_paths.items():
            if os.path.exists(caminho):
                try:
                    with open(caminho, 'rb') as file:
                        await update.callback_query.message.reply_document(
                            InputFile(file, filename=nome_exibicao),
                            caption="Clique para baixar o arquivo."
                        )
                except Exception as e:
                    await update.callback_query.message.reply_text(f"❌ Erro ao enviar: {e}")
            else:
                await update.callback_query.message.reply_text(f"⚠️ Arquivo não encontrado: {nome_exibicao}")
        return

    # ─── VENCIMENTOS ─────────────────────────────────────────────────────────
    elif d == "menu_vencimentos":
        return await handle_vencimentos_documentos(update, context)

    elif d.startswith("vencimento_"):
        return await handle_vencimento_por_mes(update, context)
    
    # SUPORTE E CONTATO DOS FORNECEDORES POR BASE ──────────────────────────────────────────────────────────
    elif d == "menu_suporte":
        kb = [
        [InlineKeyboardButton("👤 Responsável pela Base", callback_data="submenu_responsavel")],
        [InlineKeyboardButton("🗂 BASE", callback_data="submenu_base")],
        [InlineKeyboardButton("🔙 Voltar ao Início", callback_data="voltar_inicio")]
    ]
        return await update.callback_query.message.edit_text(
        "☎️ *Suporte e Contato*\n\nEscolha uma das opções para consultar as bases regionais e seus fornecedores:",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )

    elif d == "submenu_responsavel":
        texto = (
        "👤 *Responsáveis pelas Bases:*\n\n"
        "• *FILIAL – MACAÉ:\n\n*" 
        "*Compradores:* Anderson, Pryscilla & Rosane.\n\n"
        "• * FILIAL PERNAMBUCO – (RNEST, MATRIZ, PRESERV4):\n\n*" 
        "*Compradora:* Bruna.\n\n"
        "• *FILIAL – CEARÁ:*\n\n" 
        "*Compradores:* Iran.\n\n"
        "• *FILIAL SÃO PAULO – RPBC:*\n\n" 
        "*Comprador:* Gilberto."
    )
        return await update.callback_query.message.edit_text(texto, parse_mode="Markdown")
    

    elif d == "submenu_base":
        bases = [ "RJ", "SP", "PE", "PARACURU", "MG"]
        kb = [[InlineKeyboardButton(base, callback_data=f"base_{base}_0")] for base in bases]
        return await update.callback_query.message.edit_text(
        "🗂 *Selecione a base desejada:*",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
    elif re.match(r"^base_[A-Z]+_\d+$", d):
        base, pagina = d.split("_")[1], int(d.split("_")[2])
        df = pd.read_excel("DADOS DOS FORNECEDORES E COMPRADORES.xlsx", sheet_name="Fornecedores", skiprows=3)
        df.columns = ['IDX', 'FORNECEDOR', 'CONTRATO', 'CONTATO', 'FONE', 'EMAIL']
        df = df.drop(columns="IDX").dropna(subset=["FORNECEDOR"])
        df['CONTRATO'] = df['CONTRATO'].str.upper().str.strip()
        df['FORNECEDOR'] = df['FORNECEDOR'].str.strip()
        df['CONTRATO'] = df['CONTRATO'].apply(lambda x: base if isinstance(x, str) and base in x else None)
        fornecedores = df[df['CONTRATO'].notna()].sort_values(by="FORNECEDOR").reset_index(drop=True)

        total = len(fornecedores)
        por_pagina = 20
        inicio = pagina * por_pagina
        fim = inicio + por_pagina
        page_data = fornecedores.iloc[inicio:fim]

        if page_data.empty:
            return await update.callback_query.message.reply_text("⚠️ Nenhum fornecedor encontrado para essa base.")

        kb = [
            [InlineKeyboardButton(row['FORNECEDOR'], callback_data=f"forn_{base}_{inicio+i}")]
            for i, row in page_data.iterrows()
        ]

        nav = []
        if inicio > 0:
            nav.append(InlineKeyboardButton("⬅️ Anterior", callback_data=f"base_{base}_{pagina - 1}"))
        if fim < total:
            nav.append(InlineKeyboardButton("➡️ Próxima", callback_data=f"base_{base}_{pagina + 1}"))
        if nav:
            kb.append(nav)

        return await update.callback_query.message.edit_text(
            f"📋 *Fornecedores da base {base}* (pág. {pagina+1}):",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode="Markdown"
        )
    
    elif d.startswith("forn_"):
        _, base, idx = d.split("_")
        idx = int(idx)
        df = pd.read_excel("DADOS DOS FORNECEDORES E COMPRADORES.xlsx", sheet_name="Fornecedores", skiprows=3)
        df.columns = ['IDX', 'FORNECEDOR', 'CONTRATO', 'CONTATO', 'FONE', 'EMAIL']
        df = df.drop(columns="IDX").dropna(subset=["FORNECEDOR"])
        df['CONTRATO'] = df['CONTRATO'].str.upper().str.strip()
        df['FORNECEDOR'] = df['FORNECEDOR'].str.strip()
        df['BASE'] = df['CONTRATO'].str.extract(rf"({base})", expand=False)
        fornecedores = df[df['BASE'].notna()].sort_values(by="FORNECEDOR").reset_index(drop=True)
        if idx >= len(fornecedores):
            return await update.callback_query.message.reply_text("❌ Erro ao localizar fornecedor.")
        
        row = fornecedores.iloc[idx]

        responsaveis = {
            'RJ': "ANDERSON, ROSE E PRYSCILLA",
            'PE': "BRUNA",
            'CE': "IRAN",
            'SP': "GILBERTO",
            'PARACURU': "IRAN",
            'MG': "NÃO DEFINIDO"
        }
        responsavel = responsaveis.get(base.upper(), "NÃO DEFINIDO")

        msg = (
        f"🏷️ *Fornecedor:* {row['FORNECEDOR']}\n"
        f"📞 *Contato:* {row['CONTATO']}\n"
        f"📱 *Telefone:* {row['FONE']}\n"
        f"✉️ *E-mail:* {row['EMAIL']}\n\n"
        f"👤 *Responsável pela base:* {responsavel}"
    )
        
        return await update.callback_query.message.edit_text(msg, parse_mode="Markdown")   
    
    elif d.startswith("tendencia_"):
        page = int(d.split("_")[1])
        return await mostrar_lista_tendencia(update, context, page)

    elif d.startswith("trend_sel_"):
        codigo = d.replace("trend_sel_", "")
        nome = fornecedor_trend_map.get(codigo)
    if nome:
        await enviar_grafico_tendencia(update, context, nome)
    else:
        await context.bot.send_message(chat_id=update.effective_chat.id, text="❌ Código inválido.")

# GRÁFICO DE DESEMPENHO DO FORNECEDOR ──────────────────────────────────────────────────────────

async def mostrar_lista_tendencia(update: Update, context: ContextTypes.DEFAULT_TYPE, page: int):
    df = pd.read_excel("atendimento controle_qualidade.xlsx")
    fornecedores = sorted(df['nome_agente'].dropna().astype(str).str.strip().str.upper().unique())

    total_paginas = (len(fornecedores) - 1) // 50 + 1
    inicio = page * 50
    fim = inicio + 50
    fornecedores_pagina = fornecedores[inicio:fim]

    botoes = []
    for nome in fornecedores_pagina:
        codigo = sha1(nome.encode()).hexdigest()[:10]
        fornecedor_trend_map[codigo] = nome
        botoes.append([InlineKeyboardButton(nome, callback_data=f"trend_sel_{codigo}")])

    navegacao = []
    if page > 0:
        navegacao.append(InlineKeyboardButton("◀️ Anterior", callback_data=f"tendencia_{page - 1}"))
    if page < total_paginas - 1:
        navegacao.append(InlineKeyboardButton("Próxima ▶️", callback_data=f"tendencia_{page + 1}"))
    if navegacao:
        botoes.append(navegacao)

    await update.callback_query.message.edit_text(
        text="📉 *Selecione o fornecedor para ver a tendência de desempenho:*",
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode="Markdown"
    )

# Função: gerar e enviar o gráfico do fornecedor
async def enviar_grafico_tendencia(update: Update, context: ContextTypes.DEFAULT_TYPE, nome: str):
    df = pd.read_excel("atendimento controle_qualidade.xlsx")
    df['data'] = pd.to_datetime(df['data'], errors='coerce')
    df['nota'] = pd.to_numeric(df['nota'], errors='coerce')
    df['mes'] = df['data'].dt.month

    df['nome_agente'] = df['nome_agente'].astype(str).str.strip().str.upper()
    nome = nome.strip().upper()
    df_fornecedor = df[df['nome_agente'] == nome]

    if df_fornecedor.empty:
        await context.bot.send_message(chat_id=update.effective_chat.id, text="⚠️ Nenhum dado encontrado para este fornecedor.")
        return

    df_grouped = df_fornecedor.groupby('mes')['nota'].mean().reset_index()

    if df_grouped.empty:
        await context.bot.send_message(chat_id=update.effective_chat.id, text="⚠️ Não há notas disponíveis para gerar o gráfico.")
        return

    meses_ordem = ['janeiro','fevereiro','março','abril','maio','junho','julho',
                   'agosto','setembro','outubro','novembro','dezembro']
    df_grouped['mes_nome'] = df_grouped['mes'].apply(lambda x: meses_ordem[x - 1])

    # Criação do gráfico
    fig,ax = plt.subplots(figsize=(8,5))
    ax.plot(df_grouped["mes_nome"], df_grouped['nota'], marker='o', color="#FF7B00",linewidth=2.5)

    # RÓTULOS NOS PONTOS PRINCIPAIS
    for i, valor in enumerate(df_grouped['nota']):
        ax.text(df_grouped['mes_nome'][i], valor + 1.5, f"{valor:.1f}", ha='center', fontsize=9, color='gray')

    # TÍTULOS AJUSTADOS 
    titulo = f"Gráfico IQF - {nome}"
    titulo_quebrado = '\n'.join(wrap(titulo,width=50))
    ax.set_title(titulo_quebrado, fontsize=11, fontweight='bold', pad=20)

    # ESTÉTICA DE DADOS
    ax.set_xlabel('Mês', fontsize=10)
    ax.set_ylabel("Nota IQF", fontsize=10)
    ax.set_ylim(0,100)
    ax.grid(True, linestyle='--', alpha=0.5)
    plt.xticks(rotation=45)

    # AJUSTE NO LAYOUT

    fig.subplots_adjust(top=0.80)
    # EXPORTAÇÃO 
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    plt.close()     
    await context.bot.send_photo(chat_id=update.effective_chat.id, photo=buf) 
# Dispatcher de callbacks
async def dispatcher(update: Update, context: ContextTypes.DEFAULT_TYPE):
    d = update.callback_query.data

    if d == "indicadores":
        await menu_indicadores(update, context)

    elif d.startswith("tendencia_"):
        page = int(d.split("_")[1])
        await mostrar_lista_tendencia(update, context, page)

    elif d.startswith("trend_sel_"):
        nome = d.replace("trend_sel_", "")
        await enviar_grafico_tendencia(update, context, nome)

    elif d.startswith("aprovados:"):
        try:
            # Exemplo: aprovados:07-2025:page:1
            partes = d.split(":")
            mes = partes[1].replace("-", "/")
            page = int(partes[3]) if len(partes) == 4 and partes[2] == 'page' else 0
            await listar_aprovados_por_mes(update, context, mes, page)
        except Exception as e:
            await update.callback_query.message.edit_text("❌ Erro ao interpretar o callback.")
            print(f"[Dispatcher erro aprovados] {e}")

    elif d.startswith("ai_"):
        fornecedor = fornecedor_id_map.get(d)
        if fornecedor:
            await mostrar_detalhes_do_fornecedor(update, context, fornecedor)
        else:
            await update.callback_query.message.edit_text("❌ Fornecedor não encontrado.")


        
    # ENVIO DE E-MAIL
    if d.startswith("email_"):
        _, tipo, iden = d.split("_", 2)
        dest = []
        if tipo == "ind":
            dest = df_emails.loc[df_emails['Fornecedor'] == iden, 'E-mail'].dropna().tolist()
        else:
            dest = [SMTP_USER]
        if not dest:
            return await update.callback_query.message.reply_text("⚠️ E-mail do fornecedor não encontrado.")

        corpo = context.user_data.get('ultimo_texto', 'Segue a análise abaixo.')
        msg = EmailMessage()
        msg['Subject'] = f"Engeman – Análise {tipo.upper()} – {iden}"
        msg['From'] = SMTP_USER
        msg['To'] = dest[0]
        msg.set_content(montar_corpo_email(corpo))

        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
                s.starttls()
                s.login(SMTP_USER, SMTP_PASSWORD)
                s.send_message(msg)
            return await update.callback_query.message.reply_text(f"✅ E-mail enviado para {dest[0]}")
        except Exception as e:
            return await update.callback_query.message.reply_text(f"❌ Erro ao enviar o e-mail: {e}")

    if d == "nao_enviar":
        return await update.callback_query.message.reply_text("Envio por e-mail cancelado.")
    

# ENVIOS DE E-MAIL CONFIGURADO ──────────────────────────────────────────────────────────
async def handle_email_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if context.user_data.get('aguardando_email') != True:
        return  # Ignora se não estiver esperando e-mail

    email_destino = update.message.text.strip()
    if not re.match(r"[^@]+@[^@]+\.[^@]+", email_destino):
        return await update.message.reply_text("❌ E-mail inválido. Por favor, digite um e-mail válido.")

    path = context.user_data.get('arquivo_pdf')
    corpo = context.user_data.get('ultimo_texto')
    tipo = context.user_data.get('ultimo_tipo', 'análise')
    iden = context.user_data.get('ultimo_iden', '')

    # ─── TÍTULO DO E-MAIL ─────────────────────────────────────────────
    if tipo == 'men':
        mes_formatado = iden.replace('-', '/')
        titulo_email = f"ANÁLISE MENSAL ({mes_formatado})"
    elif tipo == 'ind':
        nomes = iden.strip().split()
        primeiro_nome_maiusculo = nomes[0].upper() if len(nomes) > 0 else ''
        titulo_email = f"FEEDBACK ENGEMAN – {primeiro_nome_maiusculo} ({datetime.now().year})"
    else:
        titulo_email = f"ANÁLISE ENGEMAN – {iden}"

    msg = EmailMessage()
    msg['Subject'] = titulo_email
    msg['From'] = SMTP_USER
    msg['To'] = email_destino
    msg.set_content("Este e-mail requer um cliente que suporte HTML.")

    # ─── CORPO DO E-MAIL ──────────────────────────────────────────────
    if tipo == 'ind':
        iqf_formatado = context.user_data.get('iqf_formatado', '0.00')
        msg.add_alternative(montar_corpo_email(corpo, iqf_formatado), subtype='html')
    elif tipo == 'men':
        msg.add_alternative(montar_corpo_email(corpo), subtype='html')
    else:
        msg.add_alternative(f"<html><body><p>{corpo.replace('\n','<br>')}</p></body></html>", subtype='html')

    # ─── ENVIO DO E-MAIL ──────────────────────────────────────────────
    try:
        if tipo == 'men':  # Apenas para análise mensal, anexa o PDF
            with open(path, 'rb') as f:
                msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(path))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
            s.starttls()
            s.login(SMTP_USER, SMTP_PASSWORD)
            s.send_message(msg)

        await update.message.reply_text(f"✅ E-mail enviado com sucesso para {email_destino}.")
    except Exception as e:
        await update.message.reply_text(f"❌ Erro ao enviar o e-mail: {e}")

    context.user_data['aguardando_email'] = False

def mes_ano_portugues(data):
    meses_pt = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    return f"{meses_pt[data.month]}/{data.year}"

# 🟦 Handler principal do botão "📇 Vencimentos de Fornecedores"
async def handle_vencimentos_documentos(update, context):
    query = update.callback_query
    await query.answer()

    df = carregar_dados_homologados()  # ← Sua função de carregamento deve retornar o DataFrame
    df['data vencimento'] = pd.to_datetime(df['data vencimento'], errors='coerce', dayfirst=True)
    df = df[df['data vencimento'].notna()]
    df = df[df['data vencimento'] >= pd.Timestamp.now()]
    df['mes_ano'] = df['data vencimento'].dt.strftime("%m/%Y")

    meses_disponiveis = sorted(
        df['mes_ano'].unique(),
        key=lambda m: datetime.strptime(m, "%m/%Y")
    )

    if not meses_disponiveis:
        await query.edit_message_text("⚠️ Nenhum vencimento futuro encontrado.")
        return

    botoes = [
        [InlineKeyboardButton(text=mes, callback_data=f"vencimento_{mes}")]
        for mes in meses_disponiveis
    ]

    await query.edit_message_text(
        "📅 *Selecione um mês para visualizar os vencimentos:*",
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode="Markdown"
    )

# ─── HANDLER PARA LISTAR OS VENCIMENTOS POR MÊS ──────────────────────────────
# 📌 Handler que exibe fornecedores com vencimento no mês escolhido
async def handle_vencimento_por_mes(update, context):
    query = update.callback_query
    await query.answer()

    mes_selecionado = query.data.replace("vencimento_", "")  # Ex: "05/2026"

    df = carregar_dados_homologados()
    df['data vencimento'] = pd.to_datetime(df['data vencimento'], errors='coerce', dayfirst=True)
    df = df[df['data vencimento'].notna()]
    df['mes_ano'] = df['data vencimento'].dt.strftime("%m/%Y")

    df_mes = df[df['mes_ano'] == mes_selecionado]
    if df_mes.empty:
        return await query.edit_message_text(f"⚠️ Nenhum fornecedor com vencimento em {mes_selecionado}.")

    mensagem = f"📆 *Vencimentos em {mes_selecionado}:*\n"
    for _, row in df_mes.iterrows():
        nome = row['agente']
        vencimento = row['data vencimento'].strftime("%d/%m/%Y")
        mensagem += f"• {nome} – vence em {vencimento}\n"

    mensagem += "\n🔎 *Verifique as documentações do fornecedor antes do prazo de validade.*"

    await query.edit_message_text(mensagem, parse_mode="Markdown")
    
async def listar_aprovados_individual(update: Update, context: ContextTypes.DEFAULT_TYPE, page: int = 0):
    df = carregar_dados_qualidade()
    df['nota'] = pd.to_numeric(df['nota'], errors='coerce')
    df['nome_agente'] = df['nome_agente'].astype(str).str.strip().str.upper()
    df['mes'] = pd.to_datetime(df['mes'], errors='coerce')
    df = df.dropna(subset=['mes'])  # remove linhas com datas inválidas
    df['mes'] = df['mes'].dt.strftime('%m/%Y')


    fornecedores = (
        df.groupby('nome_agente')['nota'].mean()
        .loc[lambda x: x > 75]
        .sort_index()
    )

    nomes = list(fornecedores.index)
    total = len(nomes)
    por_pagina = 50
    inicio = page * por_pagina
    fim = inicio + por_pagina

    if inicio >= total:
        return await update.callback_query.message.reply_text("❌ Página inválida.")

    botoes = [
        [InlineKeyboardButton(nome, callback_data=f"ai_{sha1(nome.encode()).hexdigest()}")]
        for nome in nomes[inicio:fim]
    ]

    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("⬅️ Anterior", callback_data=f"aprovados_individual_{page-1}"))
    if fim < total:
        nav.append(InlineKeyboardButton("➡️ Próxima", callback_data=f"aprovados_individual_{page+1}"))
    if nav:
        botoes.append(nav)

    for nome in nomes:
        fornecedor_id_map[f"ai_{sha1(nome.encode()).hexdigest()}"] = nome

    await update.callback_query.message.edit_text(
        f"✅ *Fornecedores Aprovados* (IQF > 75)\n\nPágina {page+1}",
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode="Markdown"
    )


async def listar_atencao_individual(update: Update, context: ContextTypes.DEFAULT_TYPE, page: int = 0):
    df = carregar_dados_qualidade()
    df['nota'] = pd.to_numeric(df['nota'], errors='coerce')
    df['nome_agente'] = df['nome_agente'].astype(str).str.strip().str.upper()
    df['mes'] = pd.to_datetime(df['mes'], errors='coerce')
    df = df.dropna(subset=['mes'])  # remove linhas com datas inválidas
    df['mes'] = df['mes'].dt.strftime('%m/%Y')


    fornecedores = (
        df.groupby('nome_agente')['nota'].mean()
        .loc[lambda x: (x >= 70) & (x <= 75)]
        .sort_index()
    )

    nomes = list(fornecedores.index)
    total = len(nomes)
    por_pagina = 50
    inicio = page * por_pagina
    fim = inicio + por_pagina

    if inicio >= total:
        return await update.callback_query.message.reply_text("❌ Página inválida.")

    botoes = [
        [InlineKeyboardButton(nome, callback_data=f"ai_{sha1(nome.encode()).hexdigest()}")]
        for nome in nomes[inicio:fim]
    ]

    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("⬅️ Anterior", callback_data=f"atencao_individual_{page-1}"))
    if fim < total:
        nav.append(InlineKeyboardButton("➡️ Próxima", callback_data=f"atencao_individual_{page+1}"))
    if nav:
        botoes.append(nav)

    for nome in nomes:
        fornecedor_id_map[f"ai_{sha1(nome.encode()).hexdigest()}"] = nome

    await update.callback_query.message.edit_text(
        f"⚠️ *Fornecedores em Atenção* (70 ≤ IQF ≤ 75)\n\nPágina {page+1}",
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode="Markdown"
    )


async def listar_reprovados_individual(update: Update, context: ContextTypes.DEFAULT_TYPE, page: int = 0):
    df = carregar_dados_qualidade()
    df['nota'] = pd.to_numeric(df['nota'], errors='coerce')
    df['nome_agente'] = df['nome_agente'].astype(str).str.strip().str.upper()
    df['mes'] = pd.to_datetime(df['mes'], format='%d/%m/%Y', errors='coerce').dt.strftime('%m/%Y')

    fornecedores = (
        df.groupby('nome_agente')['nota'].mean()
        .loc[lambda x: x < 70]
        .sort_index()
    )

    nomes = list(fornecedores.index)
    total = len(nomes)
    por_pagina = 50
    inicio = page * por_pagina
    fim = inicio + por_pagina

    if inicio >= total:
        return await update.callback_query.message.reply_text("❌ Página inválida.")

    botoes = [
        [InlineKeyboardButton(nome, callback_data=f"ai_{sha1(nome.encode()).hexdigest()}")]
        for nome in nomes[inicio:fim]
    ]

    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("⬅️ Anterior", callback_data=f"reprovados_individual_{page-1}"))
    if fim < total:
        nav.append(InlineKeyboardButton("➡️ Próxima", callback_data=f"reprovados_individual_{page+1}"))
    if nav:
        botoes.append(nav)

    for nome in nomes:
        fornecedor_id_map[f"ai_{sha1(nome.encode()).hexdigest()}"] = nome

    await update.callback_query.message.edit_text(
        f"❌ *Fornecedores Reprovados* (IQF < 70)\n\nPágina {page+1}",
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode="Markdown"
    )

async def listar_aprovados_mensal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    df = carregar_dados_qualidade()
    
    # Garantir que as colunas sejam convertidas corretamente
    df['nota'] = pd.to_numeric(df['nota'], errors='coerce')
    df['mes'] = pd.to_datetime(df['mes'], errors='coerce')
    df['nome_agente'] = df['nome_agente'].astype(str).str.strip().str.upper()

    # Remover registros com valores nulos em 'mes' ou 'nota'
    df = df.dropna(subset=['mes', 'nota'])

    # Obter os meses disponíveis
    meses_disponiveis = sorted(df['mes'].dt.strftime('%m/%Y').unique())
    
    # Verificar se existem meses disponíveis
    if not meses_disponiveis:
        await update.callback_query.message.edit_text("⚠️ Nenhum mês disponível para análise.")
        return

    # Criar os botões para cada mês
    botoes = [
        [InlineKeyboardButton(m, callback_data=f"mes_aprovados_{m.replace('/', '-')}")]
        for m in meses_disponiveis
    ]

    botoes.append([InlineKeyboardButton("🔙 Voltar", callback_data="grupo_aprovados")])

    # Enviar a mensagem com os botões
    await update.callback_query.message.edit_text(
        "✅ *Fornecedores Aprovados*\n\nSelecione o mês desejado para visualizar os fornecedores com IQF acima de 75:",
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode="Markdown"
    )

# Função que gera a página "Próxima" ou "Anterior"
def gerar_botoes_nav(mes_formatado, page, total, categoria):
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("⬅️ Anterior", callback_data=f"{categoria}_{mes_formatado.replace('/', '-')}_{page - 1}"))
    if page + 1 < total:
        nav.append(InlineKeyboardButton("➡️ Próxima", callback_data=f"{categoria}_{mes_formatado.replace('/', '-')}_{page + 1}"))
    return nav

async def listar_aprovados_por_mes(update: Update, context: ContextTypes.DEFAULT_TYPE, mes: str, page: int = 0):
    # Carrega os dados de qualidade
    df = carregar_dados_qualidade()

    # Convertendo a coluna 'nota' para numérica, se necessário
    df['nota'] = pd.to_numeric(df['nota'], errors='coerce')

    # Garantir que o mês no DataFrame seja formatado como 'MM/YYYY'
    df['mes'] = pd.to_datetime(df['mes'], errors='coerce').dt.strftime('%m/%Y')

    # Garantir que os nomes dos fornecedores sejam tratados
    df['nome_agente'] = df['nome_agente'].astype(str).str.strip().str.upper()

    # Filtra os dados pelo mês recebido no formato 'MM/YYYY'
    df_mes = df.loc[df['mes'] == mes]

    # Se não houver dados para o mês, exibe a mensagem de erro
    if df_mes.empty:
        await update.callback_query.message.edit_text(f"⚠️ Nenhum dado para {mes} (aprovados).")
        return

    # Filtra fornecedores com IQF > 75
    aprovados = df_mes.groupby('nome_agente')['nota'].mean()
    aprovados = aprovados[aprovados > 75].sort_index()

    # Obtém os fornecedores e a quantidade de fornecedores
    nomes = list(aprovados.index)
    total = len(nomes)
    por_pagina = 50  # Número de fornecedores por página
    inicio = page * por_pagina
    fim = inicio + por_pagina

    # Se o número da página estiver incorreto, exibe a mensagem de erro
    if inicio >= total:
        await update.callback_query.message.edit_text("❌ Página inválida.")
        return

    # Gera botões para os fornecedores
    botoes = [
        [InlineKeyboardButton(nome, callback_data=f"ai_{sha1(nome.encode()).hexdigest()}")]
        for nome in nomes[inicio:fim]
    ]

    # Navegação entre páginas (Próxima e Anterior)
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("⬅️ Anterior", callback_data=f"mes_aprovados_{mes.replace('/', '-')}_{page-1}"))
    if fim < total:
        nav.append(InlineKeyboardButton("➡️ Próxima", callback_data=f"mes_aprovados_{mes.replace('/', '-')}_{page+1}"))
    if nav:
        botoes.append(nav)

    botoes.append([InlineKeyboardButton("🔙 Voltar", callback_data="aprovados_mensal")])

    # Armazena os fornecedores e seus identificadores
    for nome in nomes:
        fornecedor_id_map[f"ai_{sha1(nome.encode()).hexdigest()}"] = nome

    # Exibe os fornecedores aprovados com base no mês
    await update.callback_query.message.edit_text(
        f"✅ *Fornecedores Aprovados em {mes}* (Página {page+1})",
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode="Markdown"
    )


async def listar_atencao_mensal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    df = carregar_dados_qualidade()
    df['nota'] = pd.to_numeric(df['nota'], errors='coerce')
    df['mes'] = pd.to_datetime(df['mes'].astype(str) + '/01', format='%m/%Y/%d', errors='coerce')
    df['mes_formatado'] = df['mes'].dt.strftime('%m/%Y')
    df['nome_agente'] = df['nome_agente'].astype(str).str.strip().str.upper()

    meses_disponiveis = sorted(df['mes'].dropna().unique())
    botoes = [
    [InlineKeyboardButton(m.strftime('%m/%Y'), callback_data=f"mes_atencao_{m.strftime('%m-%Y')}")]
    for m in meses_disponiveis
    ]
    botoes.append([InlineKeyboardButton("🔙 Voltar", callback_data="grupo_atencao")])

    await update.callback_query.message.edit_text(
        "⚠️ *Fornecedores em Atenção*\n\nSelecione o mês desejado para visualizar os fornecedores com IQF entre 70 e 75:",
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode="Markdown"
    )


async def listar_reprovados_mensal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    df = carregar_dados_qualidade()
    df['nota'] = pd.to_numeric(df['nota'], errors='coerce')
    df['mes'] = pd.to_datetime(df['mes'].astype(str) + '/01', format='%m/%Y/%d', errors='coerce')
    df['mes_formatado'] = df['mes'].dt.strftime('%m/%Y')
    df['nome_agente'] = df['nome_agente'].astype(str).str.strip().str.upper()

    meses_disponiveis = sorted(df['mes'].dropna().unique())
    botoes = [[InlineKeyboardButton(m, callback_data=f"mes_reprovados_{m}")] for m in meses_disponiveis]
    botoes.append([InlineKeyboardButton("🔙 Voltar", callback_data="grupo_reprovados")])

    await update.callback_query.message.edit_text(
        "❌ *Fornecedores Reprovados*\n\nSelecione o mês desejado para visualizar os fornecedores com IQF inferior a 70:",
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode="Markdown"
    )


# ─── INICIALIZAÇÃO E EXECUÇÃO ────────────────────────────────────────────────
async def main():
    nest_asyncio.apply()
    app = ApplicationBuilder().token(TOKEN_TELEGRAM).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(handle_vencimento_por_mes, pattern=r"^vencimento_\d{2}/\d{4}$"))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(CallbackQueryHandler(dispatcher))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_email_input))
    print("Bot iniciado!")
    await app.run_polling()

if __name__ == "__main__":
    import asyncio
    asyncio.run(main()) 