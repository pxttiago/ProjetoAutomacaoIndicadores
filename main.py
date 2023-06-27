# importação das bibliotecas
import pandas as pd
import pathlib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


# 1 - IMPORTAÇÃO E TRATAMENTO DAS BASES DE DADOS

# importar as bases de dados
emails = pd.read_excel('Bases de Dados/Emails.xlsx')  # importa a base de dados para o Python
print(emails.head())
lojas = pd.read_csv('Bases de Dados/Lojas.csv', encoding='latin-1', sep=';')  # importa a base de dados para o Python
print(lojas.head())
vendas = pd.read_excel('Bases de Dados/Vendas.xlsx')  # importa a base de dados para o Python
print(vendas.head())

# 2 - CRIAR UM ARQUIVO/TABELA PARA CADA LOJA

# incluir nomes das lojas na tabela vendas
vendas = vendas.merge(lojas, on='ID Loja')  # tabela = tabela.merge(tabela_juntar,on='coluna_juntar')
print(vendas.head())

# criando as tabelas
dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]  # dicionario onde cada item é a tabela de vendas de uma das loja

print(dicionario_lojas['Ribeirão Shopping'])
print(dicionario_lojas['Iguatemi Esplanada'])

# definir o dia do indicador
dia_indicador = vendas['Data'].max()  # pegando a maior data da tabela de vendas
print(dia_indicador)
print('{}/{}'.format(dia_indicador.day, dia_indicador.month))

# 3 - SALVAR OS BACKUPS NAS PASTAS

# identificar se a pasta ja existe
caminho_pasta = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_pasta.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]
    
for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_pasta / loja
        nova_pasta.mkdir()

    # salvar dentro da pasta
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    local_arquivo = caminho_pasta / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)

# 4 - CALCULAR OS INDICADORES E ENVIAR O ONEPAGE AOS GERENTES DE CADA LOJA

# definição de metas
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

for loja in dicionario_lojas:

    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

    # faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    # diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    # ticket médio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    print(ticket_medio_ano)

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    print(ticket_medio_dia)

    # envio do e-mail via smtp
    email_remetente = 'SEU_EMAIL'
    senha = 'SUA_SENHA'

    destinatario = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]  # .values[0] retorna como resposta apenas o valor de loc[] ao invés de uma tabela
    destinatario_copia = ''
    destinatario_copia_oculta = ''
    nome_gerente = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]

    # Configurações do servidor SMTP
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = email_remetente
    smtp_password = senha

    # Construir o objeto do e-mail
    msg = MIMEMultipart()
    msg['From'] = email_remetente
    msg['To'] = destinatario
    msg['Subject'] = 'OnePage {}/{} - Loja {}'.format(dia_indicador.day, dia_indicador.month, loja)
    msg['Cc'] = destinatario_copia  # Destinatário CC
    msg['Bcc'] = destinatario_copia_oculta  # Destinatário BCC

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'

    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'

    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'

    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'

    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    # Corpo do e-mail
    body = f'''
    <p>Bom dia, {nome_gerente}.</p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>loja {loja}</strong> foi: </p>

    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
    </tr>
    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
    </tr>
    </table>
    <br>
    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
    </tr>
    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
    </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>
    <p>Qualquer dúvida estou à disposição</p>
    <p>Att., SeuNome</p>
    '''
    msg.attach(MIMEText(body, 'html')) # 'plain' tipo de conteúdo = texto simples / 'html' tipo de conteúdo = html

    # Anexo
    filename = pathlib.Path.cwd() / caminho_pasta / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    attachment = open(str(filename), 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')

    msg.attach(part)

    # Conectar ao servidor SMTP e enviar o e-mail
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(smtp_username, [destinatario, destinatario_copia, destinatario_copia_oculta], msg.as_string())
        server.quit()
        print('E-mail da loja {} enviado com sucesso!'.format(loja))
    except Exception as e:
        print('Ocorreu um erro ao enviar o e-mail:', str(e)) # 'str(e)' converte o objeto de exceção 'e' em uma representação de string


# 5 - CRIAR RANKING DAS LOJAS PARA A DIRETORIA E ENVIAR POR E-MAIL

# gerar os arqruivos com o ranking das lojas
faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

vendas_dia = vendas.loc[vendas['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
local_arquivo = caminho_pasta / nome_arquivo
faturamento_lojas_ano.to_excel(caminho_pasta)

nome_arquivo = '{}_{}_Ranking Diário.xlsx'.format(dia_indicador.month, dia_indicador.day)
local_arquivo =caminho_pasta / nome_arquivo
faturamento_lojas_dia.to_excel(caminho_pasta)

# envio do e-mail via smtp
    
email_remetente = ''
senha = ''

destinatario = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
destinatario_copia = ''
destinatario_copia_oculta = ''

# Configurações do servidor SMTP
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = email_remetente
smtp_password = senha

# Construir o objeto do e-mail
msg = MIMEMultipart()
msg['From'] = email_remetente
msg['To'] = destinatario
msg['Subject'] = 'Ranking Lojas {}/{}'.format(dia_indicador.day, dia_indicador.month)
msg['Cc'] = destinatario_copia  # Destinatário CC
msg['Bcc'] = destinatario_copia_oculta  # Destinatário BCC

# Corpo do e-mail
body = f'''
<p>Prezados, bom dia.</p>

<p>Segue em anexo a planilha com o ranking de faturamento diário e anual de todas as lojas.</p>

<p>Maior faturamento do dia: Loja {faturamento_lojas_dia.index[0]} - Faturamento: R$ {faturamento_lojas_dia.iloc[0, 0]:.2f}.</p
<p>Pior faturamento do dia: Loja {faturamento_lojas_dia.index[-1]} - Faturamento: R$ {faturamento_lojas_dia.iloc[-1, 0]:.2f}.</p
<p>Maior faturamento do ano: Loja {faturamento_lojas_ano.index[0]} - Faturamento: R$ {faturamento_lojas_ano.iloc[0, 0]:.2f}.</p
<p>Pior faturamento do ano: Loja {faturamento_lojas_ano.index[-1]} - Faturamento: R$ {faturamento_lojas_ano.iloc[-1, 0]:.2f}.</p

<p>Qualquer dúvida estou à disposição</p>
<p>Att., Tiago</p>
'''
msg.attach(MIMEText(body, 'plain'))  # 'plain' tipo de conteúdo = texto simples / 'html' tipo de conteúdo = html

# Anexo
filename = pathlib.Path.cwd() / caminho_pasta / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
attachment = open(str(filename), 'rb')
filename = pathlib.Path.cwd() / caminho_pasta / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Diário.xlsx'
attachment = open(str(filename), 'rb')

part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', f'attachment; filename="{filename}"')

msg.attach(part)

# Conectar ao servidor SMTP e enviar o e-mail
try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(smtp_username, smtp_password)
    server.sendmail(smtp_username, [destinatario, destinatario_copia, destinatario_copia_oculta], msg.as_string())
    server.quit()
    print('E-mail enviado com sucesso à diretoria!'.format(loja))
except Exception as e:
    print('Ocorreu um erro ao enviar o e-mail:', str(e)) # 'str(e)' converte o objeto de exceção 'e' em uma representação de string