from tkinter import *
from tkinter import ttk
from openpyxl import Workbook, load_workbook
import os
from datetime import datetime

# Função para calcular o custo total de impressão 3D
def calcular():
    # Obtém os valores de entrada do usuário e converte para float
    peso_peca = float(var_pesoFatiador.get().replace(',', '.'))
    preco_fil = float(var_precoFilaKG.get().replace(',', '.'))
    preco_fix = float(var_precoFix.get().replace(',', '.'))
    
    # Calcula o custo dos materiais (filamento + fixador)
    total_mat = preco_fil / 1000 * peso_peca + preco_fix * 0.05

    # Calcula o custo de energia com base no preço do kWh e o tempo de impressão
    preco_kwh = float(var_precoEner.get().replace(',', '.'))
    tempo_imp = float(var_tempImpressao.get().replace(',', '.'))
    potencia_impressora = 500 # Potência da impressora em Watts
    total_ene = (potencia_impressora / 1000 * 0.6) * tempo_imp * preco_kwh

    # Calcula o valor de venda com base no lucro, finalização e falhas
    tx_lucro = float(var_lucro.get().replace(',', '.'))
    tx_finalizao = float(var_finalizacao.get().replace(',', '.'))
    tx_falhas = float(var_falhas.get().replace(',', '.'))
    valor_Venda = (total_ene + total_mat) * (tx_lucro / 100)
    valor_Venda = valor_Venda + valor_Venda * (tx_falhas / 100) + valor_Venda * (tx_falhas / 100)

    # Atualiza as variáveis de saída com os valores calculados
    var_materiais.set(str(round(total_mat, 2)))
    var_energia.set(str(round(total_ene, 2)))
    var_total.set(str(round(valor_Venda, 2)))

# Função para limpar os campos de entrada
def limpar():
    var_pesoFatiador.set('0')
    var_tempImpressao.set('0')
    var_materiais.set('0')
    var_energia.set('0')
    var_total.set('0')

# Função para abrir uma nova janela e registrar um orçamento
def registrar():
    btn_registrar['state'] = 'disabled'  # Desabilita o botão de registrar
    # Cria uma nova janela para inserir o nome e telefone do cliente
    janela2 = Toplevel()
    janela2.geometry('500x130')
    janela2.title("Registrar Orçamento")
    janela2.resizable(FALSE, FALSE)

    # Variáveis para armazenar o nome e telefone do cliente
    nome_cliente = StringVar()
    telefone_cliente = StringVar()

    # Criação dos widgets (rótulos e caixas de entrada) para o nome e telefone do cliente
    Label(janela2, text='Nome Cliente:', font='Arial 10').grid(row=0, column=0, padx=10, pady=10, sticky='e')
    Entry(janela2, textvariable=nome_cliente, font='Arial 10', width=50).grid(row=0, column=1, padx=10, pady=10, sticky='we')

    Label(janela2, text='Telefone Cliente:', font='Arial 10').grid(row=1, column=0, padx=10, pady=10, sticky='e')
    Entry(janela2, textvariable=telefone_cliente, font='Arial 10', width=50).grid(row=1, column=1, padx=10, pady=10, sticky='we')

    # Botão para salvar o registro
    btn_salvar = Button(janela2, text='Salvar', font='Arial 10 bold', command=lambda: salvar(nome_cliente.get(), telefone_cliente.get(), janela2), bg='#28a745', fg='white')
    btn_salvar.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky='WE')

    janela2.mainloop()  # Inicia o loop da nova janela

# Função para obter o caminho completo do arquivo Excel onde os dados serão salvos
def get_dir():
    caminho_completo = os.path.abspath(__file__)  # Obtém o caminho absoluto deste arquivo
    caminho_sem_dois_ultimos = os.path.dirname(os.path.dirname(caminho_completo))  # Remove os dois últimos diretórios do caminho
    caminho_excel = os.path.join(caminho_sem_dois_ultimos, 'docs', 'Orçamentos.xlsx')  # Define o caminho para o arquivo Excel
    print(caminho_excel)
    return caminho_excel  # Retorna o caminho do arquivo Excel

# Função para salvar os dados no arquivo Excel
def salvar(nome, telefone, janela):
    # Coletar os dados do orçamento
    dados = {
        'Data Orçamento': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # Data e hora atual
        'Nome': nome,  # Nome do cliente
        'Telefone': telefone,  # Telefone do cliente
        'Preço kWh': var_precoEner.get().replace('.', ','),  # Preço do kWh
        'Tempo Impressão': var_tempImpressao.get().replace('.', ','),  # Tempo de impressão
        'Energia': var_energia.get().replace('.', ','),  # Custo da energia
        'Preço Filamento (kg)': var_precoFilaKG.get().replace('.', ','),  # Preço do filamento por kg
        'Peso Impressão': var_pesoFatiador.get().replace('.', ','),  # Peso da impressão
        'Valor Fixador': var_precoFix.get().replace('.', ','),  # Valor do fixador
        'Materiais': var_materiais.get().replace('.', ','),  # Custo dos materiais
        'Lucro': var_lucro.get().replace('.', ','),  # Porcentagem de lucro
        'Finalização': var_finalizacao.get().replace('.', ','),  # Porcentagem de finalização
        'Falhas': var_falhas.get().replace('.', ','),  # Porcentagem de falhas
        'Total': var_total.get().replace('.', ',')  # Custo total
    }

    # Carregar o arquivo Excel e a planilha ativa
    wb = load_workbook(get_dir())
    ws = wb.active

    # Adicionar uma nova linha com os dados coletados
    ws.append([dados['Data Orçamento'], dados['Nome'], dados['Telefone'], dados['Preço kWh'], dados['Tempo Impressão'], dados['Energia'], dados['Preço Filamento (kg)'], dados['Peso Impressão'], dados['Valor Fixador'], dados['Materiais'], dados['Lucro'], dados['Finalização'], dados['Falhas'], dados['Total']])

    # Salvar o arquivo Excel
    wb.save(get_dir())

    # Habilitar o botão de registrar novamente e fechar a janela
    btn_registrar['state'] = 'normal'
    janela.destroy()

# Configuração da janela principal do Tkinter
janela = Tk()
janela.geometry('960x400')  # Define o tamanho da janela
janela.title('Orçamento de Custo de Impressão 3D')  # Título da janela
janela.configure(bg='#d9d9d9')  # Cor de fundo da janela
janela.resizable(FALSE, FALSE)  # Desabilita o redimensionamento da janela

# Configuração de estilo para os widgets usando ttk
style = ttk.Style()
style.configure('TLabel', font=('Arial', 10))
style.configure('TEntry', font=('Arial', 10))
style.configure('TButton', font=('Arial', 10, 'bold'), background='green', foreground='white')

# --------- GASTOS MATERIAIS E ENERGÉTICOS -----------
# Frame de gastos materiais
lf_mat = LabelFrame(janela, text='Gastos Materiais', font=('Arial', 12, 'bold'), bg='#d9d9d9', padx=10, pady=10)
lf_mat.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
lf_mat.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

# Entrada de preço do kWh
var_precoEner = StringVar(value='1.03')
Label(lf_mat, text='Preço kW/h:', font='Arial 10', bg='#d9d9d9').grid(row=0, column=0, padx=10, pady=10, sticky='e')
Entry(lf_mat, textvariable=var_precoEner, font='Arial 10').grid(row=0, column=1, padx=10, pady=10, sticky='ew')

# Entrada de preço do filamento por kg
var_precoFilaKG = StringVar(value='75.00')
Label(lf_mat, text='Preço Filamento (Kg):', font='Arial 10', bg='#d9d9d9').grid(row=0, column=2, padx=10, pady=10, sticky='e')
Entry(lf_mat, textvariable=var_precoFilaKG, font='Arial 10').grid(row=0, column=3, padx=10, pady=10, sticky='ew')

# Entrada de preço do fixador
var_precoFix = StringVar(value='25.00')
Label(lf_mat, text='Preço Fixador:', font='Arial 10', bg='#d9d9d9').grid(row=0, column=4, padx=10, pady=10, sticky='e')
Entry(lf_mat, textvariable=var_precoFix, font='Arial 10').grid(row=0, column=5, padx=10, pady=10, sticky='ew')

# Entrada de peso da peça
var_pesoFatiador = StringVar(value='30')
Label(lf_mat, text='Peso Peça (g):', font='Arial 10', bg='#d9d9d9').grid(row=1, column=0, padx=10, pady=10, sticky='e')
Entry(lf_mat, textvariable=var_pesoFatiador, font='Arial 10').grid(row=1, column=1, padx=10, pady=10, sticky='ew')

# Entrada de tempo de impressão
var_tempImpressao = StringVar(value='2.5')
Label(lf_mat, text='Tempo Impressão (h):', font='Arial 10', bg='#d9d9d9').grid(row=1, column=2, padx=10, pady=10, sticky='e')
Entry(lf_mat, textvariable=var_tempImpressao, font='Arial 10').grid(row=1, column=3, padx=10, pady=10, sticky='ew')



# --------- OUTROS GASTOS ----------- 
# Cria um frame rotulado 'Outros' para agrupar campos de entrada de dados relacionados a lucros, finalização e falhas
lf_outros = LabelFrame(janela, text='Outros', font=('Arial', 12, 'bold'), bg='#d9d9d9', padx=10, pady=10)
lf_outros.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")  # Define a posição e o preenchimento do frame na janela

# Configura as colunas do frame para que o conteúdo se ajuste dinamicamente
lf_outros.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

# Campo de entrada para o lucro (percentual)
var_lucro = StringVar(value='200.00')
Label(lf_outros, text='Lucro (%)', font='Arial 10', bg='#d9d9d9').grid(row=0, column=0, padx=10, pady=10, sticky='e')
Entry(lf_outros, textvariable=var_lucro, font='Arial 10').grid(row=0, column=1, padx=10, pady=10, sticky='ew')

# Campo de entrada para a finalização (percentual)
var_finalizacao = StringVar(value='20.00')
Label(lf_outros, text='Finalização (%)', font='Arial 10', bg='#d9d9d9').grid(row=0, column=2, padx=10, pady=10, sticky='e')
Entry(lf_outros, textvariable=var_finalizacao, font='Arial 10').grid(row=0, column=3, padx=10, pady=10, sticky='ew')

# Campo de entrada para falhas (percentual)
var_falhas = StringVar(value='10.00')
Label(lf_outros, text='Falhas (%)', font='Arial 10', bg='#d9d9d9').grid(row=0, column=4, padx=10, pady=10, sticky='e')
Entry(lf_outros, textvariable=var_falhas, font='Arial 10').grid(row=0, column=5, padx=10, pady=10, sticky='ew')

# --------- TOTAIS ----------- 
# Cria um frame rotulado 'Totais (R$)' para exibir os valores calculados de materiais, energia e total
lf_totais = LabelFrame(janela, text='Totais (R$)', font=('Arial', 12, 'bold'), bg='#d9d9d9', padx=10, pady=10)
lf_totais.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")  # Define a posição e o preenchimento do frame na janela

# Configura as colunas do frame para que o conteúdo se ajuste dinamicamente
lf_totais.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

# Campo de saída para exibir o custo dos materiais
var_materiais = StringVar()
Label(lf_totais, text='Materiais: ', font='Arial 10', bg='#d9d9d9').grid(row=0, column=0, padx=10, pady=10, sticky='e')
Entry(lf_totais, textvariable=var_materiais, font='Arial 10', state='readonly').grid(row=0, column=1, padx=10, pady=10, sticky='ew')

# Campo de saída para exibir o custo da energia
var_energia = StringVar()
Label(lf_totais, text='Energia:', font='Arial 10', bg='#d9d9d9').grid(row=0, column=2, padx=10, pady=10, sticky='e')
Entry(lf_totais, textvariable=var_energia, font='Arial 10', state='readonly').grid(row=0, column=3, padx=10, pady=10, sticky='ew')

# Campo de saída para exibir o total calculado
var_total = StringVar()
Label(lf_totais, text='Total:', font='Arial 10', bg='#d9d9d9').grid(row=0, column=4, padx=10, pady=10, sticky='e')
Entry(lf_totais, textvariable=var_total, font='Arial 10', state='readonly').grid(row=0, column=5, padx=10, pady=10, sticky='ew')

# Botão para calcular os custos com base nas entradas
btn_calcular = Button(lf_totais, text='CALCULAR', font='Arial 10 bold', bg='#28a745', fg='white', command=calcular)
btn_calcular.grid(row=1, column=0, padx=10, pady=10, sticky='we', columnspan=2)

# Botão para registrar o orçamento no arquivo Excel
btn_registrar = Button(lf_totais, text='REGISTRAR', font='Arial 10 bold', bg='#dc3545', fg='white', command=registrar)
btn_registrar.grid(row=1, column=2, padx=10, pady=10, sticky='we', columnspan=2)

# Botão para limpar os campos de entrada
btn_limpar = Button(lf_totais, text='LIMPAR', font='Arial 10 bold', bg='#dc3545', fg='white', command=limpar)
btn_limpar.grid(row=1, column=4, padx=10, pady=10, sticky='we', columnspan=2)

# Inicia o loop principal da interface gráfica para manter a janela aberta
janela.mainloop()

