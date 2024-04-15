import tkinter as tk
from tkinter import ttk
from tkinter import *
import pandas as pd


def deletar_linha():
    df = pd.read_excel('produtos.xlsx')
    df = df.iloc[:-1]
    df.to_excel('produtos.xlsx', index=False)
    atualizar_tabela()

def create_table():
    
    nome_aluno = campo_nome.get()
    cep_aluno = campo_cep.get()
    cor_aluno = campo_cor.get()
    endereco_aluno = campo_endereco.get()
    numerero_casa_aluno = campo_numero_casa.get()
    complemento_casa = campo_complemento_casa.get()
    bairro_aluno = campo_bairro.get()
    estado_aluno = campo_estado.get()
    cidade_aluno = campo_cidade.get()
    email_aluno = campo_email.get()
    telefone_aluno = campo_telefone.get()

      # Adicionar dados ao DataFrame
    
    if nome_aluno and cep_aluno and cor_aluno and endereco_aluno \
    and numerero_casa_aluno and complemento_casa and bairro_aluno \
    and estado_aluno and cidade_aluno and email_aluno and telefone_aluno:
        
        
        df = pd.DataFrame({'Nome': [nome_aluno], 'Cep': [cep_aluno], 'Cor': [cor_aluno],
        'Endereço': [endereco_aluno],'Número': [numerero_casa_aluno],'Complemento':[complemento_casa],
        'Bairro': [bairro_aluno],'Estado':[estado_aluno],'Cidade':[cidade_aluno],
        'Email':[email_aluno],'Telefone':[telefone_aluno]})
        
        # Verificar se o arquivo Excel já existe
        try:
            existing_df = pd.read_excel('produtos.xlsx')
            df = pd.concat([existing_df, df], ignore_index=True)
        except FileNotFoundError:
            pass
        
        # Salvar o DataFrame no arquivo Excel
        df.to_excel('produtos.xlsx', index=False)
        
        # Atualizar a tabela na interface
        atualizar_tabela()
        limpar_campos()
    else :
        mensagem['text'] = 'Você deve preencher todos os campos'



def limpar_campos():
    campo_nome.delete(0, 'end')
    campo_cep.delete(0, 'end')
    campo_cor.set('')
    campo_endereco.delete(0, 'end')
    campo_numero_casa.delete(0, 'end')
    campo_complemento_casa.delete(0, 'end')
    campo_bairro.delete(0, 'end')
    campo_estado.delete(0, 'end')
    campo_cidade.delete(0, 'end')
    campo_email.delete(0, 'end')
    campo_telefone.delete(0, 'end')

# Função para atualizar a tabela na interface
def atualizar_tabela():
    try:
        df = pd.read_excel('produtos.xlsx')
        tree.delete(*tree.get_children())
        for index, row in df.iterrows():
            tree.insert('', 'end', values=(row['Nome'], row['Cep'], row['Cor'],
            row['Endereço'],row['Número'],row['Complemento'],row['Bairro'],row['Estado'],
            row['Cidade'], row['Email'], row['Telefone']))
    except FileNotFoundError:
        pass

            
janela = tk.Tk()
janela.title("Sistema de Cadastro de Alunos")
janela.columnconfigure([0,2,3], weight=1)

nome = tk.Label(text='Nome Social', anchor='e',height = 2)
nome.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)

campo_nome = tk.Entry(borderwidth=2)
campo_nome.grid(row=0, column=1, sticky='nsew', columnspan=3, padx=10, pady=10)

cep = tk.Label(text='CEP', anchor='e', height = 2)
cep.grid(row=1, column=0, sticky='nsew', padx=10, pady=10)

campo_cep = tk.Entry(borderwidth=2)
campo_cep.grid(row=1, column=1,padx=10, pady=10,sticky='nsew')

cor = tk.Label(text='Cor/Raça',anchor='e', height = 2)
cor.grid(row=1, column=2, sticky='nsew', padx=10, pady=10)

lista_cores = ['Preta', 
               'Branca', "Indigena", "Não declarado"]
campo_cor = ttk.Combobox(values=lista_cores)
campo_cor.grid(row=1, column=3, sticky='nsew', padx=10, pady=10)

endereco = tk.Label(text='Endereço', anchor='e', height = 2)
endereco.grid(row=2, column=0, sticky='nsew', padx=10, pady=10)

campo_endereco = tk.Entry(borderwidth=2)
campo_endereco.grid(row=2, column=1, sticky='nsew', padx=10, pady=10, columnspan=3)

numero_casa = tk.Label(text='Numero', anchor='e', height = 2)
numero_casa.grid(row=3, column=0, sticky='nsew', padx=10, pady=10)

campo_numero_casa = tk.Entry(borderwidth=2)
campo_numero_casa.grid(row=3, column=1, sticky='nsew', padx=10, pady=10)

completo_casa = tk.Label(text='Complemento', anchor='e',height = 2)
completo_casa.grid(row=3, column=2, sticky='nsew', padx=10, pady=10)

campo_complemento_casa = tk.Entry(borderwidth=2)
campo_complemento_casa.grid(row=3, column=3, sticky='nsew', padx=10, pady=10)

bairo = tk.Label(text='Bairro', anchor='e',height = 2)
bairo.grid(row=4, column=0, padx=10, pady=10, sticky='nsew')

campo_bairro = tk.Entry(borderwidth=2)
campo_bairro.grid(row=4, column=1, padx=10, pady=10, sticky='nwes', columnspan=3)

states = {
    'AC': 'Acre',
    'AL': 'Alagoas',
    'AP': 'Amapá',
    'AM': 'Amazonas',
    'BA': 'Bahia',
    'CE': 'Ceará',
    'DF': 'Distrito Federal',
    'ES': 'Espírito Santo',
    'GO': 'Goiás',
    'MA': 'Maranhão',
    'MT': 'Mato Grosso',
    'MS': 'Mato Grosso do Sul',
    'MG': 'Minas Gerais',
    'PA': 'Pará',
    'PB': 'Paraíba',
    'PR': 'Paraná',
    'PE': 'Pernambuco',
    'PI': 'Piauí',
    'RJ': 'Rio de Janeiro',
    'RN': 'Rio Grande do Norte',
    'RS': 'Rio Grande do Sul',
    'RO': 'Rondônia',
    'RR': 'Roraima',
    'SC': 'Santa Catarina',
    'SP': 'São Paulo',
    'SE': 'Sergipe',
    'TO': 'Tocantins'
}

lista_estados = [estado for estado in states.keys()]
estado = tk.Label(text='Estado', anchor='e',height = 2)
estado.grid(row=5, column=0, padx=10, pady=10, sticky='nsew')

campo_estado = ttk.Combobox(values=lista_estados)
campo_estado.grid(row=5,column=1, padx=10, pady=10,sticky='nswe')

cidade = tk.Label(text='Cidade', anchor='e',height = 2)
cidade.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

lista_cidades = [estado for estado in states.values()]
campo_cidade = ttk.Combobox(values=lista_cidades)
campo_cidade.grid(row=5, column=3, padx=10, pady=10, sticky='nsew')

email = tk.Label(text="Email", anchor='e',height = 2)
email.grid(row=6, column=0, padx=10, pady=10, sticky='nswe')

campo_email = tk.Entry(borderwidth=2)
campo_email.grid(row=6, column=1, padx=10, pady=10, sticky='nsew', columnspan=3)

telefone = tk.Label(text='Telefone', anchor='e',height = 2)
telefone.grid(row =7, column=0,padx=10, pady=10, sticky='nwes')


campo_telefone = tk.Entry(borderwidth=2)
campo_telefone.grid(row=7, column=1, padx=10, pady=10, sticky='nsew', columnspan=2)

mensagem = tk.Label(text='',anchor='e',height = 2,font= ("Helvetica", 12, "bold italic"),fg='crimson')
mensagem.grid(row=7, column=3, padx=10, pady=10, sticky='nsew')

botao_cadastra = tk.Button(text="Cadastrar",bg="lightblue", anchor='center',height = 2, width=5, command=create_table)
botao_cadastra.grid(row=8, column=1, padx=15, pady= 15, sticky='nswe')

botao_cancelar = tk.Button(text="Cancelar",bg='crimson', anchor='center',height = 1, width=1, command=deletar_linha)
botao_cancelar.grid(row=8, column=3, padx=15, pady= 15, sticky='nswe')



# Criar a tabela para exibir os produtos
tree = ttk.Treeview(janela, columns=('Nome', 'Cep', 'Cor','Endereço',
'Número','Complemento', 'Bairro','Estado','Cidade','Email','Telefone'), show='headings')
tree.heading('Nome', text='Nome')
tree.heading('Cep', text='Cep')
tree.heading('Cor', text='Cor')
tree.heading('Endereço', text='Endereço')
tree.heading('Número', text='Número')
tree.heading('Complemento', text='Complemento')
tree.heading('Bairro', text='Bairro')
tree.heading('Estado', text='Estado')
tree.heading('Cidade', text='Cidade')
tree.heading('Email', text='Email')
tree.heading('Telefone', text='Telefone')

tree.grid(row=9, column=0, columnspan=4, padx=5, pady=5,sticky=tk.NSEW)


# Definir a largura das colunas
tree.column('Nome', width=100,anchor='center')  
tree.column('Cep', width=40,anchor='center')  
tree.column('Cor', width=5,anchor='center')
tree.column('Endereço', width=125,anchor='center')
tree.column('Número', width=30,anchor='center')
tree.column('Bairro', width=90,anchor='center')
tree.column('Complemento', width=80,anchor='center')
tree.column('Complemento', width=110,anchor='center')
tree.column('Estado', width=50,anchor='center')
tree.column('Cidade', width=80,anchor='center')
tree.column('Email', width=150,anchor='center')
tree.column('Telefone', width=90,anchor='center')



vsb = ttk.Scrollbar(janela, orient="vertical", command=tree.yview)
vsb.grid(row=9, column=4, sticky='ns')
tree.configure(yscrollcommand=vsb.set)

# Atualizar a tabela na inicialização
atualizar_tabela()

janela.mainloop()

