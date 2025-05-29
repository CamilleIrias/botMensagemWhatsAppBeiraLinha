import pandas as pd
from IPython.display import display

import urllib.parse
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import urllib
import time
import pandas as pd
from IPython.display import display


class Curso:
    def __init__(self,nome, data_inicio, horario, convite_comunidade):
        self.nome = nome
        self.data_inicio = data_inicio
        self.horario = horario
        self.convite_comunidade = convite_comunidade

    def __repr__(self):
        return f"{self.nome} {self.horario} - Início: {self.data_inicio}"

class Aluno:
    def __init__(self, nome, sobrenome, telefone):
        self.nome = nome
        self.sobrenome = sobrenome
        self.telefone = telefone
        self.cursos = []

    def adicionar_curso(self, curso):
        self.cursos.append(curso)

    def __repr__(self):
        return f"{self.nome} {self.sobrenome} ({self.numero_formatado})"
    
    def mostrar_cursos(self):
        print(f"\nCursos de {self.nome} {self.sobrenome}:")
        for curso in sorted(self.cursos, key=lambda c: c.data_inicio):
            print(curso)


def criar_alunos_cursos(df):
    alunos_dic = {}

    for i, row in df.iterrows():
        nome = row['Nome']
        sobrenome = row['Sobrenome']
        numero_formatado = row['Numero Formatado']
        curso_nome = row['Tipo de ingresso']
        data_inicio = row['Data de Inicio']
        horario = row['Horario']
        convite_comunidade = row['Convite Comunidade']

        if(nome, sobrenome, numero_formatado) not in alunos_dic:
            aluno = Aluno(nome, sobrenome, numero_formatado)
            alunos_dic[(nome, sobrenome, numero_formatado)] = aluno
        else:
            aluno = alunos_dic[(nome, sobrenome, numero_formatado)]

        curso = Curso(curso_nome, data_inicio, horario, convite_comunidade)
        aluno.adicionar_curso(curso)
    
    return list(alunos_dic.values())

insc_df = pd.read_excel("Inscricoes0204PRESENCIAL_FINALIZADA (3).xlsx")

insc_df['Nome'] = insc_df['Nome'].str.title()
insc_df['Sobrenome'] = insc_df['Sobrenome'].str.title()

insc_resum_df = insc_df[['Ordem de inscrição', 'Nome', 'Sobrenome', 'Tipo de ingresso','Horario','Data de Inicio' , 'Numero Formatado', 'Convite Comunidade']]

alunos = criar_alunos_cursos(insc_resum_df)

nav = webdriver.Chrome()

nav.get("https://web.whatsapp.com/")

while len(nav.find_elements(By.ID, "side")) < 1:
    time.sleep(1)

for aluno in alunos:
    nome = aluno.nome
    sobrenome = aluno.sobrenome
    cursos = aluno.cursos
    telefone = aluno.telefone

    mensagem = f"Olá {nome} {sobrenome}, boa tarde!\n\n"
    mensagem += "Você foi inscrito no Projeto Beira Linha da PUC São Gabriel.\n"
    mensagem += "Estarei enviando os cursos e o convite para participar da comunidade Alunos - Beira Linha. \n"
    mensagem += "Está comunidade servirá para facilitar a comunicação do aluno com os monitores dos cursos e para o envio de links das aulas e avisos.\n"
    mensagem += "*Por favor entre na comunidade para não perder nenhum aviso!*\n\n"

    if len(cursos) > 1:
        mensagem += "Os cursos em que você foi inscrito são:\n"
        for curso in cursos:
            mensagem += f"- *{curso.nome} {curso.horario}*, na modalidade *presencial*, e {curso.data_inicio}. \n"
            mensagem += f"*Convite do grupo:* {curso.convite_comunidade}\n"
    else:
        curso = cursos[0]
        mensagem += f"O curso em que você foi inscrito é:\n- *{curso.nome} {curso.horario}*, na modalidade *presencial*, e {curso.data_inicio}. "
        mensagem += f"*Convite do grupo:* {curso.convite_comunidade}\n"
    mensagem += "\nCaso tenha alguma dúvida, pode entrar em contato comigo ou pelo e-mail edwaldo@pucminas.br"
    mensagem += "\n\n*OBS: Caso este telefone seja do responsável do aluno favor informar para atualizarmos com o contato do aluno (se ele possuir telefone)*"

    print(mensagem)
    mensagem = urllib.parse.quote(mensagem)

    link = f"https://web.whatsapp.com/send?phone={telefone}&text={mensagem}"
    nav.get(link)
    while len(nav.find_elements(By.ID, "side")) < 1:
        time.sleep(10)
    nav.find_element(By.XPATH, value='//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div/p/span').send_keys(Keys.ENTER)
    time.sleep(30)

