"""
Tenho uma planilha do excel com dados dos alunos que finalizaram o curso e receberão o certificado.

Crie uma automação em python para enviar dados da planilha para preencher os dados mutáveis no certificado padrão

Dados dos alunos:
1 - Nome do curso
2 - Nome do participante
3 - Tipo de participação
4 - Data de início
5 - Data do final 
6 - Carga horária 
7 - Data de emissão do certficado

Tarefas: 
- Passe os dados para a imagem do certificado
-

"""

from PIL import ImageFont, Image, ImageDraw
import openpyxl

# Abrindo a planilha 
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

# Irá ler linhas desejadas
for indice, linha in enumerate (sheet_alunos.iter_rows(min_row=2)):
    # Cada coluna de dados dos alunos que precisamos
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value
    input('')

    # Transferir os dados da planilha para a imagem do certificado
    # Definindo a fonte dos dados
    fonte_nome = ImageFont.truetype('./tahomabd.ttf',90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf',80)
    fonte_data = ImageFont.truetype('./tahoma.ttf',40)

    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((1020,826), nome_participante, fill='black', font=fonte_nome)
    desenhar.text((1072,956), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1438,1068), tipo_participacao, fill='black', font=fonte_geral)
    desenhar.text((1490,1187), str(carga_horaria), fill='black', font=fonte_geral)

    desenhar.text((780, 1799), data_inicio, fill='blue', font=fonte_data)
    desenhar.text((780, 1950), data_final, fill='blue', font=fonte_data)

    desenhar.text((2260, 1950), data_emissao, fill= 'blue', font=fonte_data)

    image.save(f'./{indice}{nome_participante} certificado.png')

