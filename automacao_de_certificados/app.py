import openpyxl
from PIL import Image, ImageDraw, ImageFont

#abrir a planilha 
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice,linha in enumerate(sheet_alunos.iter_rows(min_row=2)):

    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participante = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horario = linha[5].value
    data_emissao = linha[6].value
    

    #transferir os arquivos para a imagem
    fonte_nome = ImageFont.truetype('./Awesome Snowflakes.ttf',120)
    fonte_geral = ImageFont.truetype('./Rows of Sunflowers.ttf',80)
    fonte_alternativa = ImageFont.truetype('./Awesome Snowflakes.ttf',60)

    imagem = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem)
    
                        #nome do participante
    desenhar.text((1000,815), nome_participante, fill='black', font = fonte_nome)

                        #curso
    desenhar.text((1060,935), nome_curso, fill='black', font = fonte_nome)

                        #participacao
    desenhar.text((1430,1055), tipo_participante, fill='black', font = fonte_nome)
                        
                        #data inicial 
    desenhar.text((750,1800), data_inicio, fill='black', font = fonte_alternativa)

                        #data final 
    desenhar.text((750,1950), data_final, fill='black', font = fonte_alternativa)

                        #data emissao 
    desenhar.text((2250,1950), data_emissao, fill='black', font = fonte_alternativa)

                        #carga horaria 
    desenhar.text((1480,1182), str(carga_horario), fill='black', font = fonte_nome)

    imagem.save(f'./{indice} {nome_participante}certificados.png')


    