import openpyxl
from PIL import Image, ImageDraw, ImageFont


workbook = openpyxl.load_workbook('planilha_alunos.xlsx'); # Carregar o arquivo Excel
sheet_alunos = workbook['Sheet1']; # Selecionar a planilha desejada

# Varredura na planilha e acesso às células com informações
for indice, linha in enumerate (sheet_alunos.iter_rows(min_row=2)):  
    curso = linha[0].value 
    nomeParticipante = linha[1].value
    tipoParticipacao = linha[2].value
    dataInicio = linha[3].value
    dataTermino = linha[4].value
    cargaHoraria = linha[5].value        
    dtEmissao = linha[6].value
    
      
    # Definindo a fonte a ser usada e seu tamanho
    fontName = ImageFont.truetype('ARLRDBD.TTF',90)
    fontDate = ImageFont.truetype('ARLRDBD.TTF',55)

    # Carregar a imagem do certificado
    image = Image.open('certificadoPadrao.jpg')
    desenhar = ImageDraw.Draw(image)

# Desenhar as informações na imagem do certificado
desenhar.text((1030,825),nomeParticipante,fill='black',font=fontName)
desenhar.text((1080,944),curso,fill='black',font=fontName)
desenhar.text((1440,1055),tipoParticipacao,fill='black',font=fontName)
desenhar.text((1485,1180),str(cargaHoraria),fill='black',font=fontName)
desenhar.text((735,1790),dataInicio,fill='black',font=fontDate)
desenhar.text((730,1942),dataTermino,fill='black',font=fontDate)
desenhar.text((2205,1937),dtEmissao,fill='black',font=fontDate)    

# Salvar a imagem do certificado com um nome único
image.save(f'{indice} {nomeParticipante} certificado.png')

# Fechar o arquivo Excel após a conclusão do processamento
workbook.close()





    








