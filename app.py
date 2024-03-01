import openpyxl
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # cada célula que contém a informação necessária
    nome_curso = nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value
    laboratorio = linha[7].value
    
    # Transferir os dados da planilha para a imagem do certificado
    font_nome = ImageFont.truetype('./tahomabd.ttf', 50)
    font_geral = ImageFont.truetype('./tahoma.ttf', 50)
    font_data = ImageFont.truetype('./tahoma.ttf', 50)
    
    image = Image.open('./certificado_padrao.png')
    desenhar = ImageDraw.Draw(image)
    
    desenhar.text((510,330), nome_participante, fill='black', font=font_nome)
    desenhar.text((660,390), str(laboratorio), fill='black', font=font_geral)
    desenhar.text((510,458), nome_curso, fill='black', font=font_geral)
    desenhar.text((900,520), tipo_participacao, fill='black', font=font_geral)
    desenhar.text((730,585), str(carga_horaria) + ' horas', fill='black', font=font_geral)
    desenhar.text((320,970), str(data_inicio), fill='blue', font=font_data)
    desenhar.text((320,1150), str(data_final), fill='blue', font=font_data)
    desenhar.text((1480,1040), str(data_emissao), fill='blue', font=font_data)
    
    image.save(f'./certificados/{indice}{nome_participante} certificado.png')
    
    