import pandas as pd
import pywhatkit
import datetime
import calendar

# Lendo os dados da planilha Excel
df = pd.read_excel("AlertaAutomacao.xlsx", sheet_name="Alerta")

# Transforma todas as células em strings
df = df.astype(str)

# Acessar células específicas da planilha
# '2' é a linha onde está os dias dos DOMINGOS.

dataDomingoPrimeiro = df.loc[2, "coluna 2"]
dataDomingoSegundo = df.loc[2, "coluna 3"]
dataDomingoTerceiro = df.loc[2, "coluna 4"]
dataDomingoQuarto = df.loc[2, "coluna 5"]
dataDomingoQuinto = df.loc[2, "coluna 6"]

# Acessar células específicas da planilha
# '3' é a linha onde está os nomes dos líderes que estarão nos DOMINGOS.

nomeDomingoPrimeiro = df.loc[3, "coluna 2"]
nomeDomingoSegundo = df.loc[3, "coluna 3"]
nomeDomingoTerceiro = df.loc[3, "coluna 4"]
nomeDomingoQuarto = df.loc[3, "coluna 5"]
nomeDomingoQuinto = df.loc[3, "coluna 6"]

# Acessar células específicas da planilha
# '26' é a linha onde está os dias das QUINTAS.

dataQuintaPrimeira = df.loc[24, "coluna 2"]
dataQuintaSegunda = df.loc[24, "coluna 3"]
dataQuintaTerceira = df.loc[24, "coluna 4"]
dataQuintaQuarta = df.loc[24, "coluna 5"]
dataQuintaQuinta = df.loc[24, "coluna 6"]

# Acessar células específicas da planilha
# '27' é a linha onde está os nomes dos líderes que estarão nas QUINTAS.

nomeQuintaPrimeira = df.loc[25, "coluna 2"]
nomeQuintaSegunda = df.loc[25, "coluna 3"]
nomeQuintaTerceira = df.loc[25, "coluna 4"]
nomeQuintaQuarta = df.loc[25, "coluna 5"]
nomeQuintaQuinta = df.loc[25, "coluna 6"]

# Acessar células específicas da planilha
# 'I' é a coluna onde está os numeros dos líderes.

numeroAndrezinho = df.loc[4, "coluna 9"]
numeroElias = df.loc[5, "coluna 9"]
numeroLidia = df.loc[6, "coluna 9"]
numeroFranklin = df.loc[7, "coluna 9"]
numeroJean = df.loc[8, "coluna 9"]
numeroMilton = df.loc[9, "coluna 9"]
numeroMotta = df.loc[10, "coluna 9"]

# Acessar células específicas da planilha
# 'H' é a coluna onde está os numeros dos líderes.

nomeAndrezinho = df.loc[4, "coluna 8"]
nomeElias = df.loc[5, "coluna 8"]
nomeFranklin = df.loc[6, "coluna 8"]
nomeLidia = df.loc[7, "coluna 8"]
nomeJean = df.loc[8, "coluna 8"]
nomeMilton = df.loc[9, "coluna 8"]
nomeMotta = df.loc[10, "coluna 8"]


# Variável para gerar link entre os nomes e os números
# Criar objeto da classe Listas
nomes = [
    nomeAndrezinho,
    nomeElias,
    nomeFranklin,
    nomeLidia,
    nomeJean,
    nomeMilton,
    nomeMotta,
]
numeros = [
    numeroAndrezinho,
    numeroElias,
    numeroLidia,
    numeroFranklin,
    numeroJean,
    numeroMilton,
    numeroMotta,
]
# Repertórios dos Domingos
repertorioDomingoUm = [
    df.loc[36, "coluna 2"],
    df.loc[37, "coluna 2"],
    df.loc[38, "coluna 2"],
    df.loc[39, "coluna 2"],
]

repertorioDomingoDois = [
    df.loc[36, "coluna 3"],
    df.loc[37, "coluna 3"],
    df.loc[38, "coluna 3"],
    df.loc[39, "coluna 3"],
]

repertorioDomingoTres = [
    df.loc[36, "coluna 4"],
    df.loc[37, "coluna 4"],
    df.loc[38, "coluna 4"],
    df.loc[39, "coluna 4"],
]

repertorioDomingoQuatro = [
    df.loc[36, "coluna 5"],
    df.loc[37, "coluna 5"],
    df.loc[38, "coluna 5"],
    df.loc[39, "coluna 5"],
]

repertorioDomingoCinco = [
    df.loc[36, "coluna 6"],
    df.loc[37, "coluna 6"],
    df.loc[38, "coluna 6"],
    df.loc[39, "coluna 6"],
]

# Repertórios das Quintas-Feiras
repertorioQuintaUm = [
    df.loc[14, "coluna 2"],
    df.loc[15, "coluna 2"],
    df.loc[16, "coluna 2"],
    df.loc[17, "coluna 2"],
]

repertorioQuintaDois = [
    df.loc[14, "coluna 3"],
    df.loc[15, "coluna 3"],
    df.loc[16, "coluna 3"],
    df.loc[17, "coluna 3"],
]

repertorioQuintaTres = [
    df.loc[14, "coluna 4"],
    df.loc[15, "coluna 4"],
    df.loc[16, "coluna 4"],
    df.loc[17, "coluna 4"],
]

repertorioQuintaQuatro = [
    df.loc[14, "coluna 5"],
    df.loc[15, "coluna 5"],
    df.loc[16, "coluna 5"],
    df.loc[17, "coluna 5"],
]

repertorioQuintaCinco = [
    df.loc[14, "coluna 6"],
    df.loc[15, "coluna 6"],
    df.loc[16, "coluna 6"],
    df.loc[17, "coluna 6"],
]

# Obtém a data atual
hoje = datetime.datetime.now()

# Verifica se hoje é uma segunda-feira ou sexta-feira
if hoje.weekday() == 0 or hoje.weekday() == 4:
    # Obtém o número de dias no mês atual
    numero_de_dias = calendar.monthrange(hoje.year, hoje.month)[1]

    # Define as mensagens para cada segunda-feira e sexta-feira do mês
    # Join é um método usado para concatenar uma lista em Strings para a mensagem poder ser enviada.
    mensagens_segunda = [
        "Fala aí "
        + nomeQuintaPrimeira
        + ", passando para lembrar que *quinta-feira* você é o (a) líder do dia.\nNão se esqueça:\n• Dos tons das músicas\n• Quem irá solar as músicas\n"
        + "\n".join(repertorioQuintaUm),
        "Fala aí "
        + nomeQuintaSegunda
        + ", passando para lembrar que *quinta-feira* você é o (a) líder do dia.\nNão se esqueça:\n• Dos tons das músicas\n• Quem irá solar as músicas\n"
        + "\n".join(repertorioQuintaDois),
        "Fala aí "
        + nomeQuintaTerceira
        + ", passando para lembrar que *quinta-feira* você é o (a) líder do dia.\nNão se esqueça:\n• Dos tons das músicas\n• Quem irá solar as músicas\n"
        + "\n".join(repertorioQuintaTres),
        "Fala aí "
        + nomeQuintaQuarta
        + ", passando para lembrar que *quinta-feira* você é o (a) líder do dia.\nNão se esqueça:\n• Dos tons das músicas\n• Quem irá solar as músicas\n"
        + "\n".join(repertorioQuintaQuatro),
        "Fala aí "
        + nomeQuintaQuinta
        + ", passando para lembrar que *quinta-feira* você é o (a) líder do dia.\nNão se esqueça:\n• Dos tons das músicas\n• Quem irá solar as músicas\n"
        + "\n".join(repertorioQuintaCinco),
    ]

    # Join é um método usado para concatenar uma lista em Strings para a mensagem poder ser enviada.
    mensagens_sexta = [
        "Fala aí "
        + nomeDomingoPrimeiro
        + ", passando para lembrar que *domingo* você é o (a) líder do dia.\nNão se esqueça:\n• Dos tons das músicas\n• Quem irá solar as músicas\n"
        + "\n".join(repertorioDomingoCinco),
        "Fala aí "
        + nomeDomingoSegundo
        + ", passando para lembrar que *domingo* você é o (a) líder do dia.\nNão se esqueça:\n• Dos tons das músicas\n• Quem irá solar as músicas\n"
        + "\n".join(repertorioDomingoDois),
        "Fala aí "
        + nomeDomingoTerceiro
        + ", passando para lembrar que *domingo* você é o (a) líder do dia.\nNão se esqueça:\n• Dos tons das músicas\n• Quem irá solar as músicas\n"
        + "\n".join(repertorioDomingoTres),
        "Fala aí "
        + nomeDomingoQuarto
        + ", passando para lembrar que *domingo* você é o (a) líder do dia.\nNão se esqueça:\n• Dos tons das músicas\n• Quem irá solar as músicas\n"
        + "\n".join(repertorioDomingoQuatro),
        "Fala aí "
        + nomeDomingoQuinto
        + ", passando para lembrar que *domingo* você é o (a) líder do dia.\nNão se esqueça:\n• Dos tons das músicas\n• Quem irá solar as músicas\n"
        + "\n".join(repertorioDomingoCinco),
    ]

    # Obtém o número da semana atual no mês (1-5)
    semana_atual = (hoje.day - 1) // 7 + 1
    print(semana_atual)

    # Verifica se hoje é uma segunda-feira ou sexta-feira da primeira semana do mês
    if hoje.weekday() == 0 and semana_atual == 1:
        mensagem = mensagens_segunda[0]
        nomeDoDia = nomeQuintaPrimeira
        indiceDoLider = nomes.index(nomeDoDia)
        whatsapp = numeros[indiceDoLider]

    elif hoje.weekday() == 4 and semana_atual == 1:
        mensagem = mensagens_sexta[0]
        nomeDoDia = nomeDomingoPrimeiro
        indiceDoLider = nomes.index(nomeDoDia)
        whatsapp = numeros[indiceDoLider]

    # Verifica se hoje é uma segunda-feira ou sexta-feira da segunda semana do mês
    elif hoje.weekday() == 0 and semana_atual == 2:
        mensagem = mensagens_segunda[1]
        nomeDoDia = nomeQuintaSegunda
        # Faz com que item de uma lista1 enterja com o item equivalente da lista2 colocando ele em uma Variável.
        indiceDoLider = nomes.index(nomeDoDia)
        whatsapp = numeros[indiceDoLider]

    elif hoje.weekday() == 4 and semana_atual == 2:
        mensagem = mensagens_sexta[1]
        nomeDoDia = nomeDomingoSegundo
        indiceDoLider = nomes.index(nomeDoDia)
        whatsapp = numeros[indiceDoLider]

    # Verifica se hoje é uma segunda-feira ou sexta-feira da terceira semana do mês
    elif hoje.weekday() == 0 and semana_atual == 3:
        mensagem = mensagens_segunda[2]
        nomeDoDia = nomeQuintaTerceira
        indiceDoLider = nomes.index(nomeDoDia)
        whatsapp = numeros[indiceDoLider]

    elif hoje.weekday() == 4 and semana_atual == 3:
        mensagem = mensagens_sexta[2]
        nomeDoDia = nomeDomingoTerceiro
        indiceDoLider = nomes.index(nomeDoDia)
        whatsapp = numeros[indiceDoLider]

    # Verifica se hoje é uma segunda-feira ou sexta-feira da quarta semana do mês
    elif hoje.weekday() == 0 and semana_atual == 4:
        mensagem = mensagens_segunda[3]
        nomeDoDia = nomeQuintaQuarta
        indiceDoLider = nomes.index(nomeDoDia)
        whatsapp = numeros[indiceDoLider]

    elif hoje.weekday() == 4 and semana_atual == 4:
        mensagem = mensagens_sexta[3]
        nomeDoDia = nomeDomingoQuarto
        indiceDoLider = nomes.index(nomeDoDia)
        whatsapp = numeros[indiceDoLider]

    # Verifica se hoje é uma segunda-feira ou sexta-feira da quinta semana do mês
    elif hoje.weekday() == 0 and semana_atual == 5:
        mensagem = mensagens_segunda[4]
        nomeDoDia = nomeQuintaQuinta
        indiceDoLider = nomes.index(nomeDoDia)
        whatsapp = numeros[indiceDoLider]

    elif hoje.weekday() == 4 and semana_atual == 5:
        mensagem = mensagens_sexta[4]
        nomeDoDia = nomeDomingoQuinto
        indiceDoLider = nomes.index(nomeDoDia)
        whatsapp = numeros[indiceDoLider]

    # Caso contrário, não envia mensagem
    else:
        mensagem = None

    # Se houver uma mensagem para enviar, encontra o número de telefone do destinatário e envia a mensagem
    if mensagem is not None:
        pywhatkit.sendwhatmsg(whatsapp, mensagem, 09, 30)
        print(f"Mensagem enviada para {whatsapp}")
