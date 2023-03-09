Projeto: ***Alerta de liderança de louvor para igreja***

Este projeto consiste em um código em Python que faz o envio de mensagens via WhatsApp para os líderes de louvor 
informando a data em que eles irão liderar o louvor da igreja, 
os seus deveres e quais serão os louvores tocados. 
O código utiliza uma planilha em Excel para armazenar essas informações.

--------------- Pré-requisitos ---------------

Para utilizar esse código, você precisará ter instalado em sua máquina:

Python 3.x
pandas
pywhatkit

--------------- Como utilizar ---------------

Clone este repositório em sua máquina.
Certifique-se de ter instalado os pré-requisitos listados acima.
Abra o arquivo "AlertaAutomacao.xlsx" e preencha as informações referentes aos líderes de louvor e repertório das músicas.
Salve o arquivo.
Execute o arquivo "alerta_lideranca_louvor.py".
O programa irá acessar a planilha e obter as informações necessárias para enviar as mensagens via WhatsApp.
Aguarde a execução do programa.

--------------- Observações ---------------

Certifique-se de que o arquivo "AlertaAutomacao.xlsx" esteja no mesmo diretório do arquivo "alerta_lideranca_louvor.py".
Caso os líderes de louvor mudem ou as músicas do repertório sejam alteradas, basta atualizar a planilha e executar o programa novamente. 
As mensagens serão enviadas com as informações atualizadas.
É necessário que o número de telefone dos líderes de louvor esteja cadastrado em seu celular e que eles estejam utilizando o WhatsApp.

--------------- Limitações ---------------

O código foi desenvolvido levando em consideração a estrutura da planilha "AlertaAutomacao.xlsx" fornecida. 
Alterações na estrutura da planilha podem afetar o funcionamento do código.
O código não verifica se o número de telefone cadastrado é válido ou se o WhatsApp do destinatário está funcionando corretamente.
O código foi testado apenas em ambiente Windows. Pode haver incompatibilidades em outros sistemas operacionais.
