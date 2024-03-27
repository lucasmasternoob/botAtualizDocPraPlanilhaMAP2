import pyautogui as pyautogui
import time

#abre uma mensagem de alerta avisando que o software vai rodar
pyautogui.alert('O CÓDIGO VAI COMEÇAR A RODAR! TIRE AS MÃOS DO TECLADO E DO MOUSE, POIS AMBOS SERÃO CONTROLADOS POR UM BOT AGORA!')
time.sleep(1)

#           ATENÇÃO!       
#           - NECESSÁRIO COLOCAR O BLOCO DE NOTAS NA SEXTA POSIÇÃO NA BARRA DE TAREFAS, E TAMBÉM COLOCAR O EXCEL NA SÉTIMA POSIÇÃO!
#           - AO ABRIR A CAIXA DE PESQUISA NO EXCEL, SELECIONAR OPÇÃO "Diferenciar Maiúsculas de Minúsculas".
#           - ATERIORMENTE ERA NECESSÁRIO BLOQUEAR AS CÉLULAS DA PLANILHAS QUE REPRESENTAM A ESTRUTURA DO ESTOQUE PARA REALIZAR AS
#             BUSCAS PRECISAS DOS ÍTENS (PAREDES, CORREDORES E PORTAS). PORÉM ESSAS PARTES FORAM SUBSTITÚÍDAS POR IMAGENS, 
#             DESSA FORMA A PESQUISA SERÁ FEITA APENAS NAS CÉLULAS QUE TEM INFORMAÇÕES E DADOS E NÃO IMAGENS. pORÉM AINDA É 
#             NECESSÁRIO ADAPTAR O CÓDIGO PARA AUTOMATIZAAR O PROCESSO DE ORGANIZAÇÃO DAS INFORMAÇÕES CONFORME EXEMPLO NO MAP 3.
#           - NA PLANILHA MAP 3 ESTÁ PENDENTE SUBSTITUIR AS PARTES ESTRUTURAIS DA PNALINHAS POR IMAGENS PARA DIRECIONAR AS BUSCAS, 
#             EXATAMENTE COMO FOI FEITO NA PLANILHA MAP 2. 
#           - AGORA O CÓDIGO PODE SER OTIMIZADO E REFATORADO, RESUMINDO OS COMANDOS PARA SELECIONAR A CÉLULAS ONDE SERÁ FEITA A PESQUISA PARA APANAS UM 
#             LIQUE NA OPÇÃO PARA SELECIONAR TODAS AS CÉLULAS (1ª CÉLULA NO CANTO SUPERIOR ESQUERTO DA PLANILHA).
#           

pyautogui.hotkey('win', '6') #muda para a janela do bloco de notas (colocar bloco de notas na sexta posição na barra de tarefas)
time.sleep(1)
pyautogui.moveTo(17, 85, 1) #(x=17, y=85) - (x=100, y=88)
time.sleep(1)
pyautogui.click(button='left')
time.sleep(1)
pyautogui.press('Home') #colocar cursor na primeiro posição da linha 1 do doc
time.sleep(1)
pyautogui.keyDown('shiftleft') #pressiona a tecla shift esquerda para selecionar
time.sleep(1)
pyautogui.keyDown('shiftright') #pressiona a tecla shift direita para selecionar
time.sleep(1)
pyautogui.press('End') #pressiona o end para selecionar todo o conteúdo da linha 
time.sleep(1)
pyautogui.keyUp('shiftleft') #solta a tecla shift esquerda 
time.sleep(1)
pyautogui.keyUp('shiftright') #solta a tecla shift direita 
time.sleep(1)

pyautogui.hotkey('ctrl', 'c') #copia conteúdo selecionado
time.sleep(1)


pyautogui.hotkey('win', '7') #muda para janela da planilha (colocar planilha na sétima posição na barra de tarefas)
time.sleep(1)

#MOVE MOUSE PARA OPÇÃO DE SELEÇÃO DE TABELA INTEIRA
pyautogui.moveTo(17, 199, 1) #(x=17, y=199)
time.sleep(1)
pyautogui.click(button='left')
time.sleep(1)


#SELECIONA ÁREA ESPECÍFICA NA PLANILHA PARA SER REALIZA A PESQUISA
'''
pyautogui.moveTo(117, 298, 1) #move mouse para a primeira célula para selecionar a pesquisa na planilha (x=117, y=298)
time.sleep(1)
pyautogui.click(button='left')
time.sleep(1)
pyautogui.moveTo(1355, 279, 1) # move mouse para barra de rolagem (x=1355, y=279)
time.sleep(1)
pyautogui.mouseDown(1355, 279) #clica na barre de rolagem e segura
time.sleep(1)    
pyautogui.moveTo(1355, 611, 1) # desse a barra de rolagem até o final da planilha (x=1357, y=611)
time.sleep(1)    
pyautogui.mouseUp(1355, 611) #solta o clique do mouse 
time.sleep(1)  
pyautogui.keyDown('shiftleft') #pressiona a tecla shift esquerda para selecionar
time.sleep(1)
pyautogui.keyDown('shiftright') #pressiona a tecla shift direita para selecionar
time.sleep(1)
pyautogui.moveTo(899, 652, 1) # move o mouse clica na última célula para selecionar a área de pesquisa (x=899, y=652) (x=851, y=649)
time.sleep(1)
pyautogui.click(button='left')
time.sleep(1)
pyautogui.keyUp('shiftleft') #solta a tecla shift esquerda 
time.sleep(1)
pyautogui.keyUp('shiftright') #solta a tecla shift direita 
time.sleep(1)
'''



#ABRIR JANELA DE PESQUISA NA PLANILHA DE ACORDO COM A ÁREA SELECIONADA (pressiona atelhgo SHIFT+F5)
pyautogui.keyDown('shiftleft') #pressiona a tecla shift esquerda para selecionar
time.sleep(1)
pyautogui.keyDown('shiftright') #pressiona a tecla shift direita para selecionar
time.sleep(1)
pyautogui.press('f5')
time.sleep(1)
pyautogui.keyUp('shiftleft') #solta a tecla shift esquerda 
time.sleep(1)
pyautogui.keyUp('shiftright') #solta a tecla shift direita 
time.sleep(1)
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)
pyautogui.press('Home')
time.sleep(1)
pyautogui.press('right')
time.sleep(1)

#SELECIONAR APENAS INICIAL DA REFERÊNCIA PARA PESQUISAR NA PLANILHA
pyautogui.keyDown('shiftleft') #pressiona a tecla shift esquerda para selecionar
time.sleep(1)
pyautogui.keyDown('shiftright') #pressiona a tecla shift direita para selecionar
time.sleep(1)
pyautogui.press('end')
time.sleep(1)
pyautogui.keyUp('shiftleft') #solta a tecla shift esquerda 
time.sleep(1)
pyautogui.keyUp('shiftright') #solta a tecla shift direita 
time.sleep(1)
pyautogui.press('backspace')
time.sleep(1)
pyautogui.press('Enter')
time.sleep(1)
pyautogui.hotkey('alt', 'f4')
time.sleep(1)
pyautogui.press('f2')
time.sleep(1)
pyautogui.write('; ')
time.sleep(1)
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)
pyautogui.press('Enter')
time.sleep(1)


#ALTERA PARA JANELA DO DOCUMENTO PARA SELECIONAR A PRÓXIMO REFERÊNCIA DE PESQUISA
x=1 #SETA O VALOR DA VÁRIAVEL x DO CONTADOR EM 1
while(x<6): #lógica de laço de repetição
    pyautogui.hotkey('win', '6') #ALTERA PARA JANELA DO DOCUMENTO PARA SELECIONAR A PRÓXIMO REFERÊNCIA DE PESQUISA
    time.sleep(1)
    pyautogui.press('Down')
    time.sleep(1)

    #INÍCIO DO LOOP DE ATUALIZAÇÃO DA PLANILHA

    pyautogui.press('Home') #colocar cursor na primeiro posição da linha 1 do doc
    time.sleep(1)
    pyautogui.keyDown('shiftleft') #pressiona a tecla shift esquerda para selecionar
    time.sleep(1)
    pyautogui.keyDown('shiftright') #pressiona a tecla shift direita para selecionar
    time.sleep(1)
    pyautogui.press('End') #pressiona o end para selecionar todo o conteúdo da linha 
    time.sleep(1)
    pyautogui.keyUp('shiftleft') #solta a tecla shift esquerda 
    time.sleep(1)
    pyautogui.keyUp('shiftright') #solta a tecla shift direita 
    time.sleep(1)

    pyautogui.hotkey('ctrl', 'c') #copia conteúdo selecionado
    time.sleep(1)

    pyautogui.hotkey('win', '7') #muda para janela da planilha (colocar planilha na sétima posição na barra de tarefas)
    time.sleep(1)

    #MOVE MOUSE PARA OPÇÃO DE SELEÇÃO DE TABELA INTEIRA
    pyautogui.moveTo(17, 199, 1) #(x=17, y=199)
    time.sleep(1)
    pyautogui.click(button='left')
    time.sleep(1)

    #SELECIONA ÁREA ESPECÍFICA NA PLANILHA PARA SER REALIZA A PESQUISA
    '''
    pyautogui.moveTo(117, 298, 1) #move mouse para a primeira célula para selecionar a pesquisa na planilha (x=117, y=298)
    time.sleep(1)
    pyautogui.click(button='left')
    time.sleep(1)
    pyautogui.moveTo(1355, 279, 1) # move mouse para barra de rolagem (x=1355, y=279)
    time.sleep(1)
    pyautogui.mouseDown(1355, 279) #clica na barre de rolagem e segura
    time.sleep(1)    
    pyautogui.moveTo(1355, 611, 1) # desse a barra de rolagem até o final da planilha (x=1357, y=611)
    time.sleep(1)    
    pyautogui.mouseUp(1355, 611) #solta o clique do mouse 
    time.sleep(1)  
    pyautogui.keyDown('shiftleft') #pressiona a tecla shift esquerda para selecionar
    time.sleep(1)
    pyautogui.keyDown('shiftright') #pressiona a tecla shift direita para selecionar
    time.sleep(1)
    pyautogui.moveTo(899, 652, 1) # move o mouse clica na última célula para selecionar a área de pesquisa (x=899, y=652) (x=851, y=649)
    time.sleep(1)
    pyautogui.click(button='left')
    time.sleep(1)
    pyautogui.keyUp('shiftleft') #solta a tecla shift esquerda 
    time.sleep(1)
    pyautogui.keyUp('shiftright') #solta a tecla shift direita 
    time.sleep(1)
    '''


    #ABRIR JANELA DE PESQUISA NA PLANILHA DE ACORDO COM A ÁREA SELECIONADA (pressiona atelhgo SHIFT+F5)
    pyautogui.keyDown('shiftleft') #pressiona a tecla shift esquerda para selecionar
    time.sleep(1)
    pyautogui.keyDown('shiftright') #pressiona a tecla shift direita para selecionar
    time.sleep(1)
    pyautogui.press('f5')
    time.sleep(1)
    pyautogui.keyUp('shiftleft') #solta a tecla shift esquerda 
    time.sleep(1)
    pyautogui.keyUp('shiftright') #solta a tecla shift direita 
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('Home')
    time.sleep(1)
    pyautogui.press('right')
    time.sleep(1)

    #SELECIONAR APENAS INICIAL DA REFERÊNCIA PARA PESQUISAR NA PLANILHA
    pyautogui.keyDown('shiftleft') #pressiona a tecla shift esquerda para selecionar
    time.sleep(1)
    pyautogui.keyDown('shiftright') #pressiona a tecla shift direita para selecionar
    time.sleep(1)
    pyautogui.press('end')
    time.sleep(1)
    pyautogui.keyUp('shiftleft') #solta a tecla shift esquerda 
    time.sleep(1)
    pyautogui.keyUp('shiftright') #solta a tecla shift direita 
    time.sleep(1)
    pyautogui.press('backspace')
    time.sleep(1)
    pyautogui.press('Enter')
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(1)
    pyautogui.press('f2')
    time.sleep(1)
    pyautogui.write('; ')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('Enter')
    time.sleep(1)

    pyautogui.moveTo(1357, 197, 1) # move mouse para barra de rolagem (x=1357, y=197)
    time.sleep(1)
    pyautogui.click(button='left')
    pyautogui.click(button='left')
    pyautogui.click(button='left')
    
    time.sleep(1)


    #FIM DO LOOP DE ATUALIZAÇÃO DA PLANILHA

    x=x+1
    time.sleep(1)


pyautogui.alert('O CÓDIGO TERMONOU DE RODAR!')