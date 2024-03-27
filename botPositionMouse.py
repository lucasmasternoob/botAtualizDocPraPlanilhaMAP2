import pyautogui as pyautogui
import time

pyautogui.alert('O CÓDIGO VAI COMEÇAR A RODAR! POSICIONE O MOUSE NA POSIÇÃO DESEJADA E AGUARTDE 10 SEGUNDOS E A APOSIÇÃO NAS COORDENAS X e Y/ APARECERÃO NO CONSOLE!')
time.sleep(10)
print(pyautogui.position())
pyautogui.alert('O CÓDIGO TERMONOU DE RODAR! Confira as coordenadas X e Y no console!')