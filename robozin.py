import pyautogui
import keyboard
import time

def alt_tab():
    keyboard.press('alt')
    keyboard.press('tab')
    time.sleep(0.2)
    keyboard.release('tab')
    keyboard.release('alt')
    time.sleep(0.5)

def copiar():
    keyboard.press_and_release('ctrl+c')
    time.sleep(0.3)

def colar():
    keyboard.press_and_release('ctrl+v')
    time.sleep(0.3)

def apertar_tab(n=1):
    for _ in range(n):
        pyautogui.press('tab')
        time.sleep(0.2)

def apertar_enter():
    pyautogui.press('enter')
    time.sleep(0.2)

# Pede ao usuário quantos números serão processados
total_numeros = int(input("Quantos números de celular você quer processar? "))

print("Preparado! Você tem 5 segundos pra colocar o Excel em foco...")
time.sleep(5)

# Primeira rodada (antes do loop)
copiar()
alt_tab()
apertar_tab(2)
colar()
time.sleep(1.5)
pyautogui.press('down')
apertar_tab(1)
alt_tab()
pyautogui.press('right')
copiar()
alt_tab()

colar()
apertar_enter()
alt_tab()

# Loop com contador
for i in range(1, total_numeros):
    print(f"Processando número {i+1} de {total_numeros}")
    pyautogui.press('down')
    pyautogui.press('left')
    copiar()
    alt_tab()
    apertar_tab(2)
    colar()
    time.sleep(1.5)
    pyautogui.press('down')
    apertar_tab(1)
    alt_tab()
    pyautogui.press('right')
    copiar()
    alt_tab()
    colar()
    apertar_enter()
    alt_tab()

print("✅ Fim do processo! Todos os números foram processados.")
