import pyautogui
import time
import ctypes
import subprocess
import os
import sys

x_pos = 0
y_pos = 0

#Iniciando programa com planilha dashboard em excel
caminho_dashboard = 'C:\\Users\\Gabriel Siza\\OneDrive\\Área de Trabalho\\Documents\\GABRIELIMPORTANTE\\Programing\\SQL\\Desafio 1'
nome_dashboard = 'Dashboard_de_vendas.xlsx'

caminho_completo = os.path.join(caminho_dashboard, nome_dashboard)

try:
    #Abrindo a planilha dashboard excel
    os.startfile(caminho_completo)
    time.sleep(5)
    
    #pressionando enter na caixa de alerta
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.keyUp('enter')
    
    #alterando para a janela com as queries
    pyautogui.hotkey('ctrl', 'pagedown')
    
    #Movendo o mouse para a posição da 1 querie
    pyautogui.moveTo(x_pos + 100, y_pos + 300, duration=1)
    time.sleep(1)
    
    # Execute o clique na querie
    pyautogui.click()

    #ctrl + a para selecionar toda a querie
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)

    #ctrl + c para copiar a querie
    pyautogui.hotkey('ctrl', 'c')

except FileNotFoundError:
    print(f'O arquivo "{nome_dashboard}" não foi encontrado no caminho especificado.')
except Exception as e:
    print(f'Erro em {Exception}')


#Comandos para abrir o PostGree SQL
caminho_postgree = "C:\\Program Files\\PostgreSQL\\16\\pgAdmin 4\\runtime\\pgAdmin4.exe"

#Abrindo o aplicativo PostgreeSQL
try:
    subprocess.Popen(caminho_postgree)
    print("postgree aberto com sucesso!")
except Exception as e:
    print(f"Ocorreu um erro ao abrir o postgreeSQL: {e}")
    print("O programa foi encerrado")
    sys.exit()
    

# Clique para abrir o SERVERS
try:
      # Mova o mouse para a posição do ícone
    time.sleep(15)
    pyautogui.moveTo(x_pos + 15, y_pos + 100, duration=1)

    # Espere um pouco antes de clicar
    time.sleep(1)

    # Execute o clique do mouse
    pyautogui.click()
    print("! click no SERVERS")
except Exception as e:
    print(f"Ocorreu um erro ao clicar no postgree: {e}")

# Clique para abrir o DATABASES
try:
      # Mova o mouse para a posição do ícone
    pyautogui.moveTo(x_pos + 45, y_pos + 140, duration=1)

    # Espere um pouco antes de clicar
    time.sleep(1)

    # Execute o clique do mouse
    pyautogui.click()
    print("! click no databases")
except Exception as e:
    print(f"Ocorreu um erro ao clicar no postgree: {e}")    

# Clique em Casts para liberar a função Open query
try:
      # Mova o mouse para a posição do ícone
    pyautogui.moveTo(x_pos + 145, y_pos + 190, duration=1)

    # Espere um pouco antes de clicar
    time.sleep(1)

    # Execute o clique do mouse
    pyautogui.click()
    print("! click no databases")
except Exception as e:
    print(f"Ocorreu um erro ao clicar no postgree: {e}")

#Pressionando Alt + Shift + Q para abrir o querry tool
def pressione_alt_shift_Q():
    # Pressionar as teclas Alt + Shift + Q
    ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)  # Tecla Alt
    ctypes.windll.user32.keybd_event(0x10, 0, 0, 0)  # Tecla Shift
    ctypes.windll.user32.keybd_event(0x51, 0, 0, 0)  # Tecla Q
def libere_alt_shift_Q():
    # Liberar as teclas Alt + Shift + Q
    ctypes.windll.user32.keybd_event(0x51, 0, 2, 0)  # Liberar a tecla Q
    ctypes.windll.user32.keybd_event(0x10, 0, 2, 0)  # Liberar a tecla Shift
    ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)  # Liberar a tecla Alt

pressione_alt_shift_Q()
time.sleep(1)
libere_alt_shift_Q()

#Pressione Ctrl + V para colar a querie e F5 para rodar
try:
    #hotkey em ctrl + v
    time.sleep(3)
    pyautogui.hotkey('ctrl', 'v')
    print("ctrl + v aplicado!")
    
    #pressione F5
    pyautogui.press('f5')
    print("F5 aplicado!")
except Exception as e:
    print(f"Ocorreu um erro ao pressionar ctrl + v: {e}")

#Copia e cola resultado da querie na planilha excel
try:
    #Movendo o mouse para a posição da 1 querie
    pyautogui.moveTo(x_pos + 350, y_pos + 785, duration=1)
    time.sleep(1)
    
    # Execute o clique na querie
    pyautogui.click()
    print("click no resultado da querie!")

    #ctrl + c para copiar a querie
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    print("ctrl + c aplicado!")

    #alt tab para trocar para o excel
    time.sleep(1)
    pyautogui.hotkey('alt', 'tab')
    print("alt tab aplicado!")

    #alterando a planilha para "Resultados"
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'pageup')
    print("ctrl + pgup aplicado com sucesso!")
    
    #Movendo o mouse para a posição da celula de colar (B4)
    pyautogui.moveTo(x_pos + 130, y_pos + 280, duration=1)
    time.sleep(1)
    
    # Execute o clique na querie
    pyautogui.click()
    print("click na celula B4!")

    #ctrl + v para colar o resultado da querie
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    print("ctrl + v aplicado!")

except Exception as e:
    print(f'Ocorreu um erro ao abrir o arquivo: {e}')

#agora teremos as iterações para as outras 4 queries do código
#devido algumas etapas que não precisarão ser realizadas como na primeira, agrupamos as 4 queries finais
#teremos 2 alterações em posição de mouse
#1 será na coleta da querie em excel
#2 será na hora de colar o resultado da querie no excel


import pyautogui
import time
import win32gui

#Função para encontrar a janela pelo título
def find_window(title):
    hwnd = win32gui.FindWindow(None, title)
    return hwnd

#Função para trazer a janela para frente
def bring_to_front(hwnd, bring_immediately=False):
    if bring_immediately:
        win32gui.ShowWindow(hwnd, 5)  #Restaura a janela se estiver minimizada
        win32gui.SetForegroundWindow(hwnd)
    else:
        print("Janela não foi trazida para frente imediatamente.")

#Função para executar o loop interno
def execute_inner_loop(x_pos, y_pos, j):
    if j == 0:
        pyautogui.moveTo(x_pos + 700, y_pos + 280, duration=1) # querie2
        print(f"Iteração {j+1}")
        time.sleep(1)
    elif j == 1:
        pyautogui.moveTo(x_pos + 1050, y_pos + 280, duration=1) # querie3
        print(f"Iteração {j+1}")
        time.sleep(1)
    elif j == 2:
        pyautogui.moveTo(x_pos + 1300, y_pos + 280, duration=1) # querie4
        print(f"Iteração {j+1}")
        time.sleep(1)
    elif j == 3:
        pyautogui.moveTo(x_pos + 1600, y_pos + 280, duration=1) # querie5
        print(f"Iteração {j+1}")

    # Aqui não faz mais parte do loop FOR
    # Hotkey ctrl + v para colar o resultado
    time.sleep(1)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')
    print("ctrl + v aplicado com sucesso!resultado de querie aplicado")
    pyautogui.hotkey('ctrl', 'pagedown')


# Encontrar a alça da janela do pgAdmin 4
pgadmin_hwnd = find_window("pgAdmin 4")

# Encontrar a alça da janela do Excel
excel_hwnd = find_window("Dashboard_de_vendas.xlsx - Excel (Falha na Ativação do Produto)")

# Verificar se a janela do pgAdmin 4 foi encontrada
if pgadmin_hwnd:
    # Não traz a janela do pgAdmin 4 para frente imediatamente
    bring_to_front(pgadmin_hwnd, bring_immediately=False)
else:
    print("A janela do pgAdmin 4 não foi encontrada.")

# Verificar se a janela do Excel foi encontrada
if excel_hwnd:
    # Não traz a janela do Excel para frente imediatamente
    bring_to_front(excel_hwnd, bring_immediately=False)
else:
    print("A janela do Excel não foi encontrada.")

x_pos = 0
y_pos = 0

try:  
    # Alterando para a janela com as queries
    pyautogui.hotkey('ctrl', 'pagedown') 
    
    # Movendo o mouse para a posição da 2 querie
    # Loop em for para alteração da posição do mouse para as próximas queries
    for i in range(4):
        if i == 0:
            pyautogui.moveTo(x_pos + 600, y_pos + 300, duration=1) # querie2
            print(f"Iteração {i+1}")
            time.sleep(1)
        elif i == 1:
            pyautogui.moveTo(x_pos + 1000, y_pos + 300, duration=1) # querie3
            print(f"Iteração {i+1}")
            time.sleep(1)
        elif i == 2:
            pyautogui.moveTo(x_pos + 1200, y_pos + 300, duration=1) # querie4
            print(f"Iteração {i+1}")
            time.sleep(1)
        elif i == 3:
            pyautogui.moveTo(x_pos + 1500, y_pos + 300, duration=1) # querie5
            print(f"Iteração {i+1}, esta foi a última iteração!")
            time.sleep(1)
        
        # Aqui não faz mais parte do loop FOR
        # Execute o clique na querie
        pyautogui.click()
        time.sleep(1)

        # Somente para teste
        pyautogui.click()
        time.sleep(1)

        # Ctrl + a para selecionar toda a querie
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(1)
        print("ctrl + a acionado!")

        # Ctrl + c para copiar a querie
        pyautogui.hotkey('ctrl', 'c')
        print("ctrl + c acionado!")
        
        # Alterar janela para o postgreSQL
        bring_to_front(pgadmin_hwnd, bring_immediately=True)

        # Somente para teste
        pyautogui.click()
        time.sleep(1)
        
        # Ctrl + a para selecionar tudo escrito no SQL
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'a')
        print("ctrl + a aplicado!")

        # Ctrl + v e F5 para executar a querie no postgreSQL
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'v')
        print("ctrl + v aplicado!")
        
        # Pressione F5
        pyautogui.press('f5')
        print("F5 aplicado!")

        # Movendo o mouse para a posição da 1 querie
        pyautogui.moveTo(x_pos + 350, y_pos + 785, duration=1)
        time.sleep(1)
        
        # Execute o clique na querie
        pyautogui.click()
        print("click no resultado da querie!")

        # Ctrl + c para copiar a querie
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'c')
        print("ctrl + c aplicado!")

        # Alt tab para trocar para o excel
        bring_to_front(excel_hwnd, bring_immediately=True)

        # Alterando a planilha para "Resultados"
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'pageup')
        print("ctrl + pgup aplicado com sucesso!")

        # Executando o loop interno
        execute_inner_loop(x_pos, y_pos, i)

except Exception as e:
    print(f'Erro em {e}')
