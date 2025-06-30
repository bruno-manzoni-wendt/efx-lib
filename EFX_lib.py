#imports
from datetime import datetime
from ctypes import windll
import pyautogui as pyg
from time import sleep
import pygetwindow
import webbrowser
import subprocess
import pyscreeze
import pyperclip
import sys
import os

#==================================================#
#Constants & Definitions
PAUSE_PADRAO = 0.3
pyg.PAUSE = PAUSE_PADRAO
pyscreeze.USE_IMAGE_NOT_FOUND_EXCEPTION = False
FIXED_PYG_IMG_PATH = r'G:\Meu Drive\Bruno\Python\pyautogui_img\padrao'
dynamic_pyg_img_path = FIXED_PYG_IMG_PATH
larguraMonitor, alturaMonitor = pyg.size()

def set_pyg_pause(time: float):
    global PAUSE_PADRAO
    PAUSE_PADRAO = time
    pyg.PAUSE = PAUSE_PADRAO

def set_pyg_confidence(conf: float = None):
    return 0.99 if conf is None else conf

def set_dynamic_pyg_path(path = None): #função para definir um caminho dinamico das imagens do pyautogui, porém com a opção de alterá-lo
    return dynamic_pyg_img_path if path is None else path

def set_fixed_pyg_path(path = None): #função para definir um caminho fixo das imagens do pyautogui, porém com a opção de alterá-lo
    return FIXED_PYG_IMG_PATH if path is None else path

#==================================================#
#TODO
# import keyboard escrever textos com o keyboard import
# keyboard.write("Isso é apenas um teste com ceçedilha", 0.01)

#TODO
#mudar funções para que quando for receber path e file, seja nessa ordem 1º path, 2º file

#==================================================#
#functions

def procurar(img: str, click: bool = False, move_mouse: bool = False, path = None,
             region: tuple = (0, 0, larguraMonitor, alturaMonitor), conf: float = None):
    conf = set_pyg_confidence(conf)
    path = set_dynamic_pyg_path(path)
    img_path = os.path.join(path, img)
    file_exists(img_path)
    action = 'Clicar' if click else 'Procurar'
    c = 0
    t = 0.50

    while True:
        # if 10 >= c > 2:
        #     conf = round(conf - 0.01*t, 4)
        print(f"\r{c:.2f} {action} {img}", end="", flush=True)
        position = pyg.locateCenterOnScreen(image= img_path, grayscale=True, confidence=conf, region=region) # type: ignore
        if position is not None:
            print('\n', position, '\n')
            break

        sleep(t)
        c = round(c + t, 2)

    if click:
        pyg.click(position)
        if move_mouse:
            sleep(0.1)
            pyg.moveTo(larguraMonitor/2, alturaMonitor*3/4)

    sleep(PAUSE_PADRAO)
    return position


def procurar_uma(img: str, img_path: str, conf: float = None, region: tuple = (0, 0, larguraMonitor, alturaMonitor), grayscale: bool = True):
    return pyg.locateCenterOnScreen(image= os.path.join(img_path, img), grayscale=grayscale, confidence= set_pyg_confidence(conf), region=region) # type: ignore


def check_pixel(x:int, y:int, rgb:tuple = None, click:bool = False, move_mouse:bool = False, tolerance:int = 0):
    c = 0
    t = 0.25
    defined_rgb = pyg.pixel(x, y) if rgb is None else rgb

    while True:
        print(f"\r{c:.2f} Cheking changes in the pixel {x}, {y}", end=" + ", flush=True)

        if not pyg.pixelMatchesColor(x, y, defined_rgb, tolerance=tolerance):
            new_rgb = pyg.pixel(x, y)
            print('\n', f'Pixel changed from {defined_rgb} to {new_rgb}', '\n')
            break

        sleep(t)
        c = round(c + t, 2)

    if click:
        pyg.click(x, y)
        if move_mouse:
            sleep(0.1)
            pyg.moveTo(larguraMonitor/2, alturaMonitor*3/4)

    sleep(PAUSE_PADRAO)
    return new_rgb


def stop_if_not_found(img: str, img_path = None):
    print('Esperando ' + img)
    path = set_dynamic_pyg_path(img_path)
    while procurar_uma(img, path) is not None:
        sleep(2)
    print(img+' sumiu.')


def window_to_main_monitor(window_name: str, img: str, path = None):
    print('Procurando janela:', window_name)
    while True:
        try:
            janela = pygetwindow.getWindowsWithTitle(window_name)[0]
            print('Encontrado\n')
            janela.maximize()
            break
        except:
            sleep(1)

    sleep(1)
    if procurar_uma(img, set_fixed_pyg_path(path), conf=0.9) is None: #TODO será que tem uma outra solução? pois posso encontrar a img de uma janela que não é a que quero
        janela.restore()
        sleep(0.7)
        janela.moveTo(0, 0)
        sleep(0.7)
        janela.maximize()
        sleep(0.7)


def open_link_chrome(url: str):
    webbrowser.get(r'"C:\Program Files\Google\Chrome\Application\chrome.exe" %s').open_new(url)
    if url.count('/') >= 3:
        short = "/".join(url.split("/")[:3])
        print(f'Opening {short}')
    else:
        print(f'Opening {url}')


def chrome_to_main_monitor(link: str):
    open_link_chrome(link)
    sleep(1)
    if procurar_uma('chrome.png', FIXED_PYG_IMG_PATH) is None:
        open_link_chrome('https://www.google.com.br/')
        window_to_main_monitor('Google - Google Chrome', 'chrome.PNG', FIXED_PYG_IMG_PATH)
        pyg.hotkey('ctrl', 'w')


def salvar_selecionar_arquivo(file_name: str, file_path: str, override: bool):
    procurar("salvar_como.PNG", path= FIXED_PYG_IMG_PATH)
    sleep(0.5)
    pyperclip.copy(file_name)
    pyg.hotkey("ctrl", "v")
    sleep(0.4)
    pyg.hotkey("ctrl", "l")
    sleep(0.4)
    pyperclip.copy(file_path)
    pyg.hotkey("ctrl", "v")
    sleep(0.3)
    pyg.press("enter", 1, 0.7)
    pyg.press("enter", 3, 0.21)

    if override:
        sleep(1.1)
        pyg.press("tab")
        pyg.press("enter")


def check_if_logged(img_logged: str, img_not_logged: str, img_path: str):
    img_path = set_dynamic_pyg_path(img_path)
    print('Verificando se está logado')
    while True:
        login = [procurar_uma(img_not_logged, img_path, conf = 0.9), procurar_uma(img_logged, img_path, conf = 0.9)]
        if any(login):
            break
        sleep(1)

    if login[0] is not None:
        procurar(img_not_logged, True, path=img_path, conf = 0.9)


def print_same_line(text):
    print(f"\r{text}", end="", flush=True)


#TODO GET FILE TIMESTAMP
def file_last_update(file_path: str, time: int = 21):
    """ time (int) = How many seconds would would you like to keep checking?
    """
    if not os.path.exists(file_path):
        sleep(1)

    last_modified = os.path.getmtime(file_path)
    file = os.path.basename(file_path)
    count = 0
    while count <= time:
        print(f"\r--- {count}/{time} WAITING FOR {file} ---", end="", flush=True)
        if os.path.exists(file_path):
            if os.path.getmtime(file_path) != last_modified:
                print(f'\n{file} UPDATED\n')
                sleep(1)
                break
        count += 1
        sleep(1)


def looking_file(file: str, path: str, time: int = 21):
    """WAITS FOR A FILE TO APPEAR IN A DEFINED PATH (WHEN DONWLOADING FOR EXAMPLE)""" #print(f"\r{round(c, 2)} {action} {img}", end="", flush=True)
    file_path = os.path.join(path, file)
    count = 0
    while count <= time:
        print(f'\r--- {count}/{time} LOOKING FOR {file} ---', end="", flush=True)
        if os.path.exists(file_path):
            print(f'\n{file} FOUND\n')
            break
        count += 1
        sleep(1)

# def split_path_file(file_path: str):
#     return os.path.split(file_path)

# def join_file_path(path: str, file: str):
#     return os.path.join(path, file)

def file_exists(file_path: str):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f'EFX ERROR, file not found in {file_path}.')


def win_r(file_path: str):
    file_exists(file_path)
    print(f'Win + R: {os.path.basename(file_path)}')
    pyg.hotkey('win', 'r')
    pyperclip.copy(file_path)
    pyg.press('delete')
    sleep(0.2)
    pyg.hotkey('ctrl', 'v')
    sleep(0.2)
    pyg.press('enter')


def press_keys(keys: list, time: float = 0.1):
    for k in keys:
        pyg.press(k)
        sleep(time)


def open_process(program_path: str):
    subprocess.Popen([program_path])
    print('ABRINDO', os.path.basename(program_path))


def copiar_colar(texto: str):
    pyperclip.copy(texto)
    pyg.hotkey('ctrl', 'v')


def capslock(on_or_off: bool):
    if on_or_off == True and windll.user32.GetKeyState(0x14) == 0:
        pyg.press('capslock')
        print('Capslock on')
    if on_or_off == False and windll.user32.GetKeyState(0x14) == 1:
        pyg.press('capslock')
        print('Capslock off')


def excel_update_all():
    print('\n- UPDATING ALL -')
    press_keys(['alt', 's', 'g', 'a'], 0.321)
    sleep(5)
    # procurar('excel_update_off.png', FIXED_PYG_IMG_PATH)
    while procurar_uma('excel_update_off.png', FIXED_PYG_IMG_PATH, 0.95) is None:
        sleep(2)
    print('- UPDATE FINISHED -\n')
    sleep(2)


def question_user(question: str, exit_script: bool = True):
    while True:
        resposta = input(question.strip() + ' (Y/N) ').upper()
        if resposta in ['Y', 'S', 'YES', 'SIM', 'OK', 'K']:
            print('--- CONFIRMED ---\n')
            return True

        elif resposta in ['N', 'NO', 'NÃO', 'NAO']:
            print('--- CANCELED ---\n')
            sleep(1)
            if exit_script:
                sys.exit()
            else:
                return False


def last_column_to_first(df):
    colunas = df.columns.tolist()  # Obtém a lista das colunas
    colunas = [colunas[-1]] + colunas[:-1]  # Reorganiza colocando a última coluna no início
    return df[colunas]  # Reordena o DataFrame


def beatiful_string(s: str):
    string = f'---> {s} <---'
    print('')
    print('='*len(string))
    print(string)
    print('='*len(string))
    print('')


def formatar_data(date, with_year: bool = True):
    return date.strftime("%d/%m/%Y") if with_year else date.strftime("%d/%m")


def hoje():
    return datetime.today()


def hoje_formatado(with_year: bool = True):
    return formatar_data(hoje(), with_year)


def download_drive_excel(url: str, file_path: str): #nome_janela: str,
    chrome_to_main_monitor(url)
    salvar_selecionar_arquivo(os.path.basename(file_path), os.path.dirname(file_path), True)
    file_last_update(file_path)


if __name__ == "__main__":
    print(f'RUNNING SMOOTHLY ON {__file__}')
else:
    print(f'{__name__}.py IMPORTED!', '\n')

#==================================================#

#TODO DEPRECATED