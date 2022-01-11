import pyautogui
import PySimpleGUI as sg
import win32gui, win32con
import threading
import ctypes
import time
import datetime
import os
import platform
import random

app_options = {
    'windows': [],
    'hwnds': [],
    'window': None,
    'active_index': -1,
    'pages': 0,
    'next_key': 'right',
    'ctrl_key': '',
    'alt_key': '',
    'shit_key': '',
    'top': 0,
    'left': 0,
    'bottom': 0,
    'right': 0,
    'downloads': r'~/Downloads'
}
pltfrm = platform.system()
if pltfrm == 'Windows':
    app_options['downloads'] = f'C:{os.sep}Users{os.sep}' + os.environ['USERNAME'] + f'{os.sep}Downloads'
user32 = ctypes.windll.user32
screensize = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)


def reset_options():
    app_options['windows'] = []
    app_options['hwnds'] = []
    app_options['active_index'] = -1
    app_options['pages'] = 0
    app_options['next_key'] = 'right'
    app_options['ctrl_key'] = ''
    app_options['alt_key'] = ''
    app_options['shit_key'] = ''
    reset_xy_options()


def reset_xy_options():
    app_options['top'] = 0
    app_options['left'] = 0
    app_options['bottom'] = 0
    app_options['right'] = 0


def init_capture(window, values):
    window = app_options['window']
    selitem = window.Element('_WINDOWS_').Widget.curselection()
    if 0 == len(selitem):
        reset_options();
        return

    idx = selitem[0]
    app_options['active_index'] = idx

    if app_options['top'] == app_options['bottom']:
        if app_options['top'] == 0 and app_options['bottom'] == 0:
            app_options['bottom'] = screensize[1];
        else:
            reset_options();
            return
    elif app_options['top'] > app_options['bottom']:
        tmpval = app_options['bottom']
        app_options['bottom'] = app_options['top']
        app_options['top'] = tmpval

    if app_options['left'] == app_options['right']:
        if app_options['left'] == 0 and app_options['right'] == 0:
            app_options['right'] = screensize[0];
        else:
            reset_options();
            return
    elif app_options['right'] < app_options['left']:
        tmpval = app_options['left']
        app_options['left'] = app_options['right']
        app_options['right'] = tmpval

    pages = int(values['_PAGE_INPUT_'])
    if pages < 1:
        reset_options();
    else:
        app_options['pages'] = pages

    app_options['next_key'] = values['_NEXT_KEY_']
    app_options['ctrl_key'] = values['_CTRL_KEY_']
    app_options['alt_key'] = values['_ALT_KEY_']
    app_options['shit_key'] = values['_SHIFT_KEY_']

    hwnd = app_options['hwnds'][idx]
    window.TKroot.wm_attributes("-topmost", 0)
    start_capture(hwnd)


def start_capture(hwnd):
    time.sleep(1)
    set_active_window(hwnd)
    page = app_options['pages']
    h_foldername = "Capture_"
    h_filename = "capture_"
    x1 = app_options['left']
    y1 = app_options['top']
    x2 = app_options['right']
    y2 = app_options['bottom']
    ctrl = app_options['ctrl_key']
    alt = app_options['alt_key']
    shift = app_options['shit_key']
    time.sleep(2)

    folder_name = h_foldername + "_" + str(datetime.datetime.now().strftime("%Y%m%d%H%M%S"))
    dlpath = app_options['downloads'] + os.sep + folder_name
    os.mkdir(dlpath)

    for p in range(page):
        out_filename = h_filename + "_" + str(p + 1).zfill(4) + '.png'
        s = pyautogui.screenshot(region=(x1, y1, x2 - x1, y2 - y1))
        s.save(dlpath + os.sep + out_filename)
        if ctrl or alt or shift:
            if ctrl:
                pyautogui.keyDown('ctrl')
            if alt:
                pyautogui.keyDown('alt')
            if shift:
                pyautogui.keyDown('shift')

            pyautogui.keyDown(app_options['next_key'])
            pyautogui.keyUp(app_options['next_key'])

            if ctrl:
                pyautogui.keyUp('ctrl')
            if alt:
                pyautogui.keyUp('alt')
            if shift:
                pyautogui.keyUp('shift')
        else:
            pyautogui.press(app_options['next_key'])

        set_active_window(hwnd)
        time.sleep(random.randint(1, 3))

    reset_options()
    time.sleep(1)
    app_options['window'].TKroot.wm_attributes("-topmost", 1)


def init_mouse_xy(top):
    print("start")
    try:
        while True:
            if ctypes.windll.user32.GetAsyncKeyState(0x01) == 0x8000:
                x, y = pyautogui.position()
                window = app_options['window']
                print(str(x) + ':' + str(y))

                if top:
                    app_options['top'] = y
                    app_options['left'] = x
                    window['_TOPLEFT_EDIT_'].update(value='Top:' + str(y) + '  ' + 'Left:' + str(x))
                else:
                    app_options['bottom'] = y
                    app_options['right'] = x
                    window['_BOTTOMRIGHT_EDIT_'].update(value='Bottom:' + str(y) + '  ' + 'Right:' + str(x))

                break

    except KeyboardInterrupt:
        reset_options()
        print('error')


def create_threading(top):
    threading.Thread(target=init_mouse_xy, args=(top,), daemon=True).start()


def set_active_window(hwnd):
    miniwin = False
    tup = win32gui.GetWindowPlacement(hwnd)
    if tup[1] == win32con.SW_SHOWMINIMIZED:
        miniwin = True
    if miniwin:
        win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
    win32gui.SetForegroundWindow(hwnd)


def get_windows():
    window = app_options['window']
    app_options['windows'].clear()
    app_options['hwnds'].clear()

    win32gui.EnumWindows(enum_handler, None)
    lst = []
    for wnd in app_options['windows']:
        lst.append(wnd)

    window['_WINDOWS_'].Update(values=lst)


def enum_handler(hwnd, ttl):
    if win32gui.IsWindowVisible(hwnd):
        ttllen = win32gui.GetWindowTextLength(hwnd)
        if 0 < ttllen:
            app_options['windows'].append(win32gui.GetWindowText(hwnd))
            app_options['hwnds'].append(hwnd)


def create_gui():
    sg.theme('Default')

    pyauto_keys = ['!', '#', '$', '%', '&', '(',
                   ')', '*', '+', ',', '-', '.', '/', '0', '1',
                   '2', '3', '4', '5', '6', '7',
                   '8', '9', ':', ';', '<', '=', '>', '?', '@',
                   '[', '\\', ']', '^', '_', '`',
                   'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i',
                   'j', 'k', 'l', 'm', 'n', 'o',
                   'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x',
                   'y', 'z', '{', '|', '}', '~',
                   'accept', 'add', 'alt', 'altleft',
                   'altright', 'apps', 'backspace',
                   'browserback', 'browserfavorites',
                   'browserforward', 'browserhome',
                   'browserrefresh', 'browsersearch',
                   'browserstop', 'capslock', 'clear',
                   'convert', 'ctrl', 'ctrlleft', 'ctrlright',
                   'decimal', 'del', 'delete',
                   'divide', 'down', 'end', 'enter', 'esc',
                   'escape', 'execute', 'f1', 'f10',
                   'f11', 'f12', 'f13', 'f14', 'f15', 'f16',
                   'f17', 'f18', 'f19', 'f2', 'f20',
                   'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5',
                   'f6', 'f7', 'f8', 'f9',
                   'final', 'fn', 'hanguel', 'hangul', 'hanja',
                   'help', 'home', 'insert', 'junja',
                   'kana', 'kanji', 'launchapp1', 'launchapp2',
                   'launchmail',
                   'launchmediaselect', 'left', 'modechange',
                   'multiply', 'nexttrack',
                   'nonconvert', 'num0', 'num1', 'num2', 'num3',
                   'num4', 'num5', 'num6',
                   'num7', 'num8', 'num9', 'numlock',
                   'pagedown', 'pageup', 'pause', 'pgdn',
                   'pgup', 'playpause', 'prevtrack', 'print',
                   'printscreen', 'prntscrn',
                   'prtsc', 'prtscr', 'return', 'right',
                   'scrolllock', 'select', 'separator',
                   'shift', 'shiftleft', 'shiftright', 'sleep',
                   'space', 'stop', 'subtract', 'tab',
                   'up', 'volumedown', 'volumemute', 'volumeup',
                   'win', 'winleft', 'winright', 'yen',
                   'command', 'option', 'optionleft',
                   'optionright']
    layout = [
        [sg.Button('Get Terget Window',
                   button_color=('white', 'royalblue'), size=(24, 2), key='_GET_WINDOW_BUTTON_')],
        [sg.Listbox((), size=(50, 10), key='_WINDOWS_')],
        [sg.T('', font='any 1')],
        [sg.Text("Click Capture Area    (empty=full)")],
        [
            sg.Button('Set Top-Left', button_color=('white', 'royalblue'), size=(20, 1), key='_SET_TOPLEFT_'),
            sg.Input(default_text='', size=(25, 1), key='_TOPLEFT_EDIT_', readonly=True)
        ],
        [
            sg.Button('Set Bottom-Right', button_color=('white', 'royalblue'), size=(20, 1), key='_SET_BOTTOMRIGHT_'),
            sg.Input(default_text='', size=(25, 1), key='_BOTTOMRIGHT_EDIT_', readonly=True)
        ],
        [sg.T('', font='any 1')],
        [sg.T('', font='any 1')],
        [sg.Text("Pages "), sg.Input(default_text='1', size=(8, 1), enable_events=True, key="_PAGE_INPUT_")],
        [sg.T('', font='any 1')],
        [sg.T('', font='any 1')],
        [sg.Text("Next Key", font=("Arial", 13))],

        [sg.Combo(pyauto_keys,
                  default_value='right', key='_NEXT_KEY_', size=(30, 1))],
        [sg.Combo(['', 'ctrl'],
                  default_value='', key='_CTRL_KEY_', size=(8, 1))],
        [sg.Combo(['', 'alt'],
                  default_value='', key='_ALT_KEY_', size=(8, 1))],
        [sg.Combo(['', 'shift'],
                  default_value='', key='_SHIFT_KEY_', size=(8, 1))],

        [sg.T('', font='any 1')],
        [sg.T('', font='any 1')],
        [sg.Button('Start Capture',
                   button_color=('white', 'royalblue'), size=(44, 2), key='_START_CAP_')
         ],
        [sg.T('', font='any 1')],
    ]

    wnd = sg.Window('Screen Capture', layout, keep_on_top=True)
    app_options['window'] = wnd
    event_loop(wnd);


def event_loop(window):
    while True:  # Event Loop
        event, values = window.read()

        if event == '_GET_WINDOW_BUTTON_':
            get_windows();

        elif event == '_SET_TOPLEFT_':
            create_threading(True)

        elif event == '_SET_BOTTOMRIGHT_':
            create_threading(False)

        elif event == "_PAGE_INPUT_":
            text = values['_PAGE_INPUT_']
            if text == '':
                pass
            else:
                try:
                    value = int(text)
                    window['_PAGE_INPUT_'].update(value=value)
                except:
                    window['_PAGE_INPUT_'].update(value='')

        elif event == '_START_CAP_':
            init_capture(window, values)

        elif event in (sg.WIN_CLOSED, 'Exit'):
            break

    window.close()


def main():
    create_gui()


if __name__ == '__main__':
    main()
