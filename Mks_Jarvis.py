# -------------------- Jarvis with Safe Tab Close (Fixed) --------------------
import tkinter as tk
from PIL import Image, ImageTk
import pyttsx3
import speech_recognition as sr
import pywhatkit
import datetime
import threading
import pygame
import os
import webbrowser
import psutil
import queue
import socket
import pyautogui
import requests
import wikipedia
import difflib
import math

# -------------------- SETUP --------------------
engine = pyttsx3.init()
engine.setProperty('rate', 150)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
pygame.init()
pygame.mixer.init()

speech_queue = queue.Queue()
def speak_worker(q):
    while True:
        text = q.get()
        if text is None:
            break
        engine.say(text)
        engine.runAndWait()
        q.task_done()
threading.Thread(target=speak_worker, args=(speech_queue,), daemon=True).start()

def speak_async(text):
    print("Jarvis:", text)
    speech_queue.put(text)

# -------------------- Global Vars --------------------
running = False
angle = 0  # For rotating arc reactor

# -------------------- Command Map --------------------
command_map = {
    "open youtube": "start chrome https://www.youtube.com",
    "close youtube": "close_tab",

    "open google": "start chrome https://www.google.com",
    "close google": "close_tab",

    "open chat gpt": "start chrome https://chat.openai.com",
    "close chat gpt": "close_tab",
   
    "open linkedin": "start chrome https://www.linkedin.com",
    "close linkedin": "close_tab",

    "open github": "start chrome https://www.github.com",
    "close github": "close_tab",

    "open gmail": "start chrome https://mail.google.com",
    "close gmail": "close_tab",
    "open mail": "start chrome https://mail.google.com",
    "close mail": "close_tab",

    # Local app commands
    "open notepad": "start notepad",
    "close notepad": "taskkill /f /im notepad.exe",

    "open calculator": "start calc",
    "close calculator": "taskkill /f /im Calculator.exe",

    "open file explorer": "start explorer",
    "close file explorer": "taskkill /f /im explorer.exe",

    "open word": "start winword",
    "close word": "taskkill /f /im winword.exe",

    "open excel": "start excel",
    "close excel": "taskkill /f /im excel.exe",

    "open powerpoint": "start powerpnt",
    "close powerpoint": "taskkill /f /im powerpnt.exe",

    "open edge": "start msedge",
    "close edge": "taskkill /f /im msedge.exe",

    "open chrome": "start chrome",
    "close chrome": "taskkill /f /im chrome.exe",

    "open vs code": "code",
    "close vs code": "taskkill /f /im Code.exe",

    "open photos": "start ms-photos:",
    "close photos": "taskkill /f /im Microsoft.Photos.exe",

    "open gallery": "start ms-photos:",
    "close gallery": "taskkill /f /im Microsoft.Photos.exe",

    "shutdown": "shutdown /s /t 1",
    "restart": "shutdown /r /t 1",
    "lock system": "rundll32.exe user32.dll,LockWorkStation",

    "turn on wifi": "netsh interface set interface \"Wi-Fi\" enabled",
    "turn off wifi": "netsh interface set interface \"Wi-Fi\" disabled",
    "open wifi settings": "start ms-settings:network",

    "turn on bluetooth": "powershell -Command \"Start-Service bthserv\"",
    "turn off bluetooth": "powershell -Command \"Stop-Service bthserv\"",
    "open bluetooth settings": "start ms-settings:bluetooth",

    "turn on camera": "start microsoft.windows.camera:",
    "turn off camera": "taskkill /f /im WindowsCamera.exe",

    "turn on mic": "nircmd.exe mutesysvolume 0 \"Microphone\"",
    "turn off mic": "nircmd.exe mutesysvolume 1 \"Microphone\"",
    "open sound settings": "control mmsys.cpl sounds",

    "next music": "powershell -Command \"(New-Object -ComObject WScript.Shell).SendKeys('{MEDIA_NEXT_TRACK}')\"",
    "pause music": "powershell -Command \"(New-Object -ComObject WScript.Shell).SendKeys('{MEDIA_PLAY_PAUSE}')\"",
    "previous music": "powershell -Command \"(New-Object -ComObject WScript.Shell).SendKeys('{MEDIA_PREV_TRACK}')\"",

    "take screenshot": (
        r'powershell -Command "$t=Get-Date -UFormat %%Y%%m%%d_%%H%%M%%S; '
        r'nircmd.exe savescreenshotfull '
        r'\"$env:USERPROFILE\\Pictures\\screenshot_$t.png\""'
    ),
        "open task manager": "start taskmgr",
    "open control panel": "start control",
    "open command prompt": "start cmd",
    "open run": "start shell:RunDialog",
    "open stack overflow": "start chrome https://stackoverflow.com",
    "open reddit": "start chrome https://www.reddit.com",
    "open netflix": "start chrome https://www.netflix.com",
    "open amazon": "start chrome https://www.amazon.in",

    "open task manager": "start taskmgr",
    "close task manager": "taskkill /f /im Taskmgr.exe",

    "open control panel": "start control",
    "close control panel": "taskkill /f /im control.exe",

    "open command prompt": "start cmd",
    "close command prompt": "taskkill /f /im cmd.exe",

    "open run": "start shell:RunDialog",
    "close run": "pyautogui.hotkey('ctrl', 'w')",  # Alt+F4 is risky, use safe close

    "open stack overflow": "start chrome https://stackoverflow.com",
    "close stack overflow": "pyautogui.hotkey('ctrl', 'w')",

    "open reddit": "start chrome https://www.reddit.com",
    "close reddit": "pyautogui.hotkey('ctrl', 'w')",

    "open netflix": "start chrome https://www.netflix.com",
    "close netflix": "pyautogui.hotkey('ctrl', 'w')",

    "open amazon": "start chrome https://www.amazon.in",
    "close amazon": "pyautogui.hotkey('ctrl', 'w')",

    "mute volume": "nircmd.exe mutesysvolume 1",
    "unmute volume": "nircmd.exe mutesysvolume 0",
    "increase volume": "nircmd.exe changesysvolume 5000",
    "decrease volume": "nircmd.exe changesysvolume -5000",

    "minimize window": "minimize_window",
    "restore window": "restore_window"

}

# -------------------- Take Command --------------------
def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        print("Listening...")
        try:
            audio = r.listen(source)
            text = r.recognize_google(audio).lower()
            print("User:", text)
            return text
        except:
            return ""

# -------------------- UI Controls --------------------
def start_ui():
    global running
    running = True
    try:
        pygame.mixer.music.load("E:/OneDrive/Documents/Desktop/Voice  Assistant/assets/jarvis.mp3")
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(1)
    except:
        pass
    speak_async("Hello Boss, how can I help you.")
    threading.Thread(target=listen_loop, daemon=True).start()
    update_arc_reactor()

def stop_ui():
    global running
    running = False
    speak_async("Jarvis stopped.")

# -------------------- Listen Loop --------------------
def listen_loop():
    while running:
        query = take_command()
        if query:
            handle_command(query)

def handle_command(q):
    q = q.strip().lower()

    if q.startswith("play "):
        song = q.replace("play ", "")
        speak_async(f"Playing {song} on YouTube")
        webbrowser.open(f"https://www.youtube.com/results?search_query={song}")
        return

    if any(q.startswith(tag) for tag in ["what", "who", "where", "when", "why", "how", "which"]):
        speak_async("Searching Google")
        webbrowser.open(f"https://www.google.com/search?q={q}")
        return

    for key, act in command_map.items():
        if key in q:
            speak_async(f"Executing {key}")
            if act == "close_tab":
                pyautogui.hotkey('ctrl', 'w')
            elif act == "minimize_window":
                os.system('powershell -Command "(New-Object -ComObject Shell.Application).MinimizeAll()"')
            elif act == "restore_window":
                os.system('powershell -Command "(New-Object -ComObject Shell.Application).UndoMinimizeAll()"')
            elif act.startswith("http") or act.startswith("start"):
                os.system(act)
            elif isinstance(act, str):
                os.system(act)
            return


# -------------------- GUI --------------------
root = tk.Tk()
root.title("Mks Jarvis")
root.geometry("1280x720")
root.resizable(False, False)

canvas = tk.Canvas(root, width=1280, height=720, highlightthickness=0, bg='black')
canvas.pack()
bg_img = Image.open("E:/OneDrive/Documents/Desktop/Voice  Assistant/assets/background.jpeg").resize((1280, 720))
bg = ImageTk.PhotoImage(bg_img)
canvas.create_image(0, 0, anchor='nw', image=bg)

center_x, center_y, radius = 640, 360, 80
def update_arc_reactor():
    global angle
    canvas.delete("arc")
    for i in range(12):
        a = math.radians(angle + i * 30)
        x = center_x + radius * math.cos(a)
        y = center_y + radius * math.sin(a)
        canvas.create_oval(x-5, y-5, x+5, y+5, fill="cyan", outline="", tags="arc")
    angle = (angle + 5) % 360
    if running:
        root.after(50, update_arc_reactor)

hud_x, hud_y = 200, 50
font_hud = ('Consolas', 12, 'bold')
def update_hud():
    cpu = psutil.cpu_percent()
    ram = psutil.virtual_memory().percent
    now = datetime.datetime.now().strftime("%I:%M:%S %p")
    canvas.delete('hud')
    canvas.create_text(hud_x, hud_y, text=f"CPU: {cpu:.0f}%", fill='red', font=font_hud, tags='hud', anchor='nw')
    canvas.create_text(hud_x, hud_y+20, text=f"RAM: {ram:.0f}%", fill='red', font=font_hud, tags='hud', anchor='nw')
    canvas.create_text(hud_x, hud_y+40, text=f"Time: {now}", fill='red', font=font_hud, tags='hud', anchor='nw')
    root.after(1000, update_hud)
update_hud()

ctrl_x, ctrl_y = 50, 480
folders = ["Documents", "Downloads", "Music", "Pictures"]
for i, name in enumerate(folders[::-1]):
    btn = tk.Button(root, text=name, fg='white', bg='darkred', font=('Arial', 10),
                    command=lambda n=name: os.startfile(f"C:/Users/{os.getlogin()}/{n}"))
    y_pos = ctrl_y - (i+1)*35
    canvas.create_window(ctrl_x, y_pos, window=btn, width=120, height=30, anchor='nw')

btn_start = tk.Button(root, text="Start Jarvis", fg='white', bg='green', font=('Arial', 12, 'bold'), command=start_ui)
btn_stop = tk.Button(root, text="Stop Jarvis", fg='white', bg='red', font=('Arial', 12, 'bold'), command=stop_ui)
canvas.create_window(50, 660, window=btn_start, width=140, height=40, anchor='nw')
canvas.create_window(1090, 660, window=btn_stop, width=140, height=40, anchor='nw')

root.mainloop()
