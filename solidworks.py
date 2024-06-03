import time
import psutil
import pygetwindow as gw
from pypresence import Presence
import configparser
import os
import signal
from PIL import Image
import sys
import win32api
import win32con
import win32gui
import tkinter as tk
from tkinter import messagebox


# Discord Rich Presence ID
client_id = '1124229378105167973'
CONFIG_FILE = 'config.ini'


def is_solidworks_running():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'SLDWORKS.exe':
            return True
    return False


def get_solidworks_display_name():
    try:
        solidworks_window = gw.getWindowsWithTitle('SOLIDWORKS Premium')[0]
        window_title = solidworks_window.title

        # Check if no document is open
        if '[' not in window_title:
            return 'Idle', None, None

        # Extract the file name from the window title
        start_index = window_title.index('[') + 1
        end_index = window_title.index(']')
        file_name = window_title[start_index:end_index]

        # Clean up the file name by removing unwanted characters
        file_name = file_name.strip('*').strip()

        if file_name.lower().endswith('.sldprt'):
            return 'Making things', 'part', file_name
        elif file_name.lower().endswith('.sldasm'):
            return 'Assembling things', 'assembly', file_name
        elif file_name.lower().endswith('.slddrw'):
            return 'Drawing things', 'drawing', file_name
        else:
            return 'Idle', None, None
    except IndexError:
        print("SolidWorks not found. Clearing presence.")
    return None, None, None


def update_discord_presence(idle_text, part_text, assembly_text, drawing_text):
    presence = Presence(client_id)
    presence.connect()
    start_time = None

    while True:
        try:
            if is_solidworks_running():
                state, small_image, file_name = get_solidworks_display_name()
                if state:
                    print_previous()
                    print(f"State: {idle_text if state == 'Idle' else part_text if state == 'Making things' else assembly_text if state == 'Assembling things' else drawing_text}")
                    print(f"Small Image: {small_image}")
                    print(f"File Name: {file_name}")
                    if state == 'Idle':
                        state = idle_text
                    elif state == 'Making things':
                        state = part_text
                    elif state == 'Assembling things':
                        state = assembly_text
                    elif state == 'Drawing things':
                        state = drawing_text
                    presence.update(
                        details=state,
                        state=file_name or 'No open file.',
                        large_image="sldworks",
                        small_image=small_image
                    )
                else:
                    presence.clear()
            else:
                presence.clear()
                start_time = None

            time.sleep(5)

        except ValueError:
            presence.clear()
            start_time = None

        time.sleep(5)


def save_config(idle_text, part_text, assembly_text, drawing_text):
    config = configparser.ConfigParser()
    config['DEFAULT'] = {
        'idle_text': idle_text,
        'part_text': part_text,
        'assembly_text': assembly_text,
        'drawing_text': drawing_text
    }
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)


def load_config():
    config = configparser.ConfigParser()
    if not os.path.exists(CONFIG_FILE):
        return None, None, None, None
    config.read(CONFIG_FILE)
    if 'DEFAULT' in config:
        return (
            config['DEFAULT'].get('idle_text'),
            config['DEFAULT'].get('part_text'),
            config['DEFAULT'].get('assembly_text'),
            config['DEFAULT'].get('drawing_text')
        )
    return None, None, None, None


def minimize_to_tray(hwnd, msg):
    tray_menu = win32gui.CreatePopupMenu()
    win32gui.AppendMenu(tray_menu, win32con.MF_STRING, 1024, "Exit")

    tray_icon = win32gui.Shell_NotifyIcon(
        win32gui.NIM_ADD,
        (
            hwnd,
            0,
            win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
            0,
            0,
            "",
            Image.open("icon.png").tobytes(),
            msg,
            1024,
            tray_menu
        )
    )

    def on_tray_icon_event(hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONUP:
            win32gui.DestroyWindow(tray_icon[0])
        elif lparam == win32con.WM_RBUTTONUP:
            menu = win32gui.CreatePopupMenu()
            win32gui.AppendMenu(menu, win32con.MF_STRING, 1024, "Exit")
            win32gui.SetForegroundWindow(hwnd)
            pos = win32gui.GetCursorPos()
            win32gui.TrackPopupMenu(
                menu,
                win32con.TPM_LEFTALIGN,
                pos[0],
                pos[1],
                0,
                hwnd,
                None
            )
            win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)

    win32gui.PumpMessages()


def set_exit_handler():
    def handler(sig, frame):
        minimize_to_tray(None, None)

    signal.signal(signal.SIGINT, handler)


def print_previous():
    # Clear the previous output
    print('\033[F\033[K' * 3, end='')


def save_config_and_start_presence():
    global idle_text_entry, part_text_entry, assembly_text_entry, drawing_text_entry
    idle_text = idle_text_entry.get()
    part_text = part_text_entry.get()
    assembly_text = assembly_text_entry.get()
    drawing_text = drawing_text_entry.get()

    if idle_text and part_text and assembly_text and drawing_text:
        save_config(idle_text, part_text, assembly_text, drawing_text)
        root.destroy()
        set_exit_handler()
        update_discord_presence(idle_text, part_text, assembly_text, drawing_text)
    else:
        messagebox.showerror("Error", "Please enter all the texts.")


# Create the main window
root = tk.Tk()
root.title("SolidWorks Discord Presence")
root.geometry("400x200")

# Load the configuration
idle_text, part_text, assembly_text, drawing_text = load_config()

# Create and place the text labels
idle_label = tk.Label(root, text="Idle Text:")
idle_label.pack()
idle_text_entry = tk.Entry(root)
idle_text_entry.pack()
if idle_text:
    idle_text_entry.insert(0, idle_text)

part_label = tk.Label(root, text="Part Text:")
part_label.pack()
part_text_entry = tk.Entry(root)
part_text_entry.pack()
if part_text:
    part_text_entry.insert(0, part_text)

assembly_label = tk.Label(root, text="Assembly Text:")
assembly_label.pack()
assembly_text_entry = tk.Entry(root)
assembly_text_entry.pack()
if assembly_text:
    assembly_text_entry.insert(0, assembly_text)

drawing_label = tk.Label(root, text="Drawing Text:")
drawing_label.pack()
drawing_text_entry = tk.Entry(root)
drawing_text_entry.pack()
if drawing_text:
    drawing_text_entry.insert(0, drawing_text)

# Create and place the save button
save_button = tk.Button(root, text="Save and Start", command=save_config_and_start_presence)
save_button.pack()

# Start the Tkinter event loop
root.mainloop()
