import speech_recognition as sr
import os
import ctypes
import win32api
import win32con
import pyautogui
import subprocess

# Commands for running programs
commands = {
    "calculator": "calc",
    "explorer": "explorer",
    "chrome": "start chrome",
    "paint": "mspaint",
    "minecraft": r"C:\XboxGames\Minecraft Launcher\Content\Minecraft.exe",
    "english": "change_language_english",
    "russian": "change_language_russian",
    "ukrainian": "change_language_ukrainian",
    "close window": "close_window",
    "shut down computer": "shut_down_computer",
    "cancel shutdown": "cancel_shutdown",
    "exit script": "exit"
}

# Language code (English, Russian, Ukrainian)
ENGLISH_LANG_ID = 0x0409
RUSSIAN_LANG_ID = 0x0419
UKRAINIAN_LANG_ID = 0x0422

def change_language_english():
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    ctypes.windll.user32.PostMessageW(hwnd, win32con.WM_INPUTLANGCHANGEREQUEST, 0, ENGLISH_LANG_ID)
    print("Keyboard language switched to English")

def change_language_russian():
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    ctypes.windll.user32.PostMessageW(hwnd, win32con.WM_INPUTLANGCHANGEREQUEST, 0, RUSSIAN_LANG_ID)
    print("Keyboard language switched to Russian")

def change_language_ukrainian():
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    ctypes.windll.user32.PostMessageW(hwnd, win32con.WM_INPUTLANGCHANGEREQUEST, 0, UKRAINIAN_LANG_ID)
    print("Keyboard language switched to Ukrainian")

def close_window():
    pyautogui.hotkey('alt', 'f4')  # Press Alt + F4 to close the window
    print("Closing active window")

def shutdown_computer():
    os.system("shutdown /s /t 600")  # Shutdown after 600 seconds (10 minutes)
    print("Shutdown scheduled in 10 minutes")

def cancel_shutdown():
    os.system("shutdown /a")  # Cancel shutdown
    print("Shutdown canceled")

# Voice input
recognizer = sr.Recognizer()

while True:
    with sr.Microphone() as source:
        print("Listening for command...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

        try:
            command = recognizer.recognize_google(audio, language="en-US").lower()
            print("You said:", command)

            if command in commands:
                if command == "english":
                    change_language_english()
                elif command == "russian":
                    change_language_russian()
                elif command == "ukrainian":
                    change_language_ukrainian()
                elif command == "close window":
                    close_window()
                elif command == "shut down computer":
                    shutdown_computer()
                elif command == "cancel shutdown":
                    cancel_shutdown()
                elif command == "exit script":
                    print("Exiting script...")
                    break
                else:
                    subprocess.Popen(commands[command], shell=True)  # Launching a program in a new process
                    print(f"Launching {command}")
            else:
                print("Unknown command")

        except sr.UnknownValueError:
            print("Could not understand the command")
        except sr.RequestError:
            print("Error connecting to Google")
