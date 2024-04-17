import subprocess
import pyautogui
import json
import pywhatkit
import pyttsx3
import google.generativeai as genai
import send2trash
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import ctypes
import time
import tkinter as tk
from tkinter import scrolledtext
import shutil
from pyjokes import pyjokes
import cv2
import numpy as np
import winshell
import pygetwindow as gw
global qa_data
def create_note():
    speak("What would you like to name the note?")
    note_name = takeCommand()
    speak("What would you like to write in the note?")
    note_content = takeCommand()
    if note_name != "None" and note_content != "None":
        directory = r'C:\Users\lamra\Documents\note'
        if not os.path.exists(directory):
            os.makedirs(directory)
        file_name = os.path.join(directory, f"{note_name}.txt")
        with open(file_name, "w") as file:
            file.write(note_content)
        speak("Note created successfully.")


def open_programs_on_desktop():
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    desktop_files = [f for f in os.listdir(desktop_path) if f.endswith('.lnk')]

    if not desktop_files:
        speak("No shortcut files found on the desktop.")
        return

    speak("Here are the available programs on the desktop:")
    for i, file in enumerate(desktop_files, 1):
        speak(f"{i}. {os.path.splitext(file)[0]}")
    for i, file in enumerate(desktop_files, 1):
        choice = input(f"Do you want to open {os.path.splitext(file)[0]}? (yes/no): ")
        if choice.lower() == 'yes':
            file_path = os.path.join(desktop_path, file)
            subprocess.Popen(file_path)
            break
        elif choice.lower() == 'no':
            continue
        else:
            speak("Invalid choice. Please answer 'yes' or 'no'.")
            return
    else:
        speak("No more programs to open.")
def read_note():
    directory = r'C:\Users\lamra\Documents\note'
    files = [f for f in os.listdir(directory) if f.endswith('.txt')]
    if not files:
        speak("No note files found.")
        return

    speak("Here are the available notes:")
    for i, file in enumerate(files, 1):
        speak(f"{i}. {os.path.splitext(file)[0]}")

    valid_choices = [str(i) for i in range(1, len(files) + 1)]
    choice = input("Please enter the number corresponding to the note you want to read: ")

    if choice in valid_choices:
        file_name = os.path.join(directory, files[int(choice) - 1])
        try:
            with open(file_name, "r") as file:
                note_content = file.read()
            speak(f"Here is the content of the note {os.path.splitext(files[int(choice) - 1])[0]}:")
            speak(note_content)
        except FileNotFoundError:
            speak("Note not found.")
    else:
        speak("Invalid choice.")
def maximize_windows():
    for window in gw.getAllWindows():
        window.maximize()
def minimize_windows():
    for window in gw.getAllWindows():
        window.minimize()
def Camera():
    cam = cv2.VideoCapture(0)

    cam.set(3, 740)
    cam.set(4, 580)

    classNames = []
    classFile = 'coco.names'

    with open(classFile, 'rt') as f:
        classNames = f.read().rstrip('\n').split('\n')

    configPath = 'ssd_mobilenet_v3_large_coco_2020_01_14.pbtxt'
    weightpath = 'frozen_inference_graph.pb'

    net = cv2.dnn_DetectionModel(weightpath, configPath)
    net.setInputSize(320, 230)
    net.setInputScale(1.0 / 127.5)
    net.setInputMean((127.5, 127.5, 127.5))
    net.setInputSwapRB(True)

    while True:
        success, img = cam.read()
        classIds, confs, bbox = net.detect(img, confThreshold=0.5)
        print(classIds, bbox)

        if len(classIds) != 0:
            for classId, confidence, box in zip(classIds.flatten(), confs.flatten(), bbox):
                cv2.rectangle(img, box, color=(0, 255, 0), thickness=2)
                cv2.putText(img, classNames[classId - 1], (box[0] + 10, box[1] + 20),
                            cv2.FONT_HERSHEY_COMPLEX, 1, (0, 255, 0), thickness=2)

        cv2.imshow('Output', img)
        if cv2.waitKey(1) == ord('q'):
            break

    cam.release()
    cv2.destroyAllWindows()

def move_mouse(direction):
    if direction == "up":
        pyautogui.move(0, -50)
    elif direction == "down":
        pyautogui.move(0, 50)
    elif direction == "left":
        pyautogui.move(-50, 0)
    elif direction == "right":
        pyautogui.move(50, 0)
    elif direction == "scroll up":
        pyautogui.press('up')
    elif direction == "scroll down":
        pyautogui.press('down')
    elif direction == "scroll right":
        pyautogui.press('right')
    elif direction == "scroll left":
        pyautogui.press('left')
    elif direction == "right click":
        pyautogui.click(button='right')
    elif direction == "click":
        pyautogui.click(button='left')
    elif direction == "press left click":
        pyautogui.mouseDown(button='left')
    elif direction == "release left click":
        pyautogui.mouseUp(button='left')

def restore_active_window():
    active_window = gw.getActiveWindow()
    if active_window:
        active_window.restore()
    else:
        print("No active window found.")


def mouse_interaction():
    print("You can now control the mouse. Say 'exit mouse' to stop.")
    speak("You can now control the mouse. Say 'exit mouse' to stop.")
    print("For a long left-click, say 'long'. To release the left-click, say 'release'.")
    speak("For a long left-click, say 'long'. To release the left-click, say 'release'.")
    print("To perform a left or right click, say 'left' or 'right' respectively.")
    speak("To perform a left or right click, say 'left' or 'right' respectively.")
    print("To use the arrows, use the arrow keys. Say 'go' followed by the direction (up, down, left, right).")
    speak("To use the arrows, use the arrow keys. Say 'go' followed by the direction (up, down, left, right).")
    print("To move the pointer continuously, say 'move' followed by the direction (up, down, left, right) .")
    speak("To move the pointer continuously, say ' move' followed by the direction (up, down, left, right) .")






    while True:
        command = takeCommand().lower()

        if 'exit mouse' in command:
            speak("Mouse control exited.")
            break
        elif 'move' in command:
            direction = extract_direction(command)
            if direction:
                move_mouse(direction)
            else:
                speak("Invalid move command. Please specify direction.")
        elif 'go' in command:
            direction = extract_direction(command)
            if direction:
                print(f"Pressing arrow key {direction}.")
                pyautogui.press(direction)
            else:
                print("Invalid direction. Please specify direction.")
        elif 'scroll up' in command:
            pyautogui.scroll(1000)
        elif 'scroll down' in command:
            pyautogui.scroll(-1000)
        elif 'right' in command:
            pyautogui.click(button='right')
        elif 'long' in command:
            pyautogui.mouseDown(button='left')
        elif 'release' in command:
            pyautogui.mouseUp(button='left')
        elif 'left' in command:
            pyautogui.click()
        elif 'maximize' in command:
            maximize_windows()
        elif 'restore' in command:
            restore_active_window()
        elif 'minimize' in command:
            minimize_windows()
        elif 'close' in command:
            pyautogui.hotkey('alt', 'f4')  # Close current window
        elif 'desktop' in command:
            pyautogui.click(x=1919, y=1050)  # Show desktop
            pyautogui.click(40, 40)  # Move the cursor to (40, 40)
        elif 'enter' in command:
            pyautogui.press('enter')
        elif 'task' in command:
            pyautogui.moveTo(x=50, y=1050)
            speak("You are now controlling the taskbar. Say 'exit task' to stop.")
        else:
            speak("Invalid command.")


def extract_direction(command):
    directions = ['up', 'down', 'left', 'right']
    for direction in directions:
        if direction in command:
            return direction
    return None

def ImgFile():
    img = cv2.imread('person.png')

    classNames = []
    classFile = 'coco.names'

    with open(classFile, 'rt') as f:
        classNames = f.read().rstrip('\n').split('\n')

    configPath = 'ssd_mobilenet_v3_large_coco_2020_01_14.pbtxt'
    weightpath = 'frozen_inference_graph.pb'

    net = cv2.dnn_DetectionModel(weightpath, configPath)
    net.setInputSize(320, 230)
    net.setInputScale(1.0 / 127.5)
    net.setInputMean((127.5, 127.5, 127.5))
    net.setInputSwapRB(True)

    classIds, confs, bbox = net.detect(img, confThreshold=0.5)
    print(classIds, bbox)

    for classId, confidence, box in zip(classIds.flatten(), confs.flatten(), bbox):
        cv2.rectangle(img, box, color=(0, 255, 0), thickness=2)
        cv2.putText(img, classNames[classId - 1], (box[0] + 10, box[1] + 20),
                    cv2.FONT_HERSHEY_COMPLEX, 1, (0, 255, 0), thickness=2)

    cv2.imshow('Output', img)
    cv2.waitKey(0)



def list_files(path):
    files = os.listdir(path)
    for i, file in enumerate(files):
        print(f"{i+1}. {file}")
    return files

def delete_file(files, path, choice):
    try:
        file_index = int(choice) - 1
        if 0 <= file_index < len(files):
            file_to_delete = os.path.join(path, files[file_index])
            send2trash.send2trash(file_to_delete)
            print(f"File '{files[file_index]}' sent to the recycle bin successfully.")
        else:
            print("Invalid choice.")
    except ValueError:
        print("Invalid choice. Please enter a number.")


def play_music():
    speak("Here you go with music")
    music_dir = "C:\\Users\\Public\\AppData\\Flixmate\\Flixmate downloads"
    songs = os.listdir(music_dir)

    for i, song in enumerate(songs, start=1):
        print(f"{i}. {song}")



    while True:
        try:
            choice = int(input("Please enter the number of the song you want to play: "))
            if 1 <= choice <= len(songs):
                chosen_song = os.path.join(music_dir, songs[choice - 1])
                speak(f"Playing {songs[choice - 1]}")
                os.startfile(chosen_song)
                break
            else:
                print("Invalid choice. Please enter a valid number.")
        except ValueError:
            print("Invalid input. Please enter a number.")
def set_wallpaper(image_path):
    SPI_SETDESKWALLPAPER = 20
    ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, image_path, 3)


def display_and_choose_image(folder):
    image_files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    for i, image in enumerate(image_files):
        print(f"{i + 1}. {image}")
    speak("Enter the number corresponding to the image you want to choose (0 to cancel)")
    choice = input("Enter the number corresponding to the image you want to choose (0 to cancel): ")
    if choice.isdigit():
        index = int(choice) - 1
        if 0 <= index < len(image_files):
            return os.path.join(folder, image_files[index])

    return None

def ai(prompt):
    #شوف انت بعتلك المشروع المبسط تع  api key الاخر لهو النهائي نبعتهولك بعد مانريقليه قدد لانو حاندير llm فيه باه يقدر يتواصل بطريقة جيدة جدا
    genai.configure(api_key="api key")


    generation_config = {
        "temperature": 0.9,
        "top_p": 1,
        "top_k": 1,
        "max_output_tokens": 2048,
    }

    safety_settings = [
        {
            "category": "HARM_CATEGORY_HARASSMENT",
            "threshold": "BLOCK_MEDIUM_AND_ABOVE"
        },
        {
            "category": "HARM_CATEGORY_HATE_SPEECH",
            "threshold": "BLOCK_MEDIUM_AND_ABOVE"
        },
        {
            "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
            "threshold": "BLOCK_MEDIUM_AND_ABOVE"
        },
        {
            "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
            "threshold": "BLOCK_MEDIUM_AND_ABOVE"
        },
    ]

    model = genai.GenerativeModel(model_name="gemini-1.0-pro",
                                  generation_config=generation_config,
                                  safety_settings=safety_settings)

    convo = model.start_chat(history=[
    ])

    convo.send_message(prompt)
    print(convo.last.text)
    return convo.last.text




def calculate(num1, operator, num2):
    if operator == "+":
        return num1 + num2
    elif operator == "-":
        return num1 - num2
    elif operator == "*":
        return num1 * num2
    elif operator == "/":
        try:
            return num1 / num2
        except ZeroDivisionError:
            return "Error: Division by zero"
    elif operator == "^":
        return num1 ** num2
    else:
        return "Error: Invalid operator"

def open_file(files, path, choice):
            try:
                file_index = int(choice) - 1
                if 0 <= file_index < len(files):
                    file_to_open = os.path.join(path, files[file_index])
                    os.startfile(file_to_open)
                    print(f"File '{files[file_index]}' opened successfully.")
                else:
                    print("Invalid choice.")
            except ValueError:
                print("Invalid choice. Please enter a number.")

def set_reminder(reminder_text, reminder_time):
    try:
        # Parse reminder time
        reminder_time_obj = datetime.datetime.strptime(reminder_time, "%H:%M")
        # Get current time
        current_time = datetime.datetime.now().time()
        # Calculate time difference
        time_diff = datetime.datetime.combine(datetime.date.today(), reminder_time_obj.time()) - datetime.datetime.combine(datetime.date.today(), current_time)
        # Calculate seconds until reminder
        seconds_until_reminder = time_diff.total_seconds()
        # Schedule reminder
        if seconds_until_reminder > 0:
            print(f"Reminder set for {reminder_time}: {reminder_text}")
            # Sleep until reminder time
            time.sleep(seconds_until_reminder)
            # Alert user at reminder time
            print(f"Reminder: {reminder_text}")
            speak(f" Reminder:You Have to {reminder_text}")
        else:
            print("Reminder time should be in the future.")
            speak("Reminder time should be in the future.")
    except ValueError:
        print("Invalid time format. Please use HH:MM format for the reminder time.")
        speak("Invalid time format. Please use HH:MM format for the reminder time.")



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def ai_interaction():
    speak("Would you like to write the prompt? ")
    print("Would you like to write the prompt? ")
    response = takeCommand().lower()
    if 'yes' in response:
        prompt = input("Enter the prompt: ")
    elif 'no' in response:
        speak("Please say the prompt.")
        prompt = takeCommand()
    else:
        speak("Invalid response. Please try again.")
        return

    ai_response = ai(prompt)
    speak(ai_response)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def custom_google_search(search_term):

    speak('Searching the web for "{}"...'.format(search_term))
    webbrowser.open('https://www.google.com/search?q=' + search_term)

def close_window(window_title):
        try:
            app = pyautogui.getWindowsWithTitle(window_title)[0]
            app.close()
        except IndexError:
            speak("Window not found.")

def list_available_actions():
                actions = {
                    "Search Wikipedia for information": ["search wikipedia"],
                    "Open YouTube, Google, facebook or instagram": ["open youtube", "open google", "open facebook , open intagram"],
                    "Play music": ["play music", "play song"],
                    "Check the current time": ["time", "what time is it"],
                    "Send emails": ["email", "send email", "mail"],
                    "Introduce myself or change my name": ["introduce yourself", "what's your name", "change my name"],
                    "Tell jokes": ["joke", "tell me a joke"],
                    "file explorer": ["open file explorer"],
                    "Change the desktop wallpaper": ["change background", "set wallpaper"],
                    "Search the web": ["search in google about", "how to"],
                    "Shutdown or log off your computer": ["shut down the computer", "log off", "sign out"],
                    "Detect colors": ["what is the color", "detect color"],
                    "Empty the recycle bin": ["empty recycle bin"],
                    "Use AI capabilities": ["mike ai"],
                    "Close windows": ["i want to close"],
                    "Perform object detection": ["object detection"],
                    "Perform color detection": ["color detection"],
                    "create notes or reading note": ["create a note , read note"],
                    "control mouse": ["control mouse"],
                    "opening files": ["open a file"],
                    "deleting files": ["delete a file"],
                    "closing a window": ["i want to close"],
                    "CMD": ["open cmd"],
                    "a reminder": ["set a reminder"],

                }

                print("Here are the tasks you can ask me to perform along with their keywords:")
                for action, keywords in actions.items():
                    print(f"- {action}: Keywords: {', '.join(keywords)}")
                    speak(f"{action}. Keywords: {', '.join(keywords)}")


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir !")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir !")
    else:
        speak("Good Evening Sir !")
    assname = ("mike 1 point o")
    speak("I am your Assistant")
    speak(assname)

def username():
    speak("Welcome Master")

    columns = shutil.get_terminal_size().columns
    print("#####################".center(columns))
    print("Welcome Master".center(columns))
    print("#####################".center(columns))
    speak("How can i Help you, Sir")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-us')
        print(f"User said: {query}")
    except Exception as e:
        print(e)
        print("Unable to Recognize your voice.")
        return "None"
    return query


def load_qa_data():
    try:
        with open('qa_data.json', 'r') as read_file:
            return json.load(read_file)
    except FileNotFoundError:
        return []

def save_qa_data(qa_data):
    with open('qa_data.json', 'w') as save_file:
        json.dump(qa_data, save_file, indent=4)

def load_qa_data():
    try:
        with open('qa_data.json', 'r') as read_file:
            return json.load(read_file)
    except FileNotFoundError:
        return []

def save_qa_data(qa_data):
    with open('qa_data.json', 'w') as save_file:
        json.dump(qa_data, save_file, indent=4)

def handle_unrecognized_query(query, qa_data):
    qa_data = load_qa_data()

    for entry in qa_data:
        if entry['query'] == query:
            if entry.get("is_function"):
                while True:
                    execute_or_modify = input("Do you want to execute or modify the saved function? (execute/modify): ").lower()
                    if execute_or_modify == 'execute':
                        try:
                            exec(entry["answer"], globals(), locals())
                        except Exception as e:
                            print(f"Error executing function: {e}")
                        break
                    elif execute_or_modify == 'modify':
                        root = tk.Tk()
                        root.title("Modify Python Function")
                        function_code = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=30)
                        function_code.insert(tk.END, entry['answer'])
                        function_code.pack()
                        function_code.focus_set()

                        def save_modified_function():
                            modified_function_code = function_code.get("1.0", tk.END)
                            entry['answer'] = modified_function_code
                            save_qa_data(qa_data)
                            print("Function modified and saved successfully!")
                            root.destroy()

                        save_button = tk.Button(root, text="Save", command=save_modified_function)
                        save_button.pack()

                        root.mainloop()
                        break
                    else:
                        print("Invalid option. Please enter 'execute' or 'modify'.")

            else:
                # Print the string answer
                print(f"Answer (string): {entry['answer']}")
                speak(f"{entry['answer']}")
            return

    print(f"Unrecognized query: {query}")
    store_new_entry = input("Would you like to store a new query and answer? (yes/no): ").lower()
    if store_new_entry == 'yes':
        is_function = input("Is the answer a function? (yes/no): ").lower()
        if is_function == 'yes':
            # Open a larger window for writing Python code
            root = tk.Tk()
            root.title("Python Function")
            function_code = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=40)
            function_code.pack()
            function_code.focus_set()

            def save_function():
                new_function_code = function_code.get("1.0", tk.END)
                qa_data.append({"query": query, "answer": new_function_code, "is_function": True})
                save_qa_data(qa_data)
                print("New function saved successfully!")
                root.destroy()

            save_button = tk.Button(root, text="Save", command=save_function)
            save_button.pack()

            root.mainloop()

        else:
            new_answer = input("Enter the answer (string): ")
            qa_data.append({"query": query, "answer": new_answer, "is_function": False})
            save_qa_data(qa_data)
            print("New entry saved successfully!")
    else:
        print("Continuing with the loop...")
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttl()

    server.login('email', 'password')
    server.sendmail('lamraouibasset11@gmail.com', to, content)
    server.close()

if __name__ == '__main__':
    clear = lambda: os.system('cls')
    clear()
    wishMe()
    username()
    while True:
        query = takeCommand().lower()
        if 'list available actions' in query or 'what can you do' in query:
            list_available_actions()
        elif 'search wikipedia about' in query :
            speak('Searching Wikipedia...')
            query = query.replace("search wikipedia about", "")
            print("Searching Wikipedia for:", query)
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        elif 'open facebook' in query:
            speak("Here you go to facebook\n")
            os.startfile("https://www.facebook.com/")
        elif 'open instagram' in query:
            speak("Here you go to instagram\n")
            os.startfile("https://www.instagram.com/")
        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            os.startfile("http://www.youtube.com")
        elif 'open google' in query:
            speak("Here you go to Google\\n")
            os.startfile("http://www.google.com")

        elif 'play music' in query or 'play song' in query:
            play_music()
        elif 'time' in query:
            now = datetime.datetime.now()
            strTime = now.strftime("%H:%M:%S")
            print(f"[Current Time: {strTime}]")
            speak(f"Sir, the time is {strTime}")
        elif 'email' in query or'mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whome should i send")
                to = input()
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'introduce yourself' in query:
            speak("Greetings. I am mike, a comprehensive virtual assistant designed by abd el Basset lamraoui. I am here to assist you with a wide range of functions and provide you with the information and support you need.")
            print("Greetings. I am mike, a comprehensive virtual assistant designed by abd el Basset lamraoui. I am here to assist you with a wide range of functions and provide you with the information and support you need.")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")
        elif 'fine' in query or "good" in query :
            speak("It's good to know that your fine")
        elif "change my name " in query:
            query = query.replace("change my name to", "")
            assname = query
        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)
        elif "who made you" in query or "who created you" in query:
            speak("I have been created by basset.")
        elif 'joke' in query:
            joke = pyjokes.get_joke()
            print(joke)
            speak(joke)
        elif "who i am" in query:
            speak("If you talk then definitely your human.")
        elif "why you came to world" in query:
            speak("Thanks to basset. further It's a secret")
        elif 'presentation' in query:
            speak("opening Power Point presentation")
            power = r"C:\\Users\\lamra\\Downloads\\project.pptx"
            os.startfile(power)
        elif "who are you" in query:
            speak("I am your virtual assistant created by basset")
        elif 'reason for you' in query:
            speak("I was created as a Minor project by Mister basset ")
        elif 'open file explorer' in query:
            os.system('explorer')
        elif 'open a file' in query:
            speak("Enter the path to the directory:")
            path = input("Enter the path to the directory: ")
            if os.path.exists(path):
                print("Listing files in the directory:")
                speak("Listing files in the directory:")
                files = list_files(path)
                speak("Enter the number corresponding to the file you want to open:")
                choice = input("Enter the number corresponding to the file you want to open: ")
                open_file(files, path, choice)
            else:
                print("The specified path does not exist.")
        elif 'change background' in query:
            speak("Please enter the path to your background folder")
            background_folder = input("Please enter the path to your background folder (e.g., C:/Users/username/Pictures/Backgrounds): ")

            chosen_image = display_and_choose_image(background_folder)

            if chosen_image:
                speak(f"You chose image: {chosen_image}")
                print(f"You chose image: {chosen_image}")
                set_wallpaper(chosen_image)
                speak("Wallpaper set successfully.")
                print("Wallpaper set successfully.")
            else:
                speak("No image chosen.")
                print("No image chosen.")
        elif "how to" in query:
            query = query.replace("how to", "")
            if query:
                pywhatkit.playonyt(query)
            else:
                speak("Please specify what you want to know how to do.")
        elif "search in google about" in query:
            query = query.replace( "search in google about", "")
            custom_google_search(query)

        elif 'shut down the computer' in query:
            speak("Are you sure you want to shut down the computer?")
            confirm = takeCommand().lower()
            if confirm == 'yes' or confirm == 'shutdown':
             os.system("shutdown /s /t 1")
            else:
             speak("Shutdown canceled.")

        elif 'log off' in query or 'sign out' in query:
            speak("Are you sure you want to log off?")
            confirm = takeCommand().lower()
            if confirm == 'yes' or confirm == 'log off':
             os.system("shutdown /l")
            else:
                speak("Log off canceled.")
        elif 'exit' in query:
            speak("Thanks for using me! Have a good day.")
            exit()
        elif 'delete a file' in query:
            speak("Enter the path to the directory:")
            path = input("Enter the path to the directory: ")
            if os.path.exists(path):
                speak("Listing files in the directory:")
                print("Listing files in the directory:")
                files = list_files(path)
                os.startfile(path)
                speak("Enter the number corresponding to the file you want to delete:")
                choice = input("Enter the number corresponding to the file you want to delete: ")
                delete_file(files, path, choice)
            else:
                print("The specified path does not exist.")
        elif 'calculator' in query:
            while True:
                try:
                    num1 = float(input("Enter first number: "))
                    operator = input("Enter operator (+, -, *, /, ^): ")
                    if operator == '^':
                        num2 = float(input("Enter second number (exponent): "))
                    else:
                        num2 = float(input("Enter second number: "))
                    break
                except ValueError:
                    print("Invalid input. Please enter numbers only.")

            result = calculate(num1, operator, num2)
            if result is not None:
                print("Result:", result)
                speak("Result: " + str(result))
        elif 'open file explorer' in query:
            os.system('explorer')
        elif 'empty recycle bin' in query:
            try:
                speak("Are you sure you want to empty the recycle bin?")
                print("Are you sure you want to empty the recycle bin?")
                confirm = takeCommand().lower()
                if confirm == 'yes' or confirm == 'empty':
                    winshell.recycle_bin().empty(confirm=True, show_progress=False, sound=True)
                    speak("Recycle bin has been emptied.")
                else:
                    speak("Emptying recycle bin canceled.")
            except Exception as e:
                print(e)
                speak(
                    "There might be issues emptying the recycle bin. It could be full or corrupt. Please check your Recycle Bin settings.")
        elif 'mike use ai' in query:
            ai_interaction()
        elif 'i want to close ' in query:
            query = query.replace("i want to close ", "")
            window_title = query.strip()
            close_window(window_title)
        elif 'object detection' in query:
            Camera()
        elif 'color detection' in query:
            webcam = cv2.VideoCapture(0)

            while True:
                _, imageFrame = webcam.read()

                hsvFrame = cv2.cvtColor(imageFrame, cv2.COLOR_BGR2HSV)

                red_lower = np.array([136, 87, 111], np.uint8)
                red_upper = np.array([180, 255, 255], np.uint8)
                red_mask = cv2.inRange(hsvFrame, red_lower, red_upper)

                green_lower = np.array([25, 52, 72], np.uint8)
                green_upper = np.array([102, 255, 255], np.uint8)
                green_mask = cv2.inRange(hsvFrame, green_lower, green_upper)

                blue_lower = np.array([94, 80, 2], np.uint8)
                blue_upper = np.array([120, 255, 255], np.uint8)
                blue_mask = cv2.inRange(hsvFrame, blue_lower, blue_upper)

                kernel = np.ones((5, 5), "uint8")

                red_mask = cv2.dilate(red_mask, kernel)
                res_red = cv2.bitwise_and(imageFrame, imageFrame, mask=red_mask)

                green_mask = cv2.dilate(green_mask, kernel)
                res_green = cv2.bitwise_and(imageFrame, imageFrame, mask=green_mask)

                blue_mask = cv2.dilate(blue_mask, kernel)
                res_blue = cv2.bitwise_and(imageFrame, imageFrame, mask=blue_mask)

                contours, hierarchy = cv2.findContours(red_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

                for pic, contour in enumerate(contours):
                    area = cv2.contourArea(contour)
                    if area > 300:
                        x, y, w, h = cv2.boundingRect(contour)
                        imageFrame = cv2.rectangle(imageFrame, (x, y), (x + w, y + h), (0, 0, 255), 2)
                        cv2.putText(imageFrame, "Red Colour", (x, y), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 255))

                contours, hierarchy = cv2.findContours(green_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

                for pic, contour in enumerate(contours):
                    area = cv2.contourArea(contour)
                    if area > 300:
                        x, y, w, h = cv2.boundingRect(contour)
                        imageFrame = cv2.rectangle(imageFrame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                        cv2.putText(imageFrame, "Green Colour", (x, y), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 255, 0))

                contours, hierarchy = cv2.findContours(blue_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
                for pic, contour in enumerate(contours):
                    area = cv2.contourArea(contour)
                    if area > 300:
                        x, y, w, h = cv2.boundingRect(contour)
                        imageFrame = cv2.rectangle(imageFrame, (x, y), (x + w, y + h), (255, 0, 0), 2)
                        cv2.putText(imageFrame, "Blue Colour", (x, y), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (255, 0, 0))

                cv2.imshow("Multiple Color Detection in Real-Time", imageFrame)
                if cv2.waitKey(1) == ord('q'):
                    break
            webcam.release()
            cv2.destroyAllWindows()

        elif 'control mouse' in query:
            mouse_interaction()
        elif 'create a note' in query:
            create_note()
        elif 'read note' in query:
            read_note()
        elif 'set a reminder' in query:
            speak("What should I remind you about?")
            reminder_text = takeCommand()
            speak("When should I remind you? Please specify the time.")
            reminder_time = input("Enter reminder time: ")
            set_reminder(reminder_text, reminder_time)
        elif 'open my web' in query:
            speak("Here you go to your web\\n")
            os.startfile("https://bassetblack.github.io/basset-s-web/")

        elif 'none' in query:
            print("")
        else:
            qa_data = load_qa_data()
            handle_unrecognized_query(query,qa_data)