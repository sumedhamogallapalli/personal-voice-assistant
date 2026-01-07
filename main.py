import win32com.client
import speech_recognition as sr
import webbrowser
import datetime
import winsound
import time
import Whats_app
import chat_app
import news_app
import weather_app
import pyautogui
import keyboard
from translate import Translator
from alarm import set_alarm
from word2number import w2n
import snake 
import pyqrcode
import turtle
import os
from youtube import play_specified_song,watch_specified_movie,play_specified_video
from joke import tell_joke,get_riddle,check_answer
from volume import set_system_volume, change_volume
import open_files
import random
import parsedatetime
import remainder
import system_info
from location import get_user_location
import schedule
from dateutil import parser
import re
import threading

news_apikey = "4d46f612033e484ea9f4ba943c1a8232"

speaker = win32com.client.Dispatch("SAPI.SpVoice")  # todo:takes command from user

def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        try:
            audio = r.listen(source, timeout=10)  
            query = r.recognize_google(audio, language="en-in")
            print(f"user: {query}")
            return query
        except sr.WaitTimeoutError:
            print("Voice Assistant: Timeout. No speech detected.")
            return "Timeout. No speech detected."
        except sr.UnknownValueError:
            print("Voice Assistant: Sorry, I could not understand that.")
            return "Sorry, I could not understand that."
        except sr.RequestError as e:
            print(f"Voice Assistant: Could not request results from Google Speech Recognition service; {e}")
            return "Could not request results from Google Speech Recognition service."

def schedule_video_play(video_name, hour, minute):
    current_time = time.localtime()
    alarm_time = time.struct_time((current_time.tm_year, current_time.tm_mon, current_time.tm_mday,
                                   int(hour), int(minute), 0, 0, 0, 0))

    if alarm_time > current_time:
        time_difference = time.mktime(alarm_time) - time.mktime(current_time)

        print(f"Video will play at {hour}:{minute}")

        # Define a function to play the video and start a new thread for it
        def play_video():
            print(f"Voice Assistant: Playing '{video_name}' on YouTube at {hour}:{minute}.")
            speaker.speak(f"Playing '{video_name}' on YouTube at {hour}:{minute}.")
            play_specified_video(video_name)

        # Create a new thread for playing the video
        threading.Timer(time_difference, play_video).start()
    else:
        print("Invalid time. Please provide a future time.")
    
    
def volume_up():
    change_volume("up")
    print("Voice Assistant: Volume increased.")
    speaker.speak("Volume increased.")

def volume_down():
    change_volume("down")
    print("Voice Assistant: Volume decreased.")
    speaker.speak("Volume decreased.")

def set_volume():
    print("Voice Assistant: Set the volume level between 0 and 100.")
    speaker.speak("Set the volume level between 0 and 100.")
    volume_level = take_command()
    try:
        volume_level = float(volume_level) / 100.0
        set_system_volume(volume_level)
        print(f"Voice Assistant: Volume set to {int(volume_level * 100)}%.")
        speaker.speak(f"Volume set to {int(volume_level * 100)}%.")
    except ValueError:
        print("Voice Assistant: Please provide a valid volume level.")
        speaker.speak("Please provide a valid volume level.")

def shutdown():
    print("Voice Assistant: Shutting down the computer.")
    speaker.speak("Shutting down the computer.")
    os.system("shutdown /s /t 1")

def restart():
    print("Voice Assistant: Restarting the computer.")
    speaker.speak("Restarting the computer.")
    os.system("shutdown /r /t 1")

def log_off():
    print("Voice Assistant: Logging off the computer.")
    speaker.speak("Logging off the computer.")
    os.system("shutdown -l")

def take_screenshot():
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    screenshot_filename = f"screenshot_{timestamp}.png"
    os.system(f"import -window root {screenshot_filename}")
    print(f"Voice Assistant: Screenshot saved as {screenshot_filename}.")
    speaker.speak(f"Screenshot saved as {screenshot_filename}.")

# todo: for playing sound
winsound.PlaySound("C:\\Users\\M.Sreedhar\\OneDrive\\Desktop\\Team Project\\harry.wav", winsound.SND_ASYNC)

# code = take_command()
print('Voice Assistant: Say the Secret Code to Start MAX VOICE ASSISTANT')
speaker.speak("say the secret code to start MAX VOICE ASSISTANT")
code = take_command()

if code.lower() == "hello max" or code.lower() == "hello maths":
    print("Voice Assistant:Hello Sir I am MAX AI \nVoice Assistant: How can I help you")
    speaker.speak("hello sir I am MAX AI")
    speaker.speak("How can I help you")
    
    while True:
        time.sleep(1)
        winsound.PlaySound("C:\\Users\\M.Sreedhar\\OneDrive\\Desktop\\Team Project\\harry.wav", winsound.SND_ASYNC)
        time.sleep(1)
        speaker.speak("listening your command")
        query = take_command()
        
        if "who invented you".lower() in query.lower():
            print("Voice Assistant: I was invented by Sumedha and her team members")
            speaker.speak("I was invented by Sumedha and her team members")

        elif "How are you".lower() in query.lower():
            print("Voice Assistant: I am great! What about you?")
            speaker.speak("I am great! What about you?")

        elif "What is your name".lower() in query.lower():
            print("Voice Assistant: My name is MAX")
            speaker.speak("My name is MAX") 
                  
        elif "open web browser".lower() in query.lower():
            print("Voice Assistant: Which web browser do you want to open?")
            speaker.speak("What do you want to open")
            open_app = take_command()
            site = f"https://www.{open_app}.com"
            print(f"Voice Assistant: opening {open_app} Sir")
            speaker.speak(f"opening {open_app} Sir")
            webbrowser.get('C:/Program Files/Google/Chrome/Application/chrome.exe %s').open(site)
            

        elif "open file".lower() in query.lower():
            print("Voice Assistant: Which file do you want to open?")
            speaker.speak("Which file do you want to open?")
            open_file = take_command().lower()

            if open_file == "notepad":
                open_files.open_notepad()
            elif open_file == "calculator":
                open_files.open_calculator()
            elif open_file == "file explorer":
                open_files.open_file_explorer()
            elif open_file == "command prompt":
                open_files.open_command_prompt()
            elif open_file == "paint":
                open_files.open_paint()
            elif open_file == "wordpad":
                openfiles.open_wordpad()
            elif open_file == "task manager":
                open_files.open_task_manager()
            elif open_file == "settings":
                open_files.open_settings()
            elif open_file == "control panel":
                open_files.open_control_panel()
            elif open_file == "chrome":
                open_files.open_chrome()
            elif open_file == "firefox":
                open_files.open_firefox()
            elif open_file == "edge":
                open_files.open_edge()
            elif open_file == "notepad++":
                open_files.open_notepadplusplus()
            elif open_file == "visual studio code":
                open_files.open_visual_studio_code()
            elif open_file == "excel":
                open_files.open_excel()
            else:
                print("Voice Assistant: Invalid File")
                speaker.speak("Invalid File")

        elif "take screenshot".lower() in query.lower():
            take_screenshot()

        elif "get system info".lower() in query.lower():
            system_info.get_system_info()

        elif "get network info".lower() in query.lower():
            system_info.get_network_info()

        elif "get disk usage".lower() in query.lower():
            system_info.get_disk_usage()

        elif "get cpu usage".lower() in query.lower():
            system_info.get_cpu_usage()

        elif "get memory usage".lower() in query.lower():
            system_info.get_memory_usage()

        elif "get battery status".lower() in query.lower():
            system_info.get_battery_status()

        elif "get location".lower() in query.lower():
            get_user_location()

        elif "check internet".lower() in query.lower():
            response = os.system("ping google.com")
            if response == 0:
                print("Voice Assistant: You are connected to the internet.")
                speaker.speak("You are connected to the internet.")
            else:
                print("Voice Assistant: You are not connected to the internet.")
                speaker.speak("You are not connected to the internet.")

        elif "tell me a riddle".lower() in query.lower():
            riddle_data = get_riddle()
            question = riddle_data["question"]
            correct_answer = riddle_data["answer"]

            print(f"Voice Assistant: Riddle: {question}. What's your answer?")
            speaker.speak(f"Riddle: {question}. What's your answer?")

            # Get the user's answer
            user_answer = take_command()

            # Check the answer
            if check_answer(user_answer, correct_answer):
                print("Voice Assistant: Correct! Well done.")
                speaker.speak("Correct! Well done.")
            else:
                print(f"Voice Assistant: That's incorrect. The correct answer is {correct_answer}.")
                speaker.speak(f"That's incorrect. The correct answer is {correct_answer}.")


        elif "play game".lower() in query.lower():
            print("Voice Assistant: Sure! Let's play a game. I will think of a number, and you try to guess it.")
            speaker.speak("Sure! Let's play a game. I will think of a number, and you try to guess it.")

            secret_number = random.randint(1, 20)

            attempts = 0
            max_attempts = 5

            while attempts < max_attempts:
                print("Voice Assistant: Guess the number (between 1 and 20):")
                speaker.speak("Guess the number (between 1 and 20):")
                user_guess = take_command()

                try:
                    user_guess = w2n.word_to_num(user_guess)
                except ValueError:
                    print("Voice Assistant: Please enter a valid number.")
                    speaker.speak("Please enter a valid number.")
                    continue

                attempts += 1

                if user_guess == secret_number:
                    print(f"Voice Assistant: Congratulations! You guessed the correct number {secret_number} in {attempts} attempts.")
                    speaker.speak(f"Congratulations! You guessed the correct number {secret_number} in {attempts} attempts.")
                    break
                elif user_guess < secret_number:
                    print("Voice Assistant: Try a higher number.")
                    speaker.speak("Try a higher number.")
                else:
                    print("Voice Assistant: Try a lower number.")
                    speaker.speak("Try a lower number.")

            if attempts == max_attempts:
                print(f"Voice Assistant: Sorry, you've reached the maximum attempts. The correct number was {secret_number}.")
                speaker.speak(f"Sorry, you've reached the maximum attempts. The correct number was {secret_number}.")
                    
        elif "what can you do".lower() in query.lower():
            print("Voice Assistant: I can tell you like friend and I can solve any doubt and also, I can generate any code or email ")
            speaker.speak("I can talk you like friend and I can solve any doubt and also I can generate any code or email ")
            
        elif "what's the time".lower() in query:
            hours = datetime.datetime.now().strftime("%H")
            minutes = datetime.datetime.now().strftime("%M")
            sec = datetime.datetime.now().strftime("%S")
            print(f"Voice Assistant: Sir the time is {hours} hours {minutes} minutes and {sec} seconds")
            speaker.speak(f"Sir the time is {hours} hours {minutes} minutes and {sec} seconds")
            
        elif "search the web".lower() in query.lower():
            print("Voice Assistant: What would you like to search for?")
            speaker.speak("What would you like to search for?")
            search_query = take_command()
            webbrowser.open(f"https://www.google.com/search?q={search_query}")

        elif "set remainder".lower() in query.lower():
            print("Voice Assistant: What's the reminder message?")
            speaker.speak("What's the reminder message?")
            reminder_message = take_command()
            print("Voice Assistant: When should I remind you? (e.g., in 5 minutes)")
            speaker.speak("When should I remind you? (e.g., in 5 minutes)")
            reminder_time = take_command()

            # Schedule the reminder using the reminder module
            remainder.schedule_reminder(reminder_message, reminder_time, speaker)


        elif "using your intelligence".lower() in query.lower():
            speaker.speak("Yes Sir what can i do for you")
            sub_write=take_command()
            speaker.speak(sub_write)
            result = chat_app.get_wikipedia_summary(sub_write)
            chat_app.text_to_speech(result)
            print("Voice Assistant: Successfully written for you Sir")
            speaker.speak("successfully written for you Sir")
        
        elif "chat reset".lower() in query.lower():
            chatstr = ""
            print("Voice Assistant: chat reseted.")
            speaker.speak("chat reseted.")
            
        elif "translator".lower() in query.lower():

            # Get the sentence to translate
            print("Voice Assistant: What do you want to translate?")
            speaker.speak("What do you want to translate?")
            sentence = take_command()

            if sentence == "sorry i can't understand say it again":
                speaker.speak("Sorry, I can't understand. Please try again.")
                break

            # Get the target language for translation
            print("Voice Assistant: In which language do you want to translate?")
            speaker.speak("In which language do you want to translate?")
            target_lang = take_command()

            if target_lang == "sorry i can't understand say it again":
                speaker.speak("Sorry, I can't understand. Please try again.")
                break

            # Perform the translation
            translator = Translator(to_lang=target_lang)
            translation = translator.translate(sentence)

            # Speak the translated sentence
            speaker.speak(f"Translating '{sentence}' into {target_lang}")
            print(f"Voice Assistant: The translated sentence is: {translation}")
            speaker.speak(f"The translated sentence is: {translation}")
            print("Voice Assistant: translated Successfully")
            speaker.speak("translated Successfully")

        elif "file operations".lower() in query.lower():
            print("Voice Assistant: What file operation do you want to perform? (create, delete, append, read, write)")
            speaker.speak("What file operation do you want to perform? (create, delete, append, read, write)")
            file_operation = take_command().lower()

            if file_operation == "create":
                print("Voice Assistant: What should be the name of the new file?")
                speaker.speak("What should be the name of the new file?")
                file_name = take_command()

                print("Voice Assistant: Enter the file extension (e.g., txt, csv):")
                speaker.speak("Enter the file extension (e.g., txt, csv):")
                file_extension = take_command()

                file_name_with_extension = f"{file_name}.{file_extension}"

                with open(file_name_with_extension, 'w') as file:
                    file.write("")  # Create an empty file
                print(f"Voice Assistant: {file_name_with_extension} created successfully.")
                speaker.speak(f"{file_name_with_extension} created successfully.")

            elif file_operation == "delete":
                print("Voice Assistant: What is the name of the file you want to delete?")
                speaker.speak("What is the name of the file you want to delete?")
                file_to_delete = take_command()

                print("Voice Assistant: Enter the file extension (e.g., txt, csv):")
                speaker.speak("Enter the file extension (e.g., txt, csv):")
                file_extension = take_command()

                file_to_delete_with_extension = f"{file_to_delete}.{file_extension}"

                try:
                    os.remove(file_to_delete_with_extension)
                    print(f"Voice Assistant: {file_to_delete_with_extension} deleted successfully.")
                    speaker.speak(f"{file_to_delete_with_extension} deleted successfully.")
                except FileNotFoundError:
                    print(f"Voice Assistant: File {file_to_delete_with_extension} not found.")
                    speaker.speak(f"File {file_to_delete_with_extension} not found.")

            elif file_operation == "append":
                print("Voice Assistant: What is the name of the file you want to append to?")
                speaker.speak("What is the name of the file you want to append to?")
                file_to_append = take_command()

                print("Voice Assistant: Enter the file extension (e.g., txt, csv):")
                speaker.speak("Enter the file extension (e.g., txt, csv):")
                file_extension = take_command()

                file_to_append_with_extension = f"{file_to_append}.{file_extension}"

                print("Voice Assistant: What content do you want to append?")
                speaker.speak("What content do you want to append?")
                content_to_append = take_command()

                with open(file_to_append_with_extension, 'a') as file:
                    file.write(f"\n{content_to_append}")  # Append content to the file
                print(f"Voice Assistant: Content appended to {file_to_append_with_extension} successfully.")
                speaker.speak(f"Content appended to {file_to_append_with_extension} successfully.")

            elif file_operation == "write" or file_operation == "right":
                print("Voice Assistant: What is the name of the file you want to write to?")
                speaker.speak("What is the name of the file you want to write to?")
                file_to_write = take_command()

                print("Voice Assistant: Enter the file extension (e.g., txt, csv):")
                speaker.speak("Enter the file extension (e.g., txt, csv):")
                file_extension = take_command()

                file_to_write_with_extension = f"{file_to_write}.{file_extension}"

                print("Voice Assistant: What content do you want to write?")
                speaker.speak("What content do you want to write?")
                content_to_write = take_command()

                with open(file_to_write_with_extension, 'w') as file:
                    file.write(content_to_write)  # Write content to the file
                print(f"Voice Assistant: Content written to {file_to_write_with_extension} successfully.")
                speaker.speak(f"Content written to {file_to_write_with_extension} successfully.")

            elif file_operation == "read":
                print("Voice Assistant: What is the name of the file you want to read?")
                speaker.speak("What is the name of the file you want to read?")
                file_to_read = take_command()

                print("Voice Assistant: Enter the file extension (e.g., txt, csv):")
                speaker.speak("Enter the file extension (e.g., txt, csv):")
                file_extension = take_command()

                file_to_read_with_extension = f"{file_to_read}.{file_extension}"

                try:
                    with open(file_to_read_with_extension, 'r') as file:
                        file_content = file.read()
                        print(f"Voice Assistant: Content of {file_to_read_with_extension}:\n{file_content}")
                        speaker.speak(f"Content of {file_to_read_with_extension}:\n{file_content}")
                except FileNotFoundError:
                    print(f"Voice Assistant: File {file_to_read_with_extension} not found.")
                    speaker.speak(f"File {file_to_read_with_extension} not found.")

            
        elif "latest news about".lower() in query.lower():
            print("Voice Assistant: news about	")
            news_app.news(query)

        elif "what's the weather at".lower() in query.lower():
            print("Voice Assistant:",end='')
            # speaker.speak(query)
            weather_app.weather(query)
            
        elif "whatsapp message".lower() in query.lower():
            j=1
            while(j>=0):
                print("Voice Assistant: Whom do you want send sir")
                speaker.speak("whom do you want send sir")
                name=take_command()
                if name=="sorry i can't understand say it again":
                    speaker.speak("i can't understand")
                else:
                    k=1
                    while(k>=0):
                        print("Voice Assistant: What message do you want to send")
                        speaker.speak("what message do you want to send")
                        message=take_command()
                        if message=="sorry i can't understand say it again":
                            speaker.speak("i can't understand")
                        else:
                            print(f"Voice Assistant: Sending the {message} message to {name}")
                            speaker.speak(f"sending the {message} message to {name}")
                            time.sleep(1)
                            pyautogui.press('super')
                            time.sleep(1)
                            pyautogui.click(x=622, y=165)
                            time.sleep(1)
                            keyboard.write("whatsapp")
                            pyautogui.sleep(2)
                            pyautogui.press('enter')
                            time.sleep(1)
                            Whats_app.Whatsapp_msg(name,message)
                            print(f"Voice Assistant: message sent successfully to {name}")
                            speaker.speak(f"message sent successfully to {name}")
                        break
                break
            
        elif "whatsapp call".lower() in query.lower():
            j=1
            while(j>=0):
                print("Voice Assistant: Whom do you want to call sir")
                speaker.speak("whom do you want to call sir")
                name=take_command()
                if name=="sorry i can't understand say it again":
                    speaker.speak("i can't understand")
                else:
                    print(f"Voice Assistant: Calling to {name}")
                    speaker.speak(f"calling to {name}")
                    time.sleep(1)
                    pyautogui.press('super')
                    time.sleep(1)
                    pyautogui.click(x=622, y=165)
                    time.sleep(1)
                    keyboard.write("whatsapp")
                    pyautogui.sleep(2)
                    pyautogui.press('enter')
                    Whats_app.Whatsapp_call(name)
                    print("Voice Assistant: Call ended")
                    speaker.speak("call ended")
                    break

        elif "whatsapp video call".lower() in query.lower():
            j=1
            while(j>=0):
                print("Voice Assistant: Whom do you want to video call sir")
                speaker.speak("whom do you want to video call sir")
                name=take_command()
                if name=="sorry i can't understand say it again":
                    speaker.speak("i can't understand")
                else:
                    print(f"Voice Assistant: Video calling to {name}")
                    speaker.speak(f"video calling to {name}")
                    time.sleep(1)
                    pyautogui.press('super')
                    time.sleep(1)
                    pyautogui.click(x=622, y=165)
                    time.sleep(1)
                    keyboard.write("whatsapp")
                    pyautogui.sleep(2)
                    pyautogui.press('enter')
                    Whats_app.Whatsapp_video_call(name)
                    print("Voice Assistant: Video call ended")
                    speaker.speak("video call ended")
                    break

        elif "play song".lower() in query.lower():
            print("Voice Assistant: What song would you like to play?")
            speaker.speak("What song would you like to play?")
            song_name = take_command()
            play_specified_song(song_name)


        elif "tell me a joke".lower() in query.lower():
            tell_joke()

        elif "volume up" in query:
            volume_up()
            
        elif "volume down" in query:
            volume_down()
            
        elif "set volume" in query:
            set_volume()

        elif "shutdown" in query:
            shutdown()
            
        elif "restart" in query:
            restart()
            
        elif "log off" in query:
            log_off()

        elif "check battery" in query:
            check_battery()
                
        elif "play snake game".lower() in query.lower():
            print("Voice Assistant: Starting Snake game.")
            speaker.speak("Starting Snake game.")
            turtle.setup(420, 420, 370, 0)
            turtle.hideturtle()
            turtle.tracer(False)
            turtle.listen()
            turtle.onkey(lambda: snake.change(10, 0), 'Right')
            turtle.onkey(lambda: snake.change(-10, 0), 'Left')
            turtle.onkey(lambda: snake.change(0, 10), 'Up')
            turtle.onkey(lambda: snake.change(0, -10), 'Down')
            snake.move()
            turtle.done()
            print("Voice Assistant: Snake game ended.")
            speaker.speak("Snake game ended.")

        elif "generate qr code".lower() in query.lower():
            print("Voice Assistant: Generating QR code.")
            speaker.speak("Generating QR code.")
            print("Voice Assistant: Say the URL or text for your QR code: ")
            speaker.speak("Say the URL or text for your QR code: ")
            custom_url = take_command()
            url = pyqrcode.create(custom_url)
            print("Voice Assistant: Enter the desired file name (without extension): ")
            speaker.speak("Enter the desired file name (without extension): ")
            file_name = take_command()
            url.svg(f"{file_name}.svg", scale=8)
            print(f"Voice Assistant: QR code for {custom_url} has been generated and saved as {file_name}.svg")
            speaker.speak(f"QR code for {custom_url} has been generated and saved as {file_name}.svg")

        elif "watch movie".lower() in query.lower():
            print("Voice Assistant: What movie would you like to watch?")
            speaker.speak("What movie would you like to watch?")
            movie_name = take_command()
            watch_specified_movie(movie_name)

        elif "play video at specified time".lower() in query.lower():
            print("Voice Assistant: Which video would you like to watch?")
            speaker.speak("Which video would you like to watch?")
            video_name = take_command()
            print("Voice Assistant: Say which hours in(24-hour format)do you want to paly video?")
            speaker.speak("say which hours in(24-hour format)do you want to paly video?")
            hour=take_command()
            print("Voice Assistant: Say which minute do you want to paly video?")
            speaker.speak("say which minute do you want to paly video?")
            minute=take_command()
            sec=00
            print(f"Voice Assistant: Playing video at {hour} hours and {minute} minutes")
            speaker.speak(f"Playing video at {hour} hours and {minute} minutes")

            try:
                alarm_hour = w2n.word_to_num(hour)
                alarm_minute = w2n.word_to_num(minute)

                schedule_video_play(video_name,alarm_hour, alarm_minute)

            except ValueError:
                print("Invalid input. Please enter valid numbers for the hour and minute.")
            
        elif "alarm".lower() in query.lower():
            print("Voice Assistant: Say which hours in(24-hour format)do you want to set alarm?")
            speaker.speak("say which hours in(24-hour format)do you want to set alarm?")
            hour=take_command()
            print("Voice Assistant: Say which minute do you want to set alarm?")
            speaker.speak("say which minute do you want to set alarm?")
            minute=take_command()
            sec=00
            print(f"Voice Assistant: Setting alarm at {hour} hours and {minute} minutes")
            speaker.speak(f"Setting alarm at {hour} hours and {minute} minutes")

            try:
                alarm_hour = w2n.word_to_num(hour)
                alarm_minute = w2n.word_to_num(minute)

                set_alarm(alarm_hour, alarm_minute)

            except ValueError:
                print("Invalid input. Please enter valid numbers for the hour and minute.")

        elif "goodbye".lower() in query.lower():
            print("Voice Assistant: goodbye sir have a great day...!")
            speaker.speak("goodbye have a great day...!")
            break

        else:
            print("Voice Assistant: Sorry I can't understand")
            speaker.speak("Sorry I can't understand")


else:
    speaker.speak("wrong code")
    
