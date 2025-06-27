import os
import subprocess
import smtplib
import speech_recognition as sr
from datetime import datetime
from utils import log_action

def listen_command():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)
    try:
        command = recognizer.recognize_google(audio)
        return command.lower()
    except sr.UnknownValueError:
        return "Could not understand the command."

def check_outlook():
    output = subprocess.getoutput("tasklist")
    if "OUTLOOK.EXE" in output:
        print("Outlook is running.")
        log_action("Outlook is running.")
    else:
        print("Outlook is not running. Attempting to open Outlook...")
        log_action("Outlook not running. Launching it.")
        subprocess.Popen("start outlook", shell=True)

def clear_cache():
    # Simulated logic for clearing Outlook cache
    print("Simulating clearing cache...")
    log_action("Cleared simulated Outlook cache.")

def send_test_email():
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login('your_email@gmail.com', 'your_app_password')  # Use App Password!
        server.sendmail('your_email@gmail.com', 'recipient@example.com', 'Test email from Outlook Bot')
        server.quit()
        print("Test email sent.")
        log_action("Sent test email.")
    except Exception as e:
        log_action(f"Email failed: {str(e)}")

def main():
    print("Say a command (check outlook, clear cache, send email):")
    command = listen_command()
    if "check outlook" in command:
        check_outlook()
    elif "clear cache" in command:
        clear_cache()
    elif "send email" in command:
        send_test_email()
    else:
        print("Unknown command.")
        log_action("Unknown command: " + command)

if __name__ == "__main__":
    main()
