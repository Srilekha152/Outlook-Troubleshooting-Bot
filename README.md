# Outlook-Troubleshooting-Bot
The **Outlook Troubleshooting Bot** is a simple Python project that automates basic Outlook support tasks like checking if Outlook is running, simulating cache clearing, and sending test emails using voice commands.
# 🛠️ Outlook Troubleshooting Bot (Python)

An intelligent desktop bot that automates the process of detecting and resolving common Microsoft Outlook issues using voice commands. Built using Python, this tool simulates Level 1 IT support tasks such as checking Outlook status, clearing cache (simulated), and sending test emails via SMTP. The bot integrates speech recognition to allow hands-free troubleshooting — perfect for learning automation in service desk scenarios.

---

## ✅ Features

- 🎙️ Voice-command control using Google Speech Recognition API  
- 📥 Detects if Outlook is running and attempts to launch if not  
- 🧹 Simulates cache clearing for Outlook issues  
- 📧 Sends test emails via Gmail SMTP  
- 🗂️ Logs all actions and errors to a local file for review  
- 🧠 Easily extendable for more service desk functions

---

## 🔧 Technologies Used

- **Python 3.x**  
- `speech_recognition` for capturing voice commands  
- `smtplib` for sending emails  
- `subprocess` and `os` for system command execution  
- `datetime` for logging  
- `pyaudio` for microphone input

---

## 🗂️ Folder Structure

