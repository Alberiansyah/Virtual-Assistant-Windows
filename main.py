import tkinter as tk
from threading import Thread
import speech_recognition as sr
import os
import subprocess
import time
import pyautogui

class SpeechRecognitionApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Speech Recognition App")
        self.master.geometry("400x300")

        self.result_text = tk.StringVar()
        self.result_label = tk.Label(self.master, textvariable=self.result_text, font=("Helvetica", 16))
        self.result_label.pack()

        # Menempatkan label di tengah dengan menggunakan anchor
        self.result_label.place(relx=0.5, rely=0.5, anchor="center")

        # Variabel untuk menyimpan hasil pengenalan suara
        self.speech_result = "Silahkan berbicara..."

        # Flag untuk memberitahu thread kapan harus berhenti
        self.is_running = True

        # Memanggil metode untuk memulai proses pengenalan suara
        self.start_recognition()

        # Menambahkan binding untuk menutup aplikasi dengan Ctrl+S
        self.master.bind("<Control-s>", self.close_app)

    def start_recognition(self):
        # Membuat thread agar GUI tidak terblokir
        recognition_thread = Thread(target=self.recognize_speech)
        recognition_thread.start()

    def recognize_speech(self):
        recognizer = sr.Recognizer()

        with sr.Microphone() as source:
            while self.is_running:
                audio = recognizer.listen(source)

                try:
                    text = recognizer.recognize_google(audio, language="id-ID")
                    self.speech_result = f"Anda mengatakan: \n{text.lower()}"
                    self.process_command(text.lower())
                except sr.UnknownValueError:
                    self.speech_result = "Maaf, suara tidak dapat dikenali."
                except sr.RequestError as e:
                    self.speech_result = f"Terjadi kesalahan pada request API: {e}"

    def update_result(self):
        self.result_text.set(self.speech_result)

        # Melanjutkan pemantauan secara live
        self.master.after(100, self.update_result)

    def close_app(self, event):
        # Fungsi untuk menutup aplikasi
        self.is_running = False  # Memberi tahu thread untuk berhenti
        self.master.destroy()

    def type_text_from_speech(self):
        while True:
            # Metode untuk mengetikkan teks dari ucapan ke Notepad
            recognizer = sr.Recognizer()

            with sr.Microphone() as source:
                print("Mulai berbicara...")
                recognizer.adjust_for_ambient_noise(source)
                audio = recognizer.listen(source)

            try:
                # Menggunakan Google Web Speech API untuk mengubah ucapan menjadi teks
                text = recognizer.recognize_google(audio, language="id-ID")
                print(f"Anda mengatakan: {text.lower()}")

                # Mensimulasikan input keyboard dengan pyautogui
                pyautogui.typewrite(text.lower())
                # pyautogui.press("enter")  # Tambahkan ini untuk membuat baris baru di Notepad

                # Hentikan loop jika mengatakan "berhenti"
                if "berhenti mengetik" in text.lower():
                    self.is_typing = False  # Set flag menjadi False
                    break
                # elif "space" in text.lower():
                #     pyautogui.press("space")
                # elif "koma" in text.lower():
                #     pyautogui.press(",")
                # elif "titik" in text.lower():
                #     pyautogui.press(".")


            except sr.UnknownValueError:
                print("Maaf, suara tidak dapat dikenali.")
            except sr.RequestError as e:
                print(f"Terjadi kesalahan pada request API: {e}")

    def process_command(self, command):
        if "buka control panel" in command:
            os.system("control")
        elif "buka explorer" in command:
            os.system("explorer")
        elif "buka task manager" in command:
            os.system("taskmgr")
        elif "buka command prompt" in command:
            os.system("cmd")
        elif "buka powershell" in command:
            os.system("powershell")
        elif "buka notepad" in command:
            subprocess.Popen(["notepad.exe"])
            # Setelah Notepad terbuka, panggil metode untuk mengetikkan teks
            self.type_text_from_speech()
        elif "buka browser" in command:
            os.system("start https://www.google.com")
        elif "buka chrome" in command:
            os.system("start chrome")
        elif "buka kalkulator" in command:
            os.system("calc")
        elif "buka word" in command:
            word_path = os.path.join(os.environ["ProgramFiles"], "Microsoft Office", "root", "Office16", "WINWORD.EXE")
            # Mengeksekusi Word
            subprocess.Popen([word_path])
            self.type_text_from_speech()
        elif "buka excel" in command:
            word_path = os.path.join(os.environ["ProgramFiles"], "Microsoft Office", "root", "Office16", "excel.EXE")
            # Mengeksekusi Word
            subprocess.Popen([word_path])
        elif "buka powerpoint" in command:
            word_path = os.path.join(os.environ["ProgramFiles"], "Microsoft Office", "root", "Office16", "powerpnt.EXE")
            # Mengeksekusi Word
            subprocess.Popen([word_path])
        elif "buka kalender" in command:
            os.system("start outlookcal:")
        elif "buka settings" in command:
            os.system("start ms-settings:")
        elif "buka paint" in command:
            os.system("start mspaint")
        elif "buka this pc" in command:
            os.system("explorer %SystemDrive%")
        elif "buka recycle bin" in command:
            os.system("explorer %SystemDrive%\$Recycle.Bin")
        elif "buka device manager" in command:
            os.system("devmgmt.msc")
        elif "buka services" in command:
            os.system("services.msc")
        elif "buka task scheduler" in command:
            os.system("control schedtasks")
        elif "buka disk management" in command:
            os.system("diskmgmt.msc")
        elif "tutup notepad" in command:
            self.close_application("notepad.exe")
        elif "tutup word" in command:
            self.close_application("WINWORD.EXE")
        elif "tutup excel" in command:
            self.close_application("excel.EXE")
        elif "tutup powerpoint" in command:
            self.close_application("powerpnt.EXE")
        elif "tulis" in command:
            self.speech_result = "Menulis..."
            self.update_result()
            self.type_text(command.split("tulis", 1)[1].strip())
        elif "tutup aplikasi" in command:
            self.speech_result = "Menutup aplikasi..."
            self.close_app(None)

    def close_application(self, process_name):
        try:
            os.system(f"taskkill /f /im {process_name}")
            self.speech_result = f"{process_name.capitalize()} berhasil ditutup."
        except Exception as e:
            self.speech_result = f"Gagal menutup {process_name.capitalize()}: {e}"

if __name__ == "__main__":
    root = tk.Tk()
    app = SpeechRecognitionApp(root)
    app.update_result()  # Memanggil metode update_result agar proses pemantauan dimulai
    root.mainloop()
