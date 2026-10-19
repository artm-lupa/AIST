print("Я жив!")

import customtkinter as ctk

app = ctk.CTk()
app.title("Тест")
app.geometry("400x200")

ctk.CTkLabel(app, text="Привет, AIST!",
             font=("Arial", 20)).pack(pady=40)

app.mainloop()