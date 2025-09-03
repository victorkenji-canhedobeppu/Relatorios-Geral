from tkinter import messagebox
import tkinter as tk


class LoginWindow(tk.Toplevel):
    """
    Janela modal para entrada de credenciais de login.
    """

    def __init__(self, master, firebase_controller):
        import tkinter as tk
        from tkinter import messagebox, font

        super().__init__(master)
        self.firebase_controller = firebase_controller
        self.login_successful = False

        # Configurações da Janela
        self.title("Firebase Login")
        self.geometry("400x320")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Centraliza a janela na tela
        self.update_idletasks()
        largura_janela = self.winfo_width()
        altura_janela = self.winfo_height()
        largura_tela = self.winfo_screenwidth()
        altura_tela = self.winfo_screenheight()
        x = (largura_tela // 2) - (largura_janela // 2)
        y = (altura_tela // 2) - (altura_janela // 2)
        self.geometry(f"+{x}+{y}")

        # Tenta criar a fonte Segoe UI ou usa Arial como fallback
        try:
            segoe_ui = font.Font(family="Segoe UI", size=10)
            segoe_ui_bold = font.Font(family="Segoe UI", size=10, weight="bold")
            segoe_ui_title = font.Font(family="Segoe UI", size=16, weight="bold")
        except:
            segoe_ui = font.Font(family="Arial", size=10)
            segoe_ui_bold = font.Font(family="Arial", size=10, weight="bold")
            segoe_ui_title = font.Font(family="Arial", size=16, weight="bold")

        # Cria e posiciona os widgets usando place() para coordenadas fixas
        # "Login" Label
        # O posicionamento x=0 e anchor='center' centraliza o texto na largura da janela
        lbl_title = tk.Label(self, text="Login", font=segoe_ui_title)
        lbl_title.place(relx=0.5, y=20, anchor=tk.CENTER)

        # "Email Address" Label e Entry
        lbl_email = tk.Label(self, text="Email Address", font=segoe_ui)
        lbl_email.place(x=40, y=80)
        self.email_entry = tk.Entry(self, font=segoe_ui)
        self.email_entry.place(x=40, y=105, width=320, height=25)

        # "Password" Label e Entry
        lbl_password = tk.Label(self, text="Password", font=segoe_ui)
        lbl_password.place(x=40, y=150)
        self.password_entry = tk.Entry(self, show="•", font=segoe_ui)
        self.password_entry.place(x=40, y=175, width=320, height=25)

        # "Login" Button
        self.login_button = tk.Button(
            self,
            text="Login",
            command=self._attempt_login,
            font=segoe_ui_bold,
            bg="DodgerBlue",
            fg="white",
            bd=0,
            relief="flat",
            activebackground="#1E90FF",
        )
        self.login_button.place(x=40, y=225, width=320, height=40)
        self.login_button.bind("<Enter>", lambda e: e.widget.config(bg="RoyalBlue"))
        self.login_button.bind("<Leave>", lambda e: e.widget.config(bg="DodgerBlue"))

        self.bind("<Return>", lambda event: self._attempt_login())

        self.grab_set()
        self.focus_set()
        self.wait_window()

    def _attempt_login(self):
        email = self.email_entry.get()
        password = self.password_entry.get()
        print(email)

        if not email or not password:
            messagebox.showwarning(
                "Campos Vazios", "Por favor, insira o email e a senha."
            )
            return

        sucesso, mensagem = self.firebase_controller.autenticar_usuario(email, password)

        if sucesso:
            self.login_successful = True
            self.destroy()
        else:
            messagebox.showerror("Falha no Login", mensagem)

    def _on_close(self):
        self.login_successful = False
        self.destroy()
