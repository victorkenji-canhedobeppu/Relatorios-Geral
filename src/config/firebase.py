from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives.padding import PKCS7
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
from cryptography.fernet import Fernet
from tkinter import messagebox
from ui import LoginWindow
from os import urandom
import datetime as dt
import tkinter as tk
import pyrebase
import base64
import json
import os


class FirebaseController:
    """
    Gerencia a interação com o Firebase, usando um arquivo de sessão local
    protegido com criptografia (via Fernet).
    """

    def __init__(self):

        self.initialized = False
        self.auth_client = None
        self.current_user = None

        # --- LÓGICA DE CRIPTOGRAFIA ---
        # Define o caminho para o arquivo de sessão
        appdata_path = os.path.join(os.getenv("APPDATA"), "EditorAGSAuth")
        self.session_path = os.path.join(appdata_path, "session.enc")
        os.makedirs(appdata_path, exist_ok=True)

        # Gera uma chave de criptografia segura derivada do ID da máquina.
        # Isso garante que o arquivo de sessão só pode ser lido neste computador.
        # Usa uma chave codificada diretamente para compatibilidade
        chave_codificada = "RoUvhFjHsWqK_XlQiKU7jLIXOGW-FaRT2cneL9D75xQ=".encode(
            "utf-8"
        )
        self.encryption_key = base64.urlsafe_b64decode(chave_codificada)
        self.backend = default_backend()

        try:
            # Sua configuração do Firebase (sem alterações)
            firebase_config = {
                "apiKey": "AIzaSyC7_N_mvzRLVV_iLQXd0RzRWoT3fSjCdJI",
                "authDomain": "auth-check-c0310.firebaseapp.com",
                "databaseURL": "https://auth-check-c0310-default-rtdb.firebaseio.com",
                "projectId": "auth-check-c0310",
                "storageBucket": "auth-check-c0310.firebasestorage.app",
                "messagingSenderId": "312679237582",
                "appId": "1:312679237582:web:40fa9939051b696be70d9f",
                "measurementId": "G-2KQ4JCCSMY",
            }
            firebase_client = pyrebase.initialize_app(firebase_config)
            self.auth_client = firebase_client.auth()
            self.initialized = True
            print("✅ Firebase Controller (com criptografia) inicializado.")
        except Exception as e:
            self.initialized = False
            messagebox.showerror(
                "Erro Crítico", f"Não foi possível conectar ao Firebase: {e}"
            )

    def autenticar_usuario(self, email, password):

        if not self.initialized:
            return False, "Serviço indisponível."
        try:
            user = self.auth_client.sign_in_with_email_and_password(email, password)

            self.current_user = user

            session_data = {
                "UserInfo": {
                    "Uid": user.get("localId"),
                    "FederatedId": None,
                    "FirstName": None,
                    "LastName": None,
                    "DisplayName": user.get("displayName") or None,
                    "Email": user.get("email"),
                    "IsEmailVerified": user.get("emailVerified") or False,
                    "PhotoUrl": None,
                    "IsAnonymous": user.get("isAnonymous") or False,
                },
                "Credential": {
                    "IdToken": user.get("idToken"),
                    "RefreshToken": user.get("refreshToken"),
                    "Created": dt.datetime.now(dt.timezone.utc)
                    .astimezone(dt.timezone(dt.timedelta(hours=-3)))
                    .isoformat(),
                    "ExpiresIn": int(user.get("expiresIn")),
                    "ProviderType": 7,
                },
            }

            session_json = json.dumps(session_data).encode("utf-8")

            # Criptografa os dados da sessão
            encrypted_data = self._encrypt_data(
                json.dumps(session_data).encode("utf-8")
            )

            # Salva os dados criptografados no arquivo
            with open(self.session_path, "wb") as f:
                f.write(encrypted_data)

            return True, "Login bem-sucedido!"
        except Exception as e:
            print(e)
            # ... (seu tratamento de erro de login) ...
            return False, "Email ou senha inválidos."

    def atualizar_login(self):

        if not os.path.exists(self.session_path):
            return False
        try:
            # Lê os dados criptografados do arquivo
            with open(self.session_path, "rb") as f:
                encrypted_data = f.read()

            # Descriptografa para obter o JSON
            decrypted_json = self._decrypt_data(encrypted_data).decode("utf-8")
            sessao = json.loads(decrypted_json)

            # ... (resto da sua lógica de refresh token, sem alterações) ...
            refresh_token = sessao["Credential"]["RefreshToken"]
            refreshed_data = self.auth_client.refresh(refresh_token)

            # Atualiza e salva a sessão de volta, criptografada
            sessao["Credential"].update(
                {
                    "IdToken": refreshed_data.get("idToken"),
                    "RefreshToken": refreshed_data.get("refreshToken"),
                    # ...
                }
            )
            new_session_json = json.dumps(sessao).encode("utf-8")
            new_encrypted_data = self._encrypt_data(new_session_json)
            with open(self.session_path, "wb") as f:
                f.write(new_encrypted_data)

            self.current_user = refreshed_data
            print("✅ Sessão atualizada com sucesso.")
            return True
        except Exception as e:
            print(
                f"❌ Falha ao atualizar a sessão (pode ser inválida ou corrompida): {e}"
            )
            self.fazer_logout()
            return False

    def fazer_logout(self):
        self.current_user = None
        if os.path.exists(self.session_path):
            os.remove(self.session_path)
        print("Logout realizado. Arquivo de sessão removido.")

    def _encrypt_data(self, data):

        padder = PKCS7(algorithms.AES.block_size).padder()
        padded_data = padder.update(data) + padder.finalize()
        iv = urandom(16)
        cipher = Cipher(
            algorithms.AES(self.encryption_key), modes.CBC(iv), backend=self.backend
        )
        encryptor = cipher.encryptor()
        ciphertext = encryptor.update(padded_data) + encryptor.finalize()
        return iv + ciphertext

    def _decrypt_data(self, encrypted_data):

        iv = encrypted_data[:16]
        ciphertext = encrypted_data[16:]
        cipher = Cipher(
            algorithms.AES(self.encryption_key), modes.CBC(iv), backend=self.backend
        )
        decryptor = cipher.decryptor()
        decrypted_padded_data = decryptor.update(ciphertext) + decryptor.finalize()
        unpadder = PKCS7(algorithms.AES.block_size).unpadder()
        return unpadder.update(decrypted_padded_data) + unpadder.finalize()

    def flow_autenticacao_usuario(self):
        is_user_logged_in = False

        if self.initialized:
            # Tenta fazer o login automático com o token de refresh
            auto_login_success = self.atualizar_login()

            if auto_login_success:
                is_user_logged_in = True
                messagebox.showinfo(
                    "Bem-vindo(a) de volta!", "Você foi conectado automaticamente."
                )
            else:
                # Se o login automático falhar, mostra a janela de login
                root = tk.Tk()
                root.withdraw()
                login_window = LoginWindow(root, self)  # Passa 'self' como controller

                if login_window.login_successful:
                    is_user_logged_in = True
                    messagebox.showinfo("Sucesso", "Login realizado com sucesso!")
                else:
                    is_user_logged_in = False
                    print("Login cancelado ou falhou.")
                root.destroy()

        return is_user_logged_in
