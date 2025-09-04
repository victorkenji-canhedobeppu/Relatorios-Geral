from config import FirebaseController
from ui import App

if __name__ == "__main__":
    # firebase_controller = FirebaseController()
    # is_user_logged_in = firebase_controller.flow_autenticacao_usuario()
    is_user_logged_in = True
    # 5. Inicie sua aplicação principal aqui se o login foi bem-sucedido
    if is_user_logged_in:
        print("\n🚀 O usuário está logado. Iniciando a aplicação principal...")
        # Inicia a aplicação
        app = App()
        app.mainloop()

    else:
        print("\n🛑 O usuário não está logado. Encerrando a aplicação.")
