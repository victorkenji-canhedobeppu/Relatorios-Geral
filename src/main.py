from config import FirebaseController

if __name__ == "__main__":
    firebase_controller = FirebaseController()
    is_user_logged_in = firebase_controller.flow_autenticacao_usuario()

    # 5. Inicie sua aplicação principal aqui se o login foi bem-sucedido
    if is_user_logged_in:
        print("\n🚀 O usuário está logado. Iniciando a aplicação principal...")

    else:
        print("\n🛑 O usuário não está logado. Encerrando a aplicação.")
