import os
import customtkinter as ctk
import datetime
from utils import (
    FileManager,
    AnttFieldsManager,
    update_doc_tags,
    update_doc_with_headings_and_toc,
)
from modules import DocumentController


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configuração da Janela ---
        self.title("Criar Novo Relatório")
        self.geometry("1100x00")
        self.minsize(1100, 800)

        # --- Configuração do Layout Principal (2 colunas) ---
        self.grid_columnconfigure(0, weight=0)  # Coluna da esquerda, fixa
        self.grid_columnconfigure(1, weight=1)  # Coluna da direita, se expande
        self.grid_rowconfigure(
            0, weight=1
        )  # Permite que a linha principal se expanda verticalmente

        # --- Painel Esquerdo: Opções do Relatório ---
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=10)
        self.sidebar_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(7, weight=1)

        # Título do Painel Esquerdo
        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame, text="Opções", font=ctk.CTkFont(size=20, weight="bold")
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Dropdown de Templates
        self.template_label = ctk.CTkLabel(
            self.sidebar_frame, text="Selecione o Template:"
        )
        self.template_label.grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")

        templates = FileManager.get_templates_from_folders()
        if "ANTT" in templates:
            templates.remove("ANTT")
            templates.insert(0, "ANTT")

        self.template_optionmenu = ctk.CTkOptionMenu(
            self.sidebar_frame, values=templates, command=self._update_formats
        )
        self.template_optionmenu.grid(
            row=2, column=0, padx=20, pady=(0, 10), sticky="ew"
        )

        # Dropdown de Formatos
        self.format_label = ctk.CTkLabel(
            self.sidebar_frame, text="Selecione o Formato:"
        )
        self.format_label.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="w")
        self.format_optionmenu = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=["Nenhum Formato Encontrado"],
            command=self._update_config_frame,
        )
        self.format_optionmenu.grid(row=4, column=0, padx=20, pady=(0, 10), sticky="ew")

        # Dropdown de Disciplinas
        self.subject_label = ctk.CTkLabel(
            self.sidebar_frame, text="Selecione a Disciplina:"
        )
        self.subject_label.grid(row=5, column=0, padx=20, pady=(10, 0), sticky="w")
        self.subject_optionmenu = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=[
                "Hidrologia",
                "Geometria",
            ],
        )
        self.subject_optionmenu.grid(
            row=6, column=0, padx=20, pady=(10, 20), sticky="ew"
        )

        # --- Painel Direito: Configurações ---
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

        self.config_frame = None
        self.chapter_entries = []
        self.antt_manager = None
        self.doc_controller = DocumentController()

        self.caminho_do_arquivo = None

        self._update_formats()

    def _update_formats(self, choice=None):
        selected_template = self.template_optionmenu.get()
        formats = FileManager.get_formats_for_template(selected_template)

        self.format_optionmenu.configure(values=formats)
        if formats:
            self.format_optionmenu.set(formats[0])

        self._update_config_frame()

    def _update_config_frame(self, choice=None):
        if self.config_frame:
            self.config_frame.destroy()

        self.config_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.config_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.config_frame.grid_columnconfigure(0, weight=1)
        self.config_frame.grid_rowconfigure(1, weight=1)

        selected_template = self.template_optionmenu.get()
        if selected_template == "ANTT":
            self._display_antt_fields()
        else:
            self._display_chapter_fields()

    def _display_chapter_fields(self):
        chapter_label = ctk.CTkLabel(
            self.config_frame,
            text="Configuração dos Capítulos",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        chapter_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        scrollable_frame = ctk.CTkScrollableFrame(self.config_frame, corner_radius=0)
        scrollable_frame.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="nsew")
        scrollable_frame.grid_columnconfigure(0, weight=1)

        self.chapter_entries = []
        scrollable_frame.grid_columnconfigure(1, weight=1)

        for i in range(1, 11):
            scrollable_frame.grid_rowconfigure(i, weight=1)

            label = ctk.CTkLabel(scrollable_frame, text=f"Nome do Capítulo {i}:")
            label.grid(row=i, column=0, padx=(10, 5), pady=5, sticky="w")

            entry = ctk.CTkEntry(
                scrollable_frame,
                placeholder_text=f"Ex: Introdução ao Capítulo {i}",
            )
            entry.grid(row=i, column=1, padx=(5, 10), pady=5, sticky="ew")
            self.chapter_entries.append(entry)

        generate_button = ctk.CTkButton(
            self.config_frame,
            text="Gerar Relatório",
            command=self.generate_report,
        )
        generate_button.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="e")

    def _display_antt_fields(self):
        antt_label = ctk.CTkLabel(
            self.config_frame,
            text="Campos ANTT",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        antt_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        selected_template = self.template_optionmenu.get()
        selected_format = self.format_optionmenu.get()
        self.caminho_do_arquivo = os.path.join(
            "src", "templates", selected_template, selected_format
        )
        self.antt_manager = AnttFieldsManager(
            self.config_frame, self.caminho_do_arquivo
        )

        generate_button = ctk.CTkButton(
            self.config_frame,
            text="Gerar Relatório",
            command=self.generate_report,
        )
        generate_button.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="e")

    def generate_report(self):
        print("Botão 'Gerar Relatório' clicado!")
        template = self.template_optionmenu.get()
        format = self.format_optionmenu.get()
        subject = self.subject_optionmenu.get()

        print(f"Template selecionado: {template}")
        print(f"Formato selecionado: {format}")
        print(f"Disciplina selecionada: {subject}")

        if template == "ANTT":
            if self.antt_manager:
                data = self.antt_manager.get_field_values()
                print("\nDados dos campos ANTT:")
                print(data)
                try:

                    # Chama a função de atualização com o novo caminho e os dados
                    self.doc_controller.update_doc_tags(self.caminho_do_arquivo, data)
                    self.doc_controller.update_doc_with_headings_and_toc(
                        self.caminho_do_arquivo
                    )

                    print("Documento atualizado com sucesso!")

                except Exception as e:
                    print(f"Erro ao salvar e atualizar o documento: {e}")
        else:
            print("\nNomes dos Capítulos:")
            for i, entry in enumerate(self.chapter_entries):
                chapter_name = entry.get()
                print(
                    f"Capítulo {i+1}: {chapter_name if chapter_name else 'Não preenchido'}"
                )
