import customtkinter as ctk
import datetime
import os
from .template_reader import read_content_controls


class AnttFieldsManager:
    """
    Gerencia a criação e coleta de dados dos campos específicos do template ANTT.
    Os campos são lidos dinamicamente de um arquivo .docx com ContentControls.
    """

    def __init__(self, parent_frame, caminho_do_arquivo):
        self.parent_frame = parent_frame
        self.general_entries = {}
        self.revision_entries = {}

        # Carrega os campos a partir do template usando a função externa
        self.campos_gerais, self.campos_revisoes = read_content_controls(
            caminho_do_arquivo
        )

        self._create_widgets()

    def _create_widgets(self):
        """Cria todos os widgets na frame do gerenciador."""
        scrollable_frame = ctk.CTkScrollableFrame(self.parent_frame, corner_radius=0)
        scrollable_frame.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="nsew")
        scrollable_frame.grid_columnconfigure(0, weight=1)

        # --- Seção de Campos Gerais ---
        general_frame = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        general_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        general_frame.grid_columnconfigure(0, weight=0)
        general_frame.grid_columnconfigure(1, weight=1)

        general_label = ctk.CTkLabel(
            general_frame,
            text="Campos Gerais",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        general_label.grid(
            row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w"
        )

        row = 1
        for label_text, default_value in self.campos_gerais.items():
            label = ctk.CTkLabel(general_frame, text=label_text + ":")
            label.grid(row=row, column=0, padx=(10, 5), pady=5, sticky="w")

            entry = ctk.CTkEntry(general_frame)
            entry.insert(0, default_value)
            entry.grid(row=row, column=1, padx=(5, 10), pady=5, sticky="ew")

            self.general_entries[label_text] = entry
            row += 1

        # --- Separador ---
        separator = ctk.CTkFrame(
            scrollable_frame, height=2, fg_color="gray", corner_radius=0
        )
        separator.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

        # --- Seção de Revisões ---
        revisions_frame = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        revisions_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)

        revisions_frame.grid_columnconfigure(0, weight=0)
        revisions_frame.grid_columnconfigure(1, weight=1)
        revisions_frame.grid_columnconfigure(2, weight=1)
        revisions_frame.grid_columnconfigure(3, weight=1)
        revisions_frame.grid_columnconfigure(4, weight=2)

        revisions_label = ctk.CTkLabel(
            revisions_frame, text="Revisões", font=ctk.CTkFont(size=16, weight="bold")
        )
        revisions_label.grid(
            row=0, column=0, columnspan=5, padx=10, pady=(10, 5), sticky="w"
        )

        headers = ["", "Revisão", "Versão", "Data Revisão", "Descrição"]
        for col, header_text in enumerate(headers):
            header_label = ctk.CTkLabel(
                revisions_frame, text=header_text, font=ctk.CTkFont(weight="bold")
            )
            header_label.grid(row=1, column=col, padx=5, sticky="ew")

        row = 2
        for i in sorted(
            self.campos_revisoes.keys(), key=lambda x: int(x), reverse=True
        ):
            rev_data = self.campos_revisoes[i]

            rev_label = ctk.CTkLabel(revisions_frame, text=f"Revisão {i}:")
            rev_label.grid(row=row, column=0, padx=(10, 5), pady=5, sticky="w")

            revisao_entry = ctk.CTkEntry(revisions_frame)
            revisao_entry.insert(0, rev_data.get("Revisão", "-"))
            revisao_entry.grid(row=row, column=1, padx=5, pady=5, sticky="ew")

            versao_entry = ctk.CTkEntry(revisions_frame)
            versao_entry.insert(0, rev_data.get("Versão", "-"))
            versao_entry.grid(row=row, column=2, padx=5, pady=5, sticky="ew")

            data_entry = ctk.CTkEntry(revisions_frame)
            data_entry.insert(0, rev_data.get("Data Revisão", "-"))
            data_entry.grid(row=row, column=3, padx=5, pady=5, sticky="ew")

            descricao_entry = ctk.CTkEntry(revisions_frame)
            descricao_entry.insert(0, rev_data.get("Descrição", "-"))
            descricao_entry.grid(row=row, column=4, padx=5, pady=5, sticky="ew")

            self.revision_entries[i] = {
                "Revisão": revisao_entry,
                "Versão": versao_entry,
                "Data Revisão": data_entry,
                "Descrição": descricao_entry,
            }
            row += 1

    def get_field_values(self):
        """Coleta todos os valores dos campos e os retorna em um dicionário."""
        data = {}
        for label, entry in self.general_entries.items():
            data[label] = entry.get()

        for rev, entries in self.revision_entries.items():
            data[rev] = {key: entry.get() for key, entry in entries.items()}

        return data
