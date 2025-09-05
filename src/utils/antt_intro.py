import customtkinter as ctk
import datetime
import os
from .template_reader import read_content_controls


class AnttFieldsManager:
    """
    Gerencia a criação e coleta de dados dos campos específicos do template ANTT.
    Os campos são lidos dinamicamente de um arquivo .docx com ContentControls.
    """

    # Define uma ordem fixa para os campos gerais
    GENERAL_FIELD_ORDER = [
        "Código Interno",
        "Código ANTT",
        "Emitente",
        "Data Emissão Inicial",
        "Rodovia",
        "Projetista",
        "Trecho",
        "Objeto",
    ]

    def __init__(self, parent_frame, caminho_do_arquivo):
        self.parent_frame = parent_frame
        self.general_entries = {}
        self.revision_entries = {}
        self.general_content_frame = None
        self.revisions_content_frame = None

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
        general_header_frame = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        general_header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        general_header_frame.grid_columnconfigure(0, weight=1)

        general_label = ctk.CTkLabel(
            general_header_frame,
            text="Campos Gerais",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        general_label.grid(row=0, column=0, sticky="w")

        general_toggle_button = ctk.CTkButton(
            general_header_frame,
            text="-",
            width=25,
            command=lambda: self._toggle_section(
                self.general_content_frame, general_toggle_button
            ),
        )
        general_toggle_button.grid(row=0, column=1, sticky="e")

        self.general_content_frame = ctk.CTkFrame(
            scrollable_frame, fg_color="transparent"
        )
        self.general_content_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        self.general_content_frame.grid_columnconfigure(0, weight=0)
        self.general_content_frame.grid_columnconfigure(1, weight=1)

        row = 0
        # Itera sobre a ordem fixa para garantir que os campos sejam exibidos na ordem correta
        for label_text in self.GENERAL_FIELD_ORDER:
            if label_text in self.campos_gerais:
                default_value = self.campos_gerais[label_text]

                label = ctk.CTkLabel(self.general_content_frame, text=label_text + ":")
                label.grid(row=row, column=0, padx=(10, 5), pady=5, sticky="w")

                entry = ctk.CTkEntry(self.general_content_frame)
                entry.insert(0, default_value)
                entry.grid(row=row, column=1, padx=(5, 10), pady=5, sticky="ew")

                self.general_entries[label_text] = entry
                row += 1

        # --- Separador ---
        separator = ctk.CTkFrame(
            scrollable_frame, height=2, fg_color="gray", corner_radius=0
        )
        separator.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        # --- Seção de Revisões ---
        revisions_header_frame = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        revisions_header_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        revisions_header_frame.grid_columnconfigure(0, weight=1)

        revisions_label = ctk.CTkLabel(
            revisions_header_frame,
            text="Revisões",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        revisions_label.grid(row=0, column=0, sticky="w")

        revisions_toggle_button = ctk.CTkButton(
            revisions_header_frame,
            text="-",
            width=25,
            command=lambda: self._toggle_section(
                self.revisions_content_frame, revisions_toggle_button
            ),
        )
        revisions_toggle_button.grid(row=0, column=1, sticky="e")

        self.revisions_content_frame = ctk.CTkFrame(
            scrollable_frame, fg_color="transparent"
        )
        self.revisions_content_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)

        self.revisions_content_frame.grid_columnconfigure(0, weight=0)
        self.revisions_content_frame.grid_columnconfigure(1, weight=1)
        self.revisions_content_frame.grid_columnconfigure(2, weight=1)
        self.revisions_content_frame.grid_columnconfigure(3, weight=1)
        self.revisions_content_frame.grid_columnconfigure(4, weight=2)

        headers = ["", "Revisão", "Versão", "Data Revisão", "Descrição"]
        for col, header_text in enumerate(headers):
            header_label = ctk.CTkLabel(
                self.revisions_content_frame,
                text=header_text,
                font=ctk.CTkFont(weight="bold"),
            )
            header_label.grid(row=0, column=col, padx=5, sticky="ew")

        row = 1
        for i in sorted(
            self.campos_revisoes.keys(), key=lambda x: int(x), reverse=True
        ):
            rev_data = self.campos_revisoes[i]

            rev_label = ctk.CTkLabel(self.revisions_content_frame, text=f"Revisão {i}:")
            rev_label.grid(row=row, column=0, padx=(10, 5), pady=5, sticky="w")

            revisao_entry = ctk.CTkEntry(self.revisions_content_frame)
            revisao_entry.insert(0, rev_data.get("Revisão", "-"))
            revisao_entry.grid(row=row, column=1, padx=5, pady=5, sticky="ew")

            versao_entry = ctk.CTkEntry(self.revisions_content_frame)
            versao_entry.insert(0, rev_data.get("Versão", "-"))
            versao_entry.grid(row=row, column=2, padx=5, pady=5, sticky="ew")

            data_entry = ctk.CTkEntry(self.revisions_content_frame)
            data_entry.insert(0, rev_data.get("Data Revisão", "-"))
            data_entry.grid(row=row, column=3, padx=5, pady=5, sticky="ew")

            descricao_entry = ctk.CTkEntry(self.revisions_content_frame)
            descricao_entry.insert(0, rev_data.get("Descrição", "-"))
            descricao_entry.grid(row=row, column=4, padx=5, pady=5, sticky="ew")

            self.revision_entries[i] = {
                "Revisão": revisao_entry,
                "Versão": versao_entry,
                "Data Revisão": data_entry,
                "Descrição": descricao_entry,
            }
            row += 1

    def _toggle_section(self, section_frame, button):
        if section_frame.winfo_viewable():
            section_frame.grid_forget()
            button.configure(text="+")
        else:
            if section_frame == self.general_content_frame:
                section_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
            elif section_frame == self.revisions_content_frame:
                section_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
            button.configure(text="-")

    def get_field_values(self):
        """Coleta todos os valores dos campos e os retorna em um dicionário."""
        data = {}
        for label, entry in self.general_entries.items():
            data[label] = entry.get()

        for rev, entries in self.revision_entries.items():
            data[rev] = {key: entry.get() for key, entry in entries.items()}

        return data
