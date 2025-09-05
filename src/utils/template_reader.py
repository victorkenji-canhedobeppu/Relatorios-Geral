import os
import re
from docx import Document


def read_content_controls(caminho_do_arquivo):
    """
    Lê os ContentControls do documento .docx e os categoriza em
    campos gerais e campos de revisão.
    """
    campos_encontrados = {}

    # Mapeamento para converter palavras em números
    num_map = {
        "Zero": "0",
        "Um": "1",
        "Dois": "2",
        "Tres": "3",
        "Quatro": "4",
        "Cinco": "5",
        "Seis": "6",
        "Sete": "7",
    }

    # Mapeamento de tags para nomes de exibição amigáveis
    tag_map = {
        "Cod_Interno": "Código Interno",
        "Cod_ANTT": "Código ANTT",
        "Emitente": "Emitente",
        "Data_Emissao_Inicial": "Data Emissão Inicial",
        "Rodovia": "Rodovia",
        "Projetista": "Projetista",
        "Trecho": "Trecho",
        "Objeto": "Objeto",
    }

    # Mapeamento de tipo de campo para nome de exibição amigável
    display_name_map = {
        "Revisao": "Revisão",
        "Versao": "Versão",
        "Data_Revisao": "Data Revisão",
        "Descricao": "Descrição",
    }

    print(f"🕵️‍♀️ Tentando ler o arquivo: {caminho_do_arquivo}")

    try:
        doc = Document(caminho_do_arquivo)
        document_element = doc.element.body

        sdt_elements = document_element.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt"
        )

        for sdt_element in sdt_elements:
            sdt_props = sdt_element.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr"
            )
            sdt_content = sdt_element.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent"
            )

            tag = None

            tag_elem = sdt_props.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag"
            )
            if tag_elem is not None:
                tag = tag_elem.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                )

            if tag == "Identificador_Tipo":
                print("⚠️ Ignorando campo com a tag 'Identificador_Tipo'")
                continue

            texto_conteudo = ""
            if sdt_content is not None:
                # Extrai texto de todos os parágrafos dentro do sdtContent
                for paragraph in sdt_content.findall(
                    ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
                ):
                    for run in paragraph.findall(
                        ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r"
                    ):
                        text_part = run.findtext(
                            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
                        )
                        if text_part:
                            texto_conteudo += text_part
            texto_conteudo = texto_conteudo.strip()

            # Se o conteúdo estiver vazio, substitui por um hífen
            if not texto_conteudo:
                texto_conteudo = "-"

            # Armazena todos os campos em um dicionário temporário
            if tag:
                campos_encontrados[tag] = texto_conteudo

    except Exception as e:
        print(f"❌ Ocorreu um erro ao ler o template: {e}")
        return {}, {}

    # Agora, categoriza os campos após a leitura
    campos_gerais = {}
    campos_revisoes = {}

    for tag, content in campos_encontrados.items():
        match = re.search(r"(Revisao|Versao|Data_Revisao|Descricao)_(\w+)", tag)
        if match:
            field_type, rev_number_str = match.groups()
            display_type = display_name_map.get(field_type, field_type)
            rev_number_digit = num_map.get(rev_number_str, rev_number_str)

            if rev_number_digit not in campos_revisoes:
                campos_revisoes[rev_number_digit] = {}
            campos_revisoes[rev_number_digit][display_type] = content
        else:
            display_name = tag_map.get(tag, tag)
            campos_gerais[display_name] = content

    return campos_gerais, campos_revisoes
