from docx import Document


def encontrar_todos_content_controls(caminho_do_arquivo):
    """
    Tenta encontrar todos os ContentControls do documento de forma mais robusta.
    """
    try:
        doc = Document(caminho_do_arquivo)

        # Acessa o corpo principal do documento (o `body` do XML).
        document_element = doc.element.body

        # Usa `findall` com uma busca mais abrangente no namespace do XML.
        # Ele procura por qualquer tag <sdt> em qualquer nível do documento.
        sdt_elements = document_element.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt"
        )

        if not sdt_elements:
            print("Nenhum ContentControl foi encontrado. 😥")
            return

        print("🔎 ContentControls encontrados:")
        for sdt_element in sdt_elements:
            # Pega as propriedades do ContentControl.
            sdt_props = sdt_element.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr"
            )

            # Extrai o título e a tag.
            title_elem = sdt_props.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}title"
            )
            tag_elem = sdt_props.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag"
            )

            titulo = (
                title_elem.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                )
                if title_elem is not None
                else "N/A"
            )
            tag = (
                tag_elem.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                )
                if tag_elem is not None
                else "N/A"
            )

            # Acessa o conteúdo do ContentControl, que pode ser o texto.
            sdt_content = sdt_element.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent"
            )
            texto_conteudo = ""
            if sdt_content is not None:
                runs = sdt_content.findall(
                    ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r"
                )
                texto_conteudo = "".join(
                    [
                        r.findtext(
                            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
                        )
                        for r in runs
                        if r.findtext(
                            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
                        )
                        is not None
                    ]
                )

            print("---")
            print(f"Título: {titulo}")
            print(f"Tag: {tag}")
            print(f"Conteúdo: '{texto_conteudo.strip()}'")

    except Exception as e:
        print(f"❌ Ocorreu um erro: {e}")


# Substitua 'seu_documento.docx' pelo caminho do seu arquivo
encontrar_todos_content_controls(
    r"C:\Users\Geologia\PythonProjects\Relatorios\RelatoriosGeral\src\templates\ANTT\RioSP.docx"
)
