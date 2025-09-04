from docx import Document


def editar_content_control_em_tabela(caminho_do_arquivo, tag_procurada, novo_texto):
    """
    Função para editar o conteúdo de um ContentControl específico,
    com lógica para não quebrar tabelas.
    """
    try:
        doc = Document(caminho_do_arquivo)
        document_element = doc.element.body
        sdt_elements = document_element.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt"
        )

        encontrado = False

        for sdt_element in sdt_elements:
            sdt_props = sdt_element.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr"
            )
            tag_elem = sdt_props.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag"
            )

            if (
                tag_elem is not None
                and tag_elem.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                )
                == tag_procurada
            ):
                sdt_content = sdt_element.find(
                    ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent"
                )

                if sdt_content is not None:
                    # Passo 1: Limpar o conteúdo existente de forma segura
                    # Isso remove os parágrafos e "runs" antigos, mas mantém a estrutura da célula
                    for child in list(sdt_content):
                        sdt_content.remove(child)

                    # Passo 2: Adicionar um novo parágrafo com o texto desejado
                    # Cria um novo elemento de parágrafo <w:p>
                    p_element = sdt_content.makeelement(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
                    )
                    sdt_content.append(p_element)

                    # Adiciona um novo "run" <w:r> dentro do parágrafo
                    run_element = p_element.makeelement(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r"
                    )
                    p_element.append(run_element)

                    # Adiciona o elemento de texto <w:t> dentro do "run"
                    text_element = run_element.makeelement(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
                    )
                    run_element.append(text_element)

                    # Define o texto
                    text_element.text = novo_texto

                    print(
                        f"✅ Conteúdo do ContentControl com a tag '{tag_procurada}' alterado com sucesso!"
                    )
                    encontrado = True
                    break

        if not encontrado:
            print(
                f"❌ Nenhum ContentControl com a tag '{tag_procurada}' foi encontrado."
            )
            return

        novo_nome_arquivo = f"{caminho_do_arquivo.replace('.docx', '_editado.docx')}"
        doc.save(novo_nome_arquivo)
        print(f"📁 Documento salvo como: {novo_nome_arquivo}")

    except Exception as e:
        print(f"❌ Ocorreu um erro: {e}")


# Exemplo de uso:
# Altere 'seu_documento.docx' para o seu arquivo
# Altere 'MinhaTag' para a tag do ContentControl que você quer modificar
# Altere 'Texto novo inserido' para o texto que você quer colocar
editar_content_control_em_tabela(
    r"C:\Users\Geologia\PythonProjects\Relatorios\RelatoriosGeral\src\templates\ANTT\RioSP.docx",
    "Data_Revisao_Um",
    "1",
)
