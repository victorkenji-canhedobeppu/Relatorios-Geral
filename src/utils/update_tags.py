import win32com.client as win32
import os


def update_doc_tags(caminho_do_arquivo, dados_do_formulario):
    """
    Atualiza o conteúdo dos ContentControls de um documento .docx
    usando pywin32, com base nos dados coletados da interface.

    Args:
        caminho_do_arquivo (str): Caminho completo para o arquivo .docx.
        dados_do_formulario (dict): Dicionário com os dados do formulário,
                                    incluindo campos gerais e revisões.
    """
    word_app = None
    try:
        # 1. Inicia o Word
        word_app = win32.Dispatch("Word.Application")
        word_app.Visible = False

        # 2. Abre o documento
        doc = word_app.Documents.Open(os.path.abspath(caminho_do_arquivo))

        print("Iniciando a atualização das tags do documento...")

        # Mapeamento do nome do campo na interface para o nome da tag no Word
        tag_geral_map = {
            "Código Interno": "Cod_Interno",
            "Código ANTT": "Cod_ANTT",
            "Emitente": "Emitente",
            "Data Emissão Inicial": "Data_Emissao_Inicial",
            "Rodovia": "Rodovia",
            "Projetista": "Projetista",
            "Trecho": "Trecho",
            "Objeto": "Objeto",
        }

        revisao_map = {
            "Revisão": "Revisao",
            "Versão": "Versao",
            "Data Revisão": "Data_Revisao",
            "Descrição": "Descricao",
        }

        # 3. Itera por todos os ContentControls do documento
        for cc in doc.ContentControls:
            tag = cc.Tag

            # Verifica se a tag é um campo geral
            if tag in tag_geral_map.values():
                display_name = next(
                    (k for k, v in tag_geral_map.items() if v == tag), None
                )
                if display_name and display_name in dados_do_formulario:
                    novo_valor = dados_do_formulario[display_name]
                    cc.Range.Text = novo_valor
                    print(f"  ✅ Tag geral '{tag}' atualizada para '{novo_valor}'.")

            # Verifica se a tag faz parte de uma revisão
            else:
                for rev_id, rev_data in dados_do_formulario.items():
                    # Ignora as chaves que não são números de revisão
                    if not isinstance(rev_data, dict):
                        continue

                    for form_key, tag_prefix in revisao_map.items():
                        # Constrói o nome da tag completa, ex: 'Revisao_1', 'Descricao_0'
                        expected_tag = f"{tag_prefix}_{rev_id}"
                        if tag == expected_tag:
                            novo_valor = rev_data.get(form_key, "")
                            cc.Range.Text = novo_valor
                            print(
                                f"  ✅ Tag de revisão '{tag}' atualizada para '{novo_valor}'."
                            )
                            break
                    if tag == expected_tag:
                        break

        # 4. Salva e fecha o documento
        doc.Save()
        doc.Close(False)

        print("Concluído! Documento salvo com sucesso.")

    except Exception as e:
        print(f"❌ Ocorreu um erro: {e}")
    finally:
        if word_app:
            word_app.Quit()
