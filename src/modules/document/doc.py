import re
import os
from docx import Document
import win32com.client as win32


class DocumentController:
    def __init__(self):
        self.word_app = None

    def _set_win32_instance(self):
        try:
            word_app = win32.Dispatch("Word.Application")
            word_app.Visible = False
            self.word_app = word_app
        except Exception as e:
            print("Erro", e)

    def read_content_controls(self, caminho_do_arquivo):
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

    def update_doc_tags(self, caminho_do_arquivo, dados_do_formulario):
        """
        Atualiza o conteúdo dos ContentControls de um documento .docx
        usando pywin32, com base nos dados coletados da interface.

        Args:
            caminho_do_arquivo (str): Caminho completo para o arquivo .docx.
            dados_do_formulario (dict): Dicionário com os dados do formulário,
                                        incluindo campos gerais e revisões.
        """
        try:
            # 1. Inicia o Word
            word_app = win32.Dispatch("Word.Application")
            word_app.Visible = False

            # 2. Abre o documento
            doc = self.word_app.Documents.Open(os.path.abspath(caminho_do_arquivo))

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

    def _apply_heading_numbering(self, doc):
        """
        Constrói e aplica um esquema de numeração hierárquico aos estilos de Título.
        Esta função é a chave para garantir que a numeração funcione sempre.
        """
        try:
            # Pega na galeria de listas de múltiplos níveis

            # Adiciona um novo template de lista ao documento. Isto "reseta" a formatação.
            list_template = doc.ListTemplates.Add(True)

            text_position_pt = 1.5 * 28.3464567

            lvl1 = list_template.ListLevels(1)
            lvl1.NumberFormat = "%1"  # Formato "NúmeroDoNível1.NúmeroDoNível2"
            lvl1.TrailingCharacter = 2
            lvl1.NumberStyle = 0
            lvl1.LinkedStyle = "Título 1"
            lvl1.NumberPosition = 0
            lvl1.TextPosition = text_position_pt
            lvl1.TabPosition = text_position_pt

            # Define o Nível 2 para estar ligado ao estilo "Título 2"
            lvl2 = list_template.ListLevels(2)
            lvl2.NumberFormat = "%1.%2"  # Formato "NúmeroDoNível1.NúmeroDoNível2"
            lvl2.TrailingCharacter = 2
            lvl2.NumberStyle = 0
            lvl2.LinkedStyle = "Título 2"
            lvl2.NumberPosition = 0
            lvl2.TextPosition = text_position_pt
            lvl2.TabPosition = text_position_pt

            # Define o Nível 3 para estar ligado ao estilo "Título 3"
            lvl3 = list_template.ListLevels(3)
            lvl3.NumberFormat = "%1.%2.%3"  # Formato "N1.N2.N3"
            lvl3.TrailingCharacter = 2
            lvl3.NumberStyle = 0
            lvl3.LinkedStyle = "Título 3"
            lvl3.NumberPosition = 0
            lvl3.TextPosition = text_position_pt
            lvl3.TabPosition = text_position_pt

            lvl4 = list_template.ListLevels(4)
            lvl4.NumberFormat = "%1.%2.%3.%4"  # Formato "N1.N2.N3"
            lvl4.TrailingCharacter = 2
            lvl4.NumberStyle = 0
            lvl4.LinkedStyle = "Título 4"
            lvl4.NumberPosition = 0
            lvl4.TextPosition = text_position_pt
            lvl4.TabPosition = text_position_pt

            print("Numeração dos estilos de título foi configurada com sucesso.")
            return True
        except Exception as e:
            print(f"Não foi possível configurar a numeração dos títulos. Erro: {e}")
            return False

    def update_doc_with_headings_and_toc(self, caminho_do_arquivo):
        """
        Insere novos títulos e subtítulos na última página do documento
        e atualiza o sumário.

        Args:
            caminho_do_arquivo (str): O caminho para o documento .docx.
        """
        word_app = None
        try:
            # Inicia o Word (invisível)
            word_app = win32.gencache.EnsureDispatch("Word.Application")
            word_app.Visible = False
            doc = self.word_app.Documents.Open(os.path.abspath(caminho_do_arquivo))

            if not self._apply_heading_numbering(word_app, doc):
                return

            # Encontra o final do documento e insere uma quebra de página
            end_of_doc_range = doc.Content
            end_of_doc_range.Collapse(win32.constants.wdCollapseEnd)
            end_of_doc_range.InsertBreak(win32.constants.wdPageBreak)

            # Define a nova área de inserção (após a quebra de página)
            insertion_range = doc.Range(end_of_doc_range.End, end_of_doc_range.End)

            print("✏️ Inserindo conteúdo na última página do documento.")

            # --- Exemplo de conteúdo a ser inserido ---
            conteudo = [
                {"text": " VAZÕES DE PROJETO", "style": "Título 1"},
                {"text": " SUBTÍTULO NOVO", "style": "Título 2"},
                {"text": " Subtítulo Mais Interno", "style": "Título 3"},
                {"text": "Este é um texto normal.", "style": "Normal"},
            ]

            # Insere o novo conteúdo no final do documento
            for i, item in enumerate(conteudo):
                # A partir do segundo item, insere uma quebra de parágrafo antes
                if i > 0:
                    insertion_range.InsertAfter("\r")
                    insertion_range.Collapse(win32.constants.wdCollapseEnd)

                insertion_range.Text = item["text"]
                insertion_range.Style = item["style"]

            # Atualiza o sumário
            print("🔄 Atualizando o sumário...")
            for toc in doc.TablesOfContents:
                toc.Update()
                print("✅ Sumário atualizado com sucesso!")
                break
            else:
                print("⚠️ Aviso: Nenhum sumário foi encontrado para ser atualizado.")

            # Salva as alterações
            doc.Save()
            doc.Close(False)
            print("📁 Documento salvo e fechado.")

        except Exception as e:
            print(f"❌ Ocorreu um erro: {e}")
