import win32com.client as win32
import os


def _apply_heading_numbering(word_app, doc):
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


def update_doc_with_headings_and_toc(caminho_do_arquivo):
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
        doc = word_app.Documents.Open(os.path.abspath(caminho_do_arquivo))

        if not _apply_heading_numbering(word_app, doc):
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
        if word_app:
            word_app.Quit()
