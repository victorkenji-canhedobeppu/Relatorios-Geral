import os


class FileManager:
    @staticmethod
    def get_templates_from_folders():
        """Retorna os nomes das pastas dentro de /src/templates como opções de template, ignorando arquivos temporários ou ocultos."""
        templates_path = "src/templates"
        if os.path.isdir(templates_path):
            return [
                d
                for d in os.listdir(templates_path)
                if os.path.isdir(os.path.join(templates_path, d))
                and not d.startswith(".")
            ]
        return ["Nenhum Template Encontrado"]

    @staticmethod
    def get_formats_for_template(template_name):
        """Retorna os nomes dos arquivos (formatos) para um template específico, ignorando arquivos temporários ou ocultos."""
        formats_path = os.path.join("src/templates", template_name)
        if os.path.isdir(formats_path):
            files = [
                f
                for f in os.listdir(formats_path)
                if os.path.isfile(os.path.join(formats_path, f))
                and not f.startswith(".")
            ]
            # Retorna o nome completo do arquivo, em vez da extensão
            return files
        return ["Nenhum Formato Encontrado"]
