import configparser
from pathlib import Path

class ConfigManager:
    """Classe para gerenciar leitura e escrita de configurações"""
    
    def __init__(self):
        self.CONFIG_PATH = Path(__file__).parent.parent / "config.ini"
        self.load_config()
    
    def load_config(self):
        """Carrega as configurações do arquivo ini"""
        config = configparser.ConfigParser()
        templates = {}
        settings = {"theme": "Neon"}  # Neon como padrão inicial
        if self.CONFIG_PATH.exists():
            try:
                config.read(self.CONFIG_PATH, encoding="utf-8-sig")
                for section in config.sections():
                    if section == "Settings":
                        settings = dict(config[section])
                        if "theme" not in settings or settings["theme"] not in ["Neon", "Claro"]:
                            settings["theme"] = "Neon"  # Forçar Neon se inválido
                    else:
                        templates[section] = dict(config[section])
            except configparser.MissingSectionHeaderError:
                print("Erro: config.ini inválido.")
        return {"Templates": templates, "Settings": settings}
    
    def save_config(self, templates: dict, settings: dict):
        """Salva as configurações no arquivo ini"""
        config = configparser.ConfigParser()
        config["Settings"] = settings
        for name, data in templates.items():
            config[name] = data
        with open(self.CONFIG_PATH, "w", encoding="utf-8") as f:
            config.write(f)
    
    def get_current_settings(self, app_state):
        """Obtém as configurações atuais da aplicação"""
        return {
            "pastasorigem": ";".join(app_state.listbox_origem.item(i).text() for i in range(app_state.listbox_origem.count())),
            "pastadestino": app_state.textbox_destino.text(),
            "moverarquivos": str(app_state.checkbox_mover.isChecked()),
            "excluirduplicatas": str(app_state.checkbox_excluir_duplicatas.isChecked()),
            "usarlixeira": str(app_state.checkbox_lixeira.isChecked()),
            "usarsubpastas": str(app_state.checkbox_subpastas.isChecked()),
            "usarhash": str(app_state.checkbox_hash.isChecked()),
            "abrirdestino": str(app_state.checkbox_abrir_destino.isChecked()),
            "encerrarprograma": str(app_state.checkbox_encerrar.isChecked()),
            "filtro": app_state.combobox_filtro.currentText(),
            "usarregex": str(app_state.checkbox_regex.isChecked())
        }