import sys
import os
import shutil
import hashlib
from pathlib import Path
import configparser
import re
from typing import Dict, Optional
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QListWidget, QPushButton, QTextEdit, QLineEdit, QCheckBox, 
                             QComboBox, QFileDialog, QMessageBox, QProgressBar, QDialog, QSizePolicy)
from PyQt6.QtCore import QThread
from PyQt6.QtGui import QFont
from send2trash import send2trash
import win32com.client
# Configuração inicial
def get_base_path():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent

CONFIG_PATH = get_base_path() / "config.ini"

FILTERS = [
    "*.*",
    "*.jpg;*.jpeg;*.png;*.gif;*.bmp",  # Imagens
    "*.mp4;*.avi;*.mkv;*.mov;*.wmv",  # Vídeos
    "*.exe;*.msi;*.bat;*.cmd",  # Executáveis
    "*.py;*.cs;*.java;*.js;*.cpp;*.html;*.css",  # Códigos
    "*.doc;*.docx;*.pdf;*.txt;*.xlsx;*.pptx",  # Documentos
    "*.zip;*.rar;*.7z;*.tar.gz",  # Arquivos Compactados
    "*.mp3;*.wav;*.flac;*.aac",  # Áudio
    r"\.jpe?g$",  # Regex: JPEG/JPG
    r"\.mp[34]$",  # Regex: MP3/MP4
    r"^doc.*\.pdf$",  # Regex: PDFs começando com "doc"
    r"\.(cs|py|java)$"  # Regex: C#, Python, Java
]

class FileOrganizerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Organizador de Arquivos v3.4")
        self.setGeometry(100, 100, 1280, 720)
        self.setFont(QFont("Consolas", 10))
        self.is_processing_selection = False
        self.undo_stack = []
        self.setup_ui()
        self.load_initial_config()

    def setup_ui(self):
        # Central Widget e Layout Principal (Vertical)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)  # Sem margens externas
        main_layout.setSpacing(0)

        # Layout Horizontal para Config e Visualização
        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        # Painel de Configurações (Esquerda)
        self.panel_config = QWidget()
        self.panel_config.setFixedWidth(340)
        config_layout = QVBoxLayout(self.panel_config)
        config_layout.setContentsMargins(10, 10, 10, 10)
        config_layout.setSpacing(8)

        self.label_origem = QLabel("Pastas de Origem:")
        self.listbox_origem = QListWidget()
        self.listbox_origem.setMinimumHeight(150)
        self.button_add_origem = QPushButton("Adicionar")
        self.button_remove_origem = QPushButton("Remover")
        self.button_clear_form = QPushButton("Limpar Tudo")

        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.addWidget(self.button_add_origem)
        button_layout.addWidget(self.button_remove_origem)
        button_layout.addWidget(self.button_clear_form)

        self.label_destino = QLabel("Pasta de Destino:")
        self.textbox_destino = QLineEdit()
        self.button_selecionar_destino = QPushButton("Selecionar")
        self.button_clear_destino = QPushButton("Limpar")

        button_layout2 = QHBoxLayout()
        button_layout2.setContentsMargins(0, 0, 0, 0)
        button_layout2.addWidget(self.button_selecionar_destino)
        button_layout2.addWidget(self.button_clear_destino)

        self.label_config = QLabel("Configurações:")
        self.checkbox_mover = QCheckBox("Mover arquivos")
        self.checkbox_excluir_duplicatas = QCheckBox("Excluir duplicatas")
        self.checkbox_excluir_duplicatas.setChecked(True)
        self.checkbox_lixeira = QCheckBox("Usar lixeira")
        self.checkbox_lixeira.setChecked(True)
        self.checkbox_subpastas = QCheckBox("Organizar em subpastas")
        self.checkbox_hash = QCheckBox("Comparar hash (MD5)")
        self.checkbox_abrir_destino = QCheckBox("Abrir pasta destino")
        self.checkbox_encerrar = QCheckBox("Encerrar após executar")
        self.checkbox_regex = QCheckBox("Usar Regex")

        checkbox_layout = QHBoxLayout()
        checkbox_layout.setContentsMargins(0, 0, 0, 0)
        checkbox_layout.addWidget(self.checkbox_mover)
        checkbox_layout.addWidget(self.checkbox_excluir_duplicatas)

        checkbox_layout2 = QHBoxLayout()
        checkbox_layout2.setContentsMargins(0, 0, 0, 0)
        checkbox_layout2.addWidget(self.checkbox_lixeira)
        checkbox_layout2.addWidget(self.checkbox_subpastas)

        checkbox_layout3 = QHBoxLayout()
        checkbox_layout3.setContentsMargins(0, 0, 0, 0)
        checkbox_layout3.addWidget(self.checkbox_hash)
        checkbox_layout3.addWidget(self.checkbox_abrir_destino)

        self.label_filtro = QLabel("Filtro de arquivos:")
        self.combobox_filtro = QComboBox()
        self.combobox_filtro.addItems(FILTERS)
        self.combobox_filtro.setEditable(True)
        self.combobox_filtro.setCurrentText("*.jpg;*.png")
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 0)
        filter_layout.addWidget(self.label_filtro)
        filter_layout.addWidget(self.combobox_filtro)

        self.label_tema = QLabel("Tema:")
        self.combobox_tema = QComboBox()
        self.combobox_tema.addItems(["Neon", "Claro"])  # Neon primeiro
        theme_layout = QHBoxLayout()
        theme_layout.setContentsMargins(0, 0, 0, 0)
        theme_layout.addWidget(self.label_tema)
        theme_layout.addWidget(self.combobox_tema)

        self.label_templates = QLabel("Templates:")
        self.combobox_templates = QComboBox()
        self.label_template_name = QLabel("Nome do template:")
        self.textbox_template_name = QLineEdit()
        template_name_layout = QHBoxLayout()
        template_name_layout.setContentsMargins(0, 0, 0, 0)
        template_name_layout.addWidget(self.label_template_name)
        template_name_layout.addWidget(self.textbox_template_name)

        self.button_salvar_template = QPushButton("Salvar")
        self.button_editar_template = QPushButton("Editar")
        self.button_excluir_template = QPushButton("Excluir")
        template_button_layout = QHBoxLayout()
        template_button_layout.setContentsMargins(0, 0, 0, 0)
        template_button_layout.addWidget(self.button_salvar_template)
        template_button_layout.addWidget(self.button_editar_template)
        template_button_layout.addWidget(self.button_excluir_template)

        config_layout.addWidget(self.label_origem)
        config_layout.addWidget(self.listbox_origem)
        config_layout.addLayout(button_layout)
        config_layout.addWidget(self.label_destino)
        config_layout.addWidget(self.textbox_destino)
        config_layout.addLayout(button_layout2)
        config_layout.addWidget(self.label_config)
        config_layout.addLayout(checkbox_layout)
        config_layout.addLayout(checkbox_layout2)
        config_layout.addLayout(checkbox_layout3)
        config_layout.addWidget(self.checkbox_encerrar)
        config_layout.addWidget(self.checkbox_regex)
        config_layout.addLayout(filter_layout)
        config_layout.addLayout(theme_layout)
        config_layout.addWidget(self.label_templates)
        config_layout.addWidget(self.combobox_templates)
        config_layout.addLayout(template_name_layout)
        config_layout.addLayout(template_button_layout)
        config_layout.addStretch()
        content_layout.addWidget(self.panel_config)

        # Painel de Visualização (Direita)
        self.panel_visualizacao = QWidget()
        visualizacao_layout = QVBoxLayout(self.panel_visualizacao)
        visualizacao_layout.setContentsMargins(0, 0, 0, 0)
        visualizacao_layout.setSpacing(8)

        self.label_log = QLabel("Log de Operações:")
        self.logbox = QTextEdit()
        self.logbox.setReadOnly(True)
        self.logbox.setMinimumHeight(300)
        self.label_preview = QLabel("Pré-visualização:")
        preview_button_layout = QHBoxLayout()
        preview_button_layout.setContentsMargins(0, 0, 0, 0)
        preview_button_layout.addWidget(self.label_preview)
        self.button_preview = QPushButton("Pré-visualizar")
        self.button_executar = QPushButton("Executar")
        self.button_executar.setEnabled(False)
        self.button_restaurar_lixeira = QPushButton("Restaurar Lixeira")
        self.button_clear_log = QPushButton("Limpar Log")
        self.button_export_log = QPushButton("Exportar Log")
        self.button_undo = QPushButton("Desfazer")
        self.button_undo.setEnabled(False)
        preview_button_layout.addWidget(self.button_preview)
        preview_button_layout.addWidget(self.button_executar)
        preview_button_layout.addWidget(self.button_restaurar_lixeira)
        preview_button_layout.addWidget(self.button_clear_log)
        preview_button_layout.addWidget(self.button_export_log)
        preview_button_layout.addWidget(self.button_undo)
        preview_button_layout.addStretch()  # Stretch após o último botão (Desfazer)
        self.listbox_preview = QListWidget()
        self.listbox_preview.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.button_remove_preview = QPushButton("Remover")

        visualizacao_layout.addWidget(self.label_log)
        visualizacao_layout.addWidget(self.logbox)
        visualizacao_layout.addLayout(preview_button_layout)
        visualizacao_layout.addWidget(self.listbox_preview)
        visualizacao_layout.addWidget(self.button_remove_preview)
        content_layout.addWidget(self.panel_visualizacao, stretch=1)

        # Adicionar o layout horizontal ao layout principal
        main_layout.addLayout(content_layout)

        # Painel de Ações (Rodapé)
        self.panel_acoes = QWidget()
        acoes_layout = QHBoxLayout(self.panel_acoes)
        acoes_layout.setContentsMargins(0, 0, 0, 0)
        acoes_layout.setSpacing(0)
        self.progress_bar = QProgressBar()
        acoes_layout.addWidget(self.progress_bar)
        self.panel_acoes.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.panel_acoes.setFixedHeight(50)

        # Adicionar o rodapé ao layout principal
        main_layout.addWidget(self.panel_acoes)

        # Conectar eventos
        self.combobox_tema.currentTextChanged.connect(self.apply_theme)
        self.button_add_origem.clicked.connect(self.add_origem)
        self.button_remove_origem.clicked.connect(self.remove_origem)
        self.button_selecionar_destino.clicked.connect(self.select_destination)
        self.button_clear_destino.clicked.connect(self.clear_destination)
        self.button_clear_form.clicked.connect(self.clear_form)
        self.button_salvar_template.clicked.connect(self.save_template)
        self.button_editar_template.clicked.connect(self.edit_template)
        self.button_excluir_template.clicked.connect(self.delete_template)
        self.combobox_templates.currentTextChanged.connect(self.load_template)
        self.button_preview.clicked.connect(self.preview_files)
        self.button_executar.clicked.connect(self.execute)
        self.button_restaurar_lixeira.clicked.connect(self.restore_recycle_bin)
        self.button_remove_preview.clicked.connect(self.remove_preview)
        self.button_clear_log.clicked.connect(self.clear_log)
        self.button_export_log.clicked.connect(self.export_log)
        self.button_undo.clicked.connect(self.undo_action)

        # Ajustar tamanhos dos botões para consistência
        button_size = (100, 20)
        self.button_preview.setFixedSize(*button_size)
        self.button_executar.setFixedSize(*button_size)
        self.button_restaurar_lixeira.setFixedSize(*button_size)
        self.button_clear_log.setFixedSize(*button_size)
        self.button_export_log.setFixedSize(*button_size)
        self.button_undo.setFixedSize(*button_size)
        self.button_remove_preview.setFixedSize(*button_size)

        # Forçar renderização antes de aplicar o tema
        QApplication.instance().processEvents()

        # Aplicar tema inicial após setup
        # self.apply_theme(self.combobox_tema.currentText())

    def apply_theme(self, theme: str):
        if not QApplication.instance().thread() == QThread.currentThread():
            self.logbox.append("Erro: Tentativa de aplicar tema fora do thread principal.\n")
            return
        QApplication.instance().processEvents()
        config = self.load_config()
        config["Settings"]["theme"] = theme
        self.save_config(config["Templates"], config["Settings"])
        stylesheet = """
            QWidget { background: #000000; color: #00FF00; border: none; }
            QPushButton { background: #333333; color: #00FF00; border: 1px solid #333333; }
            QPushButton:hover { background: #808080; }
            QLineEdit, QTextEdit, QListWidget, QComboBox { background: #000000; color: #00FF00; border: 1px solid #333333; }
            QProgressBar { background: #333333; color: #00FF00; border: none; text-align: center; }
        """ if theme == "Neon" else """
            QWidget { background: #F0F0F0; color: #000000; border: none; }
            QPushButton { background: #E0E0E0; color: #000000; border: 1px solid #A0A0A0; }
            QPushButton:hover { background: #D0D0D0; }
            QLineEdit, QTextEdit, QListWidget, QComboBox { background: #FFFFFF; color: #000000; border: 1px solid #A0A0A0; }
            QProgressBar { background: #E0E0E0; color: #000000; border: none; text-align: center; }
        """
        self.setStyleSheet(stylesheet)
    def load_config(self) -> Dict:
        config = configparser.ConfigParser()
        templates = {}
        settings = {"theme": "Neon"}  # Neon como padrão inicial
        if CONFIG_PATH.exists():
            try:
                config.read(CONFIG_PATH, encoding="utf-8-sig")
                for section in config.sections():
                    if section == "Settings":
                        settings = dict(config[section])
                        if "theme" not in settings or settings["theme"] not in ["Neon", "Claro"]:
                            settings["theme"] = "Neon"  # Forçar Neon se inválido
                    else:
                        templates[section] = dict(config[section])
            except configparser.MissingSectionHeaderError:
                self.show_message("Erro: config.ini inválido. Criando novo arquivo.")
                self.save_config(templates, settings)
        else:
            self.save_config(templates, settings)  # Criar config.ini com Neon
        return {"Templates": templates, "Settings": settings}

    def save_config(self, templates: Dict, settings: Dict):
        config = configparser.ConfigParser()
        config["Settings"] = settings
        for name, data in templates.items():
            config[name] = data
        with CONFIG_PATH.open("w", encoding="utf-8") as f:
            config.write(f)

    def show_message(self, text: str, title: str = "Aviso"):
        QMessageBox.information(self, title, text)

    def get_file_hash_md5(self, file_path: Path) -> Optional[str]:
        try:
            with file_path.open("rb") as f:
                md5 = hashlib.md5()
                for chunk in iter(lambda: f.read(4096), b""):
                    md5.update(chunk)
            return md5.hexdigest().lower()
        except:
            return None

    def load_initial_config(self):
        config = self.load_config()
        theme = config["Settings"].get("theme", "Neon")  # Neon como padrão
        if theme not in ["Neon", "Claro"]:  # Validar tema
            theme = "Neon"
        self.combobox_tema.setCurrentText(theme)
        self.apply_theme(theme)
        self.populate_templates_dropdown()

    def populate_templates_dropdown(self):
        current = self.combobox_templates.currentText()
        self.combobox_templates.clear()
        config = self.load_config()
        templates = config["Templates"]
        for key in sorted(templates.keys()):
            self.combobox_templates.addItem(key)
        if templates and current in templates:
            self.combobox_templates.setCurrentText(current)
        elif templates:
            self.combobox_templates.setCurrentIndex(0)

    def load_template(self):
        if self.is_processing_selection:
            return
        self.is_processing_selection = True
        try:
            sel = self.combobox_templates.currentText()
            if not sel:
                return
            config = self.load_config()
            t = config["Templates"].get(sel, {})
            self.listbox_origem.clear()
            if t.get("pastasorigem"):
                for path in t["pastasorigem"].split(";"):
                    if path.strip():
                        self.listbox_origem.addItem(path.strip())
            self.textbox_destino.setText(t.get("pastadestino", ""))
            self.checkbox_mover.setChecked(t.get("moverarquivos", "False") == "True")
            self.checkbox_excluir_duplicatas.setChecked(t.get("excluirduplicatas", "True") == "True")
            self.checkbox_lixeira.setChecked(t.get("usarlixeira", "True") == "True")
            self.checkbox_subpastas.setChecked(t.get("usarsubpastas", "False") == "True")
            self.checkbox_hash.setChecked(t.get("usarhash", "False") == "True")
            self.checkbox_abrir_destino.setChecked(t.get("abrirdestino", "False") == "True")
            self.checkbox_encerrar.setChecked(t.get("encerrarprograma", "False") == "True")
            self.combobox_filtro.setCurrentText(t.get("filtro", "*.jpg;*.png"))
            self.checkbox_regex.setChecked(t.get("usarregex", "False") == "True")
            self.textbox_template_name.setText(sel)
            # Remover redefinição do tema
            # current_theme = config["Settings"].get("theme", "Claro")
            # self.combobox_tema.setCurrentText(current_theme)
            # self.apply_theme(current_theme)
        except Exception as e:
            self.logbox.append(f"Erro ao carregar template: {e}\n")
        finally:
            self.is_processing_selection = False

    def save_template(self):
        template_name = self.textbox_template_name.text().strip()
        if not template_name:
            self.show_message("Digite um nome para o template.")
            return
        config = self.load_config()
        templates = config["Templates"]
        templates[template_name] = self.get_current_settings()
        self.save_config(templates, config["Settings"])
        self.combobox_templates.addItem(template_name)
        self.combobox_templates.setCurrentText(template_name)
        self.show_message(f"Template '{template_name}' salvo com sucesso.")

    def edit_template(self):
        if not self.combobox_templates.currentText():
            self.show_message("Selecione um template para editar.")
            return
        template_name = self.textbox_template_name.text().strip()
        if not template_name:
            self.show_message("Digite um nome para o template.")
            return
        config = self.load_config()
        templates = config["Templates"]
        old_name = self.combobox_templates.currentText()
        templates[template_name] = self.get_current_settings()
        if template_name != old_name:
            templates.pop(old_name, None)
        self.save_config(templates, config["Settings"])
        self.combobox_templates.clear()
        for key in sorted(templates.keys()):
            self.combobox_templates.addItem(key)
        self.combobox_templates.setCurrentText(template_name)
        self.show_message(f"Template '{template_name}' editado com sucesso.")

    def delete_template(self):
        template_name = self.combobox_templates.currentText()
        if not template_name:
            self.show_message("Selecione um template para excluir.")
            return
        reply = QMessageBox.question(self, "Confirmar Exclusão",
                                    f"Deseja excluir o template '{template_name}'?",
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            config = self.load_config()
            templates = config["Templates"]
            templates.pop(template_name, None)
            self.save_config(templates, config["Settings"])
            self.combobox_templates.clear()
            for key in sorted(templates.keys()):
                self.combobox_templates.addItem(key)
            if self.combobox_templates.count() > 0:
                self.combobox_templates.setCurrentIndex(0)
            self.textbox_template_name.clear()
            self.show_message(f"Template '{template_name}' excluído com sucesso.")

    def get_current_settings(self) -> Dict:
        return {
            "pastasorigem": ";".join(self.listbox_origem.item(i).text() for i in range(self.listbox_origem.count())),
            "pastadestino": self.textbox_destino.text(),
            "moverarquivos": str(self.checkbox_mover.isChecked()),
            "excluirduplicatas": str(self.checkbox_excluir_duplicatas.isChecked()),
            "usarlixeira": str(self.checkbox_lixeira.isChecked()),
            "usarsubpastas": str(self.checkbox_subpastas.isChecked()),
            "usarhash": str(self.checkbox_hash.isChecked()),
            "abrirdestino": str(self.checkbox_abrir_destino.isChecked()),
            "encerrarprograma": str(self.checkbox_encerrar.isChecked()),
            "filtro": self.combobox_filtro.currentText(),
            "usarregex": str(self.checkbox_regex.isChecked())
        }

    def add_origem(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecione uma pasta de origem")
        if folder and folder not in [self.listbox_origem.item(i).text() for i in range(self.listbox_origem.count())]:
            self.listbox_origem.addItem(folder)
            self.logbox.append(f"Pasta adicionada: {folder}\n")

    def remove_origem(self):
        selected = self.listbox_origem.currentItem()
        if selected:
            self.logbox.append(f"Pasta removida: {selected.text()}\n")
            self.listbox_origem.takeItem(self.listbox_origem.row(selected))

    def select_destination(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecione a pasta de destino")
        if folder:
            self.textbox_destino.setText(folder)
            self.logbox.append(f"Destino selecionado: {folder}\n")

    def clear_destination(self):
        self.textbox_destino.clear()
        self.logbox.append("Pasta de destino limpa\n")

    def clear_form(self):
        self.listbox_origem.clear()
        self.textbox_destino.clear()
        self.checkbox_mover.setChecked(False)
        self.checkbox_excluir_duplicatas.setChecked(True)
        self.checkbox_lixeira.setChecked(True)
        self.checkbox_subpastas.setChecked(False)
        self.checkbox_hash.setChecked(False)
        self.checkbox_abrir_destino.setChecked(False)
        self.checkbox_encerrar.setChecked(False)
        self.combobox_filtro.setCurrentText("*.jpg;*.png")
        self.checkbox_regex.setChecked(False)
        self.textbox_template_name.clear()
        self.listbox_preview.clear()
        self.button_executar.setEnabled(False)
        self.logbox.append("Formulário limpo\n")

    def clear_log(self):
        self.logbox.clear()
        self.logbox.append("Log limpo\n")

    def export_log(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "Salvar Log", f"log_{os.path.basename(__file__)[:-3]}_{os.time.strftime('%Y%m%d_%H%M%S')}.txt", "Text files (*.txt);;CSV files (*.csv)")
        if file_name:
            with open(file_name, "w", encoding="utf-8") as f:
                f.write(self.logbox.toPlainText())
            self.logbox.append(f"Log exportado para {file_name}\n")

    def restore_recycle_bin(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Restaurar Arquivos")
        dialog.setMinimumSize(800, 600)
        layout = QVBoxLayout(dialog)

        listbox_restore = QListWidget()
        listbox_restore.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(listbox_restore)

        button_restore = QPushButton("Restaurar")
        layout.addWidget(button_restore)

        # Listar arquivos da lixeira do sistema
        shell = win32com.client.Dispatch("Shell.Application")
        recycle_bin = shell.NameSpace(0xA)  # Lixeira
        items = recycle_bin.Items()
        for item in items:
            listbox_restore.addItem(f"{item.Name} ({item.Path})")

        # Listar arquivos da lixeira local
        local_trash = Path(self.textbox_destino.text()) / "Lixeira"
        if local_trash.exists():
            for file in local_trash.iterdir():
                if file.is_file():
                    listbox_restore.addItem(f"{file.name} ({file})")

        def restore_selected():
            for item in listbox_restore.selectedItems():
                item_text = item.text()
                if item_text.endswith(")"):
                    file_name = item_text.split(" (")[0]
                    file_path = item_text.split(" (")[1][:-1]
                    if str(local_trash) in file_path:
                        # Restaurar da lixeira local
                        dest = Path(self.textbox_destino.text()) / file_name
                        shutil.move(file_path, dest)
                        self.logbox.append(f"Restaurado (lixeira local): {file_name} -> {dest}\n")
                    else:
                        # Restaurar da lixeira do sistema
                        for rb_item in recycle_bin.Items():
                            if rb_item.Name == file_name:
                                rb_item.InvokeVerb("Restore")
                                self.logbox.append(f"Restaurado (lixeira sistema): {file_name}\n")
                                break
                listbox_restore.takeItem(listbox_restore.row(item))
            if listbox_restore.count() == 0:
                dialog.accept()

        button_restore.clicked.connect(restore_selected)
        dialog.exec()

    def remove_preview(self):
        selected = self.listbox_preview.selectedItems()
        if selected:
            for item in selected:
                self.listbox_preview.takeItem(self.listbox_preview.row(item))
            self.button_executar.setEnabled(self.listbox_preview.count() > 0)
            self.logbox.append(f"Removidas {len(selected)} ações da pré-visualização.\n")

    def undo_action(self):
        if self.undo_stack:
            action = self.undo_stack.pop()
            if action["action"] == "move":
                shutil.move(action["dest"], action["source"])
                self.logbox.append(f"Desfeito: Movido {action['dest']} -> {action['source']}\n")
            elif action["action"] == "copy":
                os.remove(action["dest"])
                self.logbox.append(f"Desfeito: Removido {action['dest']}\n")
            self.button_undo.setEnabled(len(self.undo_stack) > 0)

    def preview_files(self):
        self.listbox_preview.clear()
        self.button_executar.setEnabled(False)
        if self.listbox_origem.count() == 0:
            self.show_message("Adicione pelo menos uma pasta de origem.")
            return
        if not self.textbox_destino.text():
            self.show_message("Selecione uma pasta de destino.")
            return
        count = 0
        filter_pattern = self.combobox_filtro.currentText()
        use_regex = self.checkbox_regex.isChecked()
        if use_regex:
            try:
                re.compile(filter_pattern)
            except re.error:
                self.show_message("Expressão regular inválida.")
                return

        for i in range(self.listbox_origem.count()):
            origem = Path(self.listbox_origem.item(i).text())
            if not origem.exists():
                self.listbox_preview.addItem(f"Erro: Pasta de origem não encontrada: {origem}")
                continue
            for file_path in origem.rglob("*"):
                if not file_path.is_file():
                    continue
                if use_regex:
                    if not re.search(filter_pattern, file_path.name):
                        continue
                else:
                    patterns = [p.strip() for p in filter_pattern.split(";") if p.strip()]
                    if not any(file_path.match(p) for p in patterns):
                        continue
                extension = file_path.suffix.lstrip('.').lower()
                dest_folder = Path(self.textbox_destino.text()) / extension if self.checkbox_subpastas.isChecked() and extension else Path(self.textbox_destino.text())
                dest_file = dest_folder / file_path.name
                action = "Mover" if self.checkbox_mover.isChecked() else "Copiar"
                if dest_file.exists():
                    if self.checkbox_excluir_duplicatas.isChecked():
                        is_duplicate = True
                        if self.checkbox_hash.isChecked():
                            src_hash = self.get_file_hash_md5(file_path)
                            dest_hash = self.get_file_hash_md5(dest_file)
                            is_duplicate = src_hash and dest_hash and src_hash == dest_hash
                        if is_duplicate:
                            if self.checkbox_mover.isChecked():
                                destino = "Lixeira" if self.checkbox_lixeira.isChecked() else "Permanentemente"
                                self.listbox_preview.addItem(f"Excluir duplicata: {file_path} -> {destino}")
                            else:
                                self.listbox_preview.addItem(f"Pular duplicata: {file_path}")
                        else:
                            dest_file = self.get_unique_filename(dest_file)
                            self.listbox_preview.addItem(f"{action} (renomear): {file_path} -> {dest_file}")
                    else:
                        dest_file = self.get_unique_filename(dest_file)
                        self.listbox_preview.addItem(f"{action} (renomear): {file_path} -> {dest_file}")
                else:
                    self.listbox_preview.addItem(f"{action}: {file_path} -> {dest_file}")
                count += 1
        if count > 0:
            self.button_executar.setEnabled(True)
            self.logbox.append(f"Pré-visualização gerada: {count} ações\n")
        else:
            self.logbox.append("Pré-visualização vazia: nenhum arquivo encontrado\n")

    def get_unique_filename(self, dest_file: Path) -> Path:
        base_name = dest_file.stem
        extension = dest_file.suffix
        counter = 1
        new_file = dest_file
        while new_file.exists():
            new_file = dest_file.parent / f"{base_name}_{counter}{extension}"
            counter += 1
        return new_file

    def move_or_copy_file(self, src: Path, dest: Path, move: bool):
        if move:
            shutil.move(src, dest)
            self.undo_stack.append({"action": "move", "source": str(src), "dest": str(dest)})
            self.logbox.append(f"Movido: {src} -> {dest}\n")
        else:
            shutil.copy2(src, dest)
            self.undo_stack.append({"action": "copy", "source": str(src), "dest": str(dest)})
            self.logbox.append(f"Copiado: {src} -> {dest}\n")
        self.button_undo.setEnabled(True)

    def process_files(self):
        self.logbox.clear()
        pastas_origem = [Path(self.listbox_origem.item(i).text()) for i in range(self.listbox_origem.count())]
        pasta_destino = Path(self.textbox_destino.text())
        filter_pattern = self.combobox_filtro.currentText()
        mover_arquivos = self.checkbox_mover.isChecked()
        excluir_duplicatas = self.checkbox_excluir_duplicatas.isChecked()
        usar_lixeira = self.checkbox_lixeira.isChecked()
        usar_subpastas = self.checkbox_subpastas.isChecked()
        usar_hash = self.checkbox_hash.isChecked()
        usar_regex = self.checkbox_regex.isChecked()

        # Contar arquivos pra barra de progresso
        total_files = 0
        for origem in pastas_origem:
            for file_path in origem.rglob("*"):
                if file_path.is_file():
                    if usar_regex:
                        if re.search(filter_pattern, file_path.name):
                            total_files += 1
                    else:
                        patterns = [p.strip() for p in filter_pattern.split(";") if p.strip()]
                        if any(file_path.match(p) for p in patterns):
                            total_files += 1
        self.progress_bar.setMaximum(total_files)
        current_file = 0

        for origem in pastas_origem:
            if not origem.exists():
                self.logbox.append(f"Pasta de origem não encontrada: {origem}\n")
                continue
            for file_path in origem.rglob("*"):
                if not file_path.is_file():
                    continue
                if usar_regex:
                    if not re.search(filter_pattern, file_path.name):
                        continue
                else:
                    patterns = [p.strip() for p in filter_pattern.split(";") if p.strip()]
                    if not any(file_path.match(p) for p in patterns):
                        continue
                current_file += 1
                self.progress_bar.setValue(current_file)
                extension = file_path.suffix.lstrip('.').lower()
                dest_folder = pasta_destino / extension if usar_subpastas and extension else pasta_destino
                dest_folder.mkdir(parents=True, exist_ok=True)
                dest_file = dest_folder / file_path.name
                if dest_file.exists():
                    if excluir_duplicatas:
                        is_duplicate = True
                        if usar_hash:
                            src_hash = self.get_file_hash_md5(file_path)
                            dest_hash = self.get_file_hash_md5(dest_file)
                            is_duplicate = src_hash and dest_hash and src_hash == dest_hash
                        if is_duplicate:
                            if mover_arquivos:
                                src_mtime = os.path.getmtime(file_path)
                                dest_mtime = os.path.getmtime(dest_file)
                                if src_mtime < dest_mtime:
                                    file_to_delete = file_path
                                    action = "source"
                                else:
                                    file_to_delete = dest_file
                                    action = "dest"
                                if usar_lixeira:
                                    try:
                                        send2trash(str(file_to_delete))
                                        self.logbox.append(f"Duplicata movida para lixeira: {file_to_delete}\n")
                                    except:
                                        lixeira_local = pasta_destino / "Lixeira"
                                        lixeira_local.mkdir(exist_ok=True)
                                        shutil.move(file_to_delete, lixeira_local / file_to_delete.name)
                                        self.logbox.append(f"Duplicata movida para lixeira local: {file_to_delete}\n")
                                else:
                                    if QMessageBox.question(self, "Confirmar Exclusão",
                                                           f"Excluir permanentemente '{file_to_delete}'?",
                                                           QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
                                        file_to_delete.unlink()
                                        self.logbox.append(f"Duplicata excluída permanentemente: {file_to_delete}\n")
                                if action == "source":
                                    continue  # Pula a movimentação/cópia
                                else:
                                    self.move_or_copy_file(file_path, dest_file, mover_arquivos)
                            else:
                                self.logbox.append(f"Pulado duplicata: {file_path}\n")
                                continue
                        else:
                            dest_file = self.get_unique_filename(dest_file)
                            self.move_or_copy_file(file_path, dest_file, mover_arquivos)
                    else:
                        dest_file = self.get_unique_filename(dest_file)
                        self.move_or_copy_file(file_path, dest_file, mover_arquivos)
                else:
                    self.move_or_copy_file(file_path, dest_file, mover_arquivos)
        self.progress_bar.setValue(0)
        self.logbox.append("Processamento concluído.\n")

    def execute(self):
        if self.listbox_origem.count() == 0:
            self.show_message("Adicione pelo menos uma pasta de origem.")
            return
        destino = Path(self.textbox_destino.text())
        if not destino.exists():
            try:
                destino.mkdir(parents=True)
                self.logbox.append(f"Pasta destino criada: {destino}\n")
            except Exception as e:
                self.show_message(f"Não foi possível criar a pasta destino: {e}")
                return
        self.button_executar.setEnabled(False)
        try:
            self.process_files()
            if self.checkbox_abrir_destino.isChecked():
                os.startfile(str(destino))
            if self.checkbox_encerrar.isChecked():
                QApplication.quit()
        except Exception as e:
            self.show_message(f"Erro durante execução: {e}")
        finally:
            self.button_executar.setEnabled(False)
            self.listbox_preview.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileOrganizerApp()
    window.show()
    sys.exit(app.exec())