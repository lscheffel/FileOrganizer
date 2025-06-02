import sys
import os
from pathlib import Path
# Importações dos módulos criados
from config.config_manager import ConfigManager
from data.file_organizer import FileOrganizer
from ui.file_organizer_ui import FileOrganizerUI
from utils.recycle_bin_manager import RecycleBinManager

# Mantém as importações originais
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QListWidget, QPushButton, QTextEdit, QLineEdit, QCheckBox, 
                             QComboBox, QFileDialog, QMessageBox, QProgressBar, QDialog, QSizePolicy)
from PyQt6.QtCore import QThread
from PyQt6.QtGui import QFont
from send2trash import send2trash
import win32com.client
import hashlib
from typing import Dict, Optional
from pathlib import Path
import configparser
import re
from typing import Dict, Optional
import shutil
import os
import sys
import time  # Importando o módulo time

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

class FileOrganizerApp(FileOrganizerUI):
    """Classe principal da aplicação que integra UI e lógica de organização de arquivos"""
    
    def __init__(self):
        super().__init__()
        self.config_manager = ConfigManager()
        self.file_organizer = FileOrganizer(self)
        self.recycle_bin_manager = RecycleBinManager()
        self.setup_connections()
        self.load_initial_config()
    
    def setup_connections(self):
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
        self.button_restaurar_lixeira.clicked.connect(lambda: self.recycle_bin_manager.restore_files(self.textbox_destino))
        self.button_remove_preview.clicked.connect(self.remove_preview)
        self.button_clear_log.clicked.connect(self.clear_log)
        self.button_export_log.clicked.connect(self.export_log)
        self.button_undo.clicked.connect(self.undo_action)
    
    def apply_theme(self, theme: str):
        """Aplica o tema selecionado à interface"""
        if not QApplication.instance().thread() == QThread.currentThread():
            self.logbox.append("Erro: Tentativa de aplicar tema fora do thread principal.\n")
            return
        QApplication.instance().processEvents()
        config = self.config_manager.load_config()
        config["Settings"]["theme"] = theme
        self.config_manager.save_config(config["Templates"], config["Settings"])
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
    
    def load_config(self):
        """Carrega as configurações do arquivo ini"""
        return self.config_manager.load_config()
    
    def save_config(self, templates, settings):
        """Salva as configurações no arquivo ini"""
        self.config_manager.save_config(templates, settings)
    
    def get_current_settings(self):
        """Obtém as configurações atuais da aplicação"""
        return self.config_manager.get_current_settings(self)
    
    def load_initial_config(self):
        """Carrega as configurações iniciais da aplicação"""
        config = self.load_config()
        theme = config["Settings"].get("theme", "Neon")  # Neon como padrão
        if theme not in ["Neon", "Claro"]:  # Validar tema
            theme = "Neon"
        self.combobox_tema.setCurrentText(theme)
        self.apply_theme(theme)
        self.populate_templates_dropdown()
    
    def process_files(self):
        """Processa os arquivos de acordo com as configurações da aplicação"""
        self.file_organizer.process_files()
    
    def execute(self):
        """Executa a organização dos arquivos"""
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
        file_name, _ = QFileDialog.getSaveFileName(self, "Salvar Log", f"log_{os.path.basename(__file__)[:-3]}_{time.strftime('%Y%m%d_%H%M%S')}.txt", "Text files (*.txt);;CSV files (*.csv)")
        if file_name:
            with open(file_name, "w", encoding="utf-8") as f:
                f.write(self.logbox.toPlainText())
            self.logbox.append(f"Log exportado para {file_name}\n")

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
        except Exception as e:
            self.logbox.append(f"Erro ao carregar template: {e}\n")
        finally:
            self.is_processing_selection = False

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
                            src_hash = self.file_organizer.get_file_hash_md5(file_path)
                            dest_hash = self.file_organizer.get_file_hash_md5(dest_file)
                            is_duplicate = src_hash and dest_hash and src_hash == dest_hash
                        if is_duplicate:
                            if self.checkbox_mover.isChecked():
                                destino = "Lixeira" if self.checkbox_lixeira.isChecked() else "Permanentemente"
                                self.listbox_preview.addItem(f"Excluir duplicata: {file_path} -> {destino}")
                            else:
                                self.listbox_preview.addItem(f"Pular duplicata: {file_path}")
                        else:
                            dest_file = self.file_organizer.get_unique_filename(dest_file)
                            self.listbox_preview.addItem(f"{action} (renomear): {file_path} -> {dest_file}")
                    else:
                        dest_file = self.file_organizer.get_unique_filename(dest_file)
                        self.listbox_preview.addItem(f"{action} (renomear): {file_path} -> {dest_file}")
                else:
                    self.listbox_preview.addItem(f"{action}: {file_path} -> {dest_file}")
                count += 1
        if count > 0:
            self.button_executar.setEnabled(True)
            self.logbox.append(f"Pré-visualização gerada: {count} ações\n")
        else:
            self.logbox.append("Pré-visualização vazia: nenhum arquivo encontrado\n")

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

    def show_message(self, text: str, title: str = "Aviso"):
        QMessageBox.information(self, title, text)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileOrganizerApp()
    window.show()
    sys.exit(app.exec())