import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QListWidget, QPushButton, QTextEdit, QLineEdit, QCheckBox, 
                             QComboBox, QFileDialog, QMessageBox, QProgressBar, QDialog, QSizePolicy)
from PyQt6.QtCore import QThread
from PyQt6.QtGui import QFont

class FileOrganizerUI(QMainWindow):
    """Classe que representa a interface do usuário da aplicação"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Organizador de Arquivos v3.4")
        self.setGeometry(100, 100, 1280, 720)
        self.setFont(QFont("Consolas", 10))
        self.is_processing_selection = False
        self.undo_stack = []
        self.setup_ui()
    
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
        self.combobox_filtro.addItems([
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
        ])
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

        # Ajustar tamanhos dos botões para consistência
        button_size = (100, 20)
        self.button_preview.setFixedSize(*button_size)
        self.button_executar.setFixedSize(*button_size)
        self.button_restaurar_lixeira.setFixedSize(*button_size)
        self.button_clear_log.setFixedSize(*button_size)
        self.button_export_log.setFixedSize(*button_size)
        self.button_undo.setFixedSize(*button_size)
        self.button_remove_preview.setFixedSize(*button_size)