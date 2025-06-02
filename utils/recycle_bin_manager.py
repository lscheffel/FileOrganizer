import win32com.client
from pathlib import Path
import shutil

class RecycleBinManager:
    """Classe para gerenciar operações de restauração da lixeira"""
    
    def restore_files(self, textbox_destino):
        """Abre um diálogo para restaurar arquivos da lixeira"""
        dialog = QDialog()
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
        local_trash = Path(textbox_destino.text()) / "Lixeira"
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
                        dest = Path(textbox_destino.text()) / file_name
                        shutil.move(file_path, dest)
                        print(f"Restaurado (lixeira local): {file_name} -> {dest}\n")
                    else:
                        # Restaurar da lixeira do sistema
                        for rb_item in recycle_bin.Items():
                            if rb_item.Name == file_name:
                                rb_item.InvokeVerb("Restore")
                                print(f"Restaurado (lixeira sistema): {file_name}\n")
                                break
                listbox_restore.takeItem(listbox_restore.row(item))
            if listbox_restore.count() == 0:
                dialog.accept()

        button_restore.clicked.connect(restore_selected)
        dialog.exec()