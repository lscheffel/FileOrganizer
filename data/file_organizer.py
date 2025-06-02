import os
import shutil
from pathlib import Path
import hashlib
import re
from typing import Dict, Optional, List
from send2trash import send2trash

class FileOrganizer:
    """Classe que contém a lógica principal de organização de arquivos"""
    
    def __init__(self, app_state):
        self.app_state = app_state
        self.undo_stack = []
    
    def get_file_hash_md5(self, file_path: Path) -> Optional[str]:
        """Calcula o hash MD5 de um arquivo"""
        try:
            with file_path.open("rb") as f:
                md5 = hashlib.md5()
                for chunk in iter(lambda: f.read(4096), b""):
                    md5.update(chunk)
            return md5.hexdigest().lower()
        except:
            return None
    
    def get_unique_filename(self, dest_file: Path) -> Path:
        """Gera um nome único para um arquivo que já existe"""
        base_name = dest_file.stem
        extension = dest_file.suffix
        counter = 1
        new_file = dest_file
        while new_file.exists():
            new_file = dest_file.parent / f"{base_name}_{counter}{extension}"
            counter += 1
        return new_file
    
    def move_or_copy_file(self, src: Path, dest: Path, move: bool):
        """Move ou copia um arquivo e adiciona a ação ao stack de desfazer"""
        if move:
            shutil.move(src, dest)
            self.undo_stack.append({"action": "move", "source": str(src), "dest": str(dest)})
            print(f"Movido: {src} -> {dest}\n")
        else:
            shutil.copy2(src, dest)
            self.undo_stack.append({"action": "copy", "source": str(src), "dest": str(dest)})
            print(f"Copiado: {src} -> {dest}\n")
    
    def process_files(self):
        """Processa os arquivos de acordo com as configurações da aplicação"""
        pastas_origem = [Path(self.app_state.listbox_origem.item(i).text()) for i in range(self.app_state.listbox_origem.count())]
        pasta_destino = Path(self.app_state.textbox_destino.text())
        filter_pattern = self.app_state.combobox_filtro.currentText()
        mover_arquivos = self.app_state.checkbox_mover.isChecked()
        excluir_duplicatas = self.app_state.checkbox_excluir_duplicatas.isChecked()
        usar_lixeira = self.app_state.checkbox_lixeira.isChecked()
        usar_subpastas = self.app_state.checkbox_subpastas.isChecked()
        usar_hash = self.app_state.checkbox_hash.isChecked()
        usar_regex = self.app_state.checkbox_regex.isChecked()

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
        self.app_state.progress_bar.setMaximum(total_files)
        current_file = 0

        for origem in pastas_origem:
            if not origem.exists():
                print(f"Pasta de origem não encontrada: {origem}\n")
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
                self.app_state.progress_bar.setValue(current_file)
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
                                        print(f"Duplicata movida para lixeira: {file_to_delete}\n")
                                    except:
                                        lixeira_local = pasta_destino / "Lixeira"
                                        lixeira_local.mkdir(exist_ok=True)
                                        shutil.move(file_to_delete, lixeira_local / file_to_delete.name)
                                        print(f"Duplicata movida para lixeira local: {file_to_delete}\n")
                                else:
                                    print(f"Duplicata excluída permanentemente: {file_to_delete}\n")
                                if action == "source":
                                    continue  # Pula a movimentação/cópia
                                else:
                                    self.move_or_copy_file(file_path, dest_file, mover_arquivos)
                            else:
                                print(f"Pulado duplicata: {file_path}\n")
                                continue
                        else:
                            dest_file = self.get_unique_filename(dest_file)
                            self.move_or_copy_file(file_path, dest_file, mover_arquivos)
                    else:
                        dest_file = self.get_unique_filename(dest_file)
                        self.move_or_copy_file(file_path, dest_file, mover_arquivos)
                else:
                    self.move_or_copy_file(file_path, dest_file, mover_arquivos)
        self.app_state.progress_bar.setValue(0)
        print("Processamento concluído.\n")