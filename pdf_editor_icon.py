import os
import sys
import ctypes
from dataclasses import dataclass
from typing import List

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QListWidget,
    QPushButton, QFileDialog, QLineEdit, QMessageBox, QDialog,
    QListWidgetItem, QInputDialog
)

from pypdf import PdfReader, PdfWriter


def resource_path(relative_path: str) -> str:
    """
    Resolve caminhos tanto no modo .py quanto quando empacotado via PyInstaller.
    Se o ícone estiver ao lado do .py (no projeto) e você usar --add-data,
    esse helper também acha no executável.
    """
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def is_pdf(path: str) -> bool:
    return path.lower().endswith(".pdf") and os.path.isfile(path)


def parse_page_ranges(spec: str, max_pages: int) -> List[int]:
    """
    Converte '1-3, 5, 7-9' em lista de índices 0-based, validando.
    max_pages = total de páginas do arquivo.
    """
    spec = spec.strip()
    if not spec:
        return list(range(max_pages))

    pages = set()
    parts = [p.strip() for p in spec.split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            a, b = [x.strip() for x in part.split("-", 1)]
            if not a.isdigit() or not b.isdigit():
                raise ValueError(f"Intervalo inválido: {part}")
            start = int(a)
            end = int(b)
            if start < 1 or end < 1 or start > end:
                raise ValueError(f"Intervalo inválido: {part}")
            for n in range(start, end + 1):
                if n > max_pages:
                    raise ValueError(f"Página {n} excede o total ({max_pages}).")
                pages.add(n - 1)
        else:
            if not part.isdigit():
                raise ValueError(f"Página inválida: {part}")
            n = int(part)
            if n < 1 or n > max_pages:
                raise ValueError(f"Página {n} excede o total ({max_pages}).")
            pages.add(n - 1)

    return sorted(pages)


@dataclass
class PageRef:
    file_path: str
    file_name: str
    page_index: int  # 0-based
    rotation: int = 0  # 0, 90, 180, 270


class PageEditorDialog(QDialog):
    def __init__(self, parent, pages: List[PageRef], icon: QIcon):
        super().__init__(parent)
        self.setWindowTitle("Editar antes de juntar (páginas)")
        self.resize(850, 500)
        self.setWindowIcon(icon)

        self.pages: List[PageRef] = list(pages)

        layout = QVBoxLayout(self)

        info = QLabel(
            "Aqui você edita a SEQUÊNCIA FINAL de páginas.\n"
            "• Reordene, remova e rotacione páginas antes de salvar o PDF juntado.\n"
            "• Dica: para filtrar páginas de um arquivo específico, use o botão 'Selecionar páginas por arquivo…'."
        )
        info.setStyleSheet("color: #555;")
        layout.addWidget(info)

        self.list_pages = QListWidget()
        self.list_pages.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        layout.addWidget(self.list_pages)

        btn_row = QHBoxLayout()

        self.btn_up = QPushButton("↑ Subir")
        self.btn_up.clicked.connect(self.move_up)
        btn_row.addWidget(self.btn_up)

        self.btn_down = QPushButton("↓ Descer")
        self.btn_down.clicked.connect(self.move_down)
        btn_row.addWidget(self.btn_down)

        self.btn_remove = QPushButton("Remover")
        self.btn_remove.clicked.connect(self.remove_selected)
        btn_row.addWidget(self.btn_remove)

        self.btn_rot90 = QPushButton("Rotacionar +90°")
        self.btn_rot90.clicked.connect(lambda: self.rotate_selected(90))
        btn_row.addWidget(self.btn_rot90)

        self.btn_rot270 = QPushButton("Rotacionar -90°")
        self.btn_rot270.clicked.connect(lambda: self.rotate_selected(-90))
        btn_row.addWidget(self.btn_rot270)

        btn_row.addStretch(1)

        self.btn_pick_ranges = QPushButton("Selecionar páginas por arquivo…")
        self.btn_pick_ranges.clicked.connect(self.pick_ranges_for_file)
        btn_row.addWidget(self.btn_pick_ranges)

        layout.addLayout(btn_row)

        ok_row = QHBoxLayout()
        ok_row.addStretch(1)

        self.btn_ok = QPushButton("OK")
        self.btn_ok.setStyleSheet("font-weight: 600; padding: 8px;")
        self.btn_ok.clicked.connect(self.accept)
        ok_row.addWidget(self.btn_ok)

        self.btn_cancel = QPushButton("Cancelar")
        self.btn_cancel.clicked.connect(self.reject)
        ok_row.addWidget(self.btn_cancel)

        layout.addLayout(ok_row)

        self.refresh()

    def refresh(self):
        self.list_pages.clear()
        for pr in self.pages:
            pnum = pr.page_index + 1
            rot = pr.rotation % 360
            rot_txt = f" | rot {rot}°" if rot else ""
            item = QListWidgetItem(f"{pr.file_name} — página {pnum}{rot_txt}")
            item.setData(Qt.ItemDataRole.UserRole, pr)
            self.list_pages.addItem(item)

    def selected_rows(self) -> List[int]:
        return sorted({self.list_pages.row(i) for i in self.list_pages.selectedItems()})

    def move_up(self):
        rows = self.selected_rows()
        if not rows or rows[0] == 0:
            return
        for r in rows:
            self.pages[r - 1], self.pages[r] = self.pages[r], self.pages[r - 1]
        self.refresh()
        for r in [x - 1 for x in rows]:
            self.list_pages.item(r).setSelected(True)

    def move_down(self):
        rows = self.selected_rows()
        if not rows or rows[-1] == len(self.pages) - 1:
            return
        for r in reversed(rows):
            self.pages[r + 1], self.pages[r] = self.pages[r], self.pages[r + 1]
        self.refresh()
        for r in [x + 1 for x in rows]:
            self.list_pages.item(r).setSelected(True)

    def remove_selected(self):
        rows = self.selected_rows()
        if not rows:
            return
        for r in reversed(rows):
            self.pages.pop(r)
        self.refresh()

    def rotate_selected(self, delta: int):
        rows = self.selected_rows()
        if not rows:
            return
        for r in rows:
            self.pages[r].rotation = (self.pages[r].rotation + delta) % 360
        self.refresh()
        for r in rows:
            self.list_pages.item(r).setSelected(True)

    def pick_ranges_for_file(self):
        files = sorted({p.file_path for p in self.pages})
        if not files:
            return

        names = [os.path.basename(f) for f in files]
        chosen_name, ok = QInputDialog.getItem(
            self, "Escolher arquivo", "Selecione o arquivo para filtrar páginas:", names, 0, False
        )
        if not ok or not chosen_name:
            return

        chosen_path = None
        for f in files:
            if os.path.basename(f) == chosen_name:
                chosen_path = f
                break
        if not chosen_path:
            return

        try:
            total = len(PdfReader(chosen_path).pages)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não consegui ler o PDF:\n{e}")
            return

        spec, ok = QInputDialog.getText(
            self,
            "Intervalos de páginas",
            f"Digite páginas para MANTER (ex.: 1-3, 7, 10-12).\nTotal: {total}\n"
            "Deixe em branco para manter todas:",
        )
        if not ok:
            return

        try:
            keep = parse_page_ranges(spec, total)
        except Exception as e:
            QMessageBox.warning(self, "Intervalo inválido", str(e))
            return

        old_indexes = [i for i, p in enumerate(self.pages) if p.file_path == chosen_path]
        if not old_indexes:
            return

        insert_at = old_indexes[0]
        for i in reversed(old_indexes):
            self.pages.pop(i)

        new_refs = [PageRef(chosen_path, chosen_name, pi, 0) for pi in keep]
        for offset, pr in enumerate(new_refs):
            self.pages.insert(insert_at + offset, pr)

        self.refresh()

    def get_pages(self) -> List[PageRef]:
        return self.pages


class DropListWidget(QListWidget):
    def __init__(self, status_label: QLabel):
        super().__init__()
        self.status_label = status_label
        self.setAcceptDrops(True)
        self.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        added = 0
        ignored = 0

        existing = {self.item(i).text() for i in range(self.count())}

        for url in urls:
            path = url.toLocalFile()
            if is_pdf(path) and path not in existing:
                self.addItem(path)
                existing.add(path)
                added += 1
            else:
                ignored += 1

        self.status_label.setText(
            f"Arquivos na lista: {self.count()} | Adicionados agora: {added} | Ignorados: {ignored}"
        )
        event.acceptProposedAction()


class PdfMergerApp(QWidget):
    def __init__(self, icon: QIcon):
        super().__init__()
        self.icon = icon
        self.setWindowTitle("PDF Editor")
        self.resize(850, 520)

        # Ícone na janela (barra superior)
        self.setWindowIcon(self.icon)

        self.pages_sequence: List[PageRef] = []

        layout = QVBoxLayout(self)

        title = QLabel("Arraste e solte PDFs abaixo • Depois clique em “Editar páginas…”")
        title.setStyleSheet("font-size: 16px; font-weight: 600;")
        layout.addWidget(title)

        self.status = QLabel("Arquivos na lista: 0")
        layout.addWidget(self.status)

        self.list_files = DropListWidget(self.status)
        self.list_files.setStyleSheet("QListWidget { border: 2px dashed #999; padding: 8px; }")
        layout.addWidget(self.list_files)

        out_row = QHBoxLayout()
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("Escolha o nome/local do PDF final (ex.: transcript_unificado.pdf)")
        out_row.addWidget(QLabel("Arquivo final:"))
        out_row.addWidget(self.output_path)

        btn_choose = QPushButton("Escolher…")
        btn_choose.clicked.connect(self.choose_output_file)
        out_row.addWidget(btn_choose)
        layout.addLayout(out_row)

        btn_row = QHBoxLayout()

        btn_add = QPushButton("Adicionar PDFs…")
        btn_add.clicked.connect(self.add_files_dialog)
        btn_row.addWidget(btn_add)

        btn_remove = QPushButton("Remover selecionados")
        btn_remove.clicked.connect(self.remove_selected_files)
        btn_row.addWidget(btn_remove)

        btn_clear = QPushButton("Limpar lista")
        btn_clear.clicked.connect(self.clear_files)
        btn_row.addWidget(btn_clear)

        btn_row.addStretch(1)

        btn_edit = QPushButton("Editar páginas…")
        btn_edit.setStyleSheet("font-weight: 600;")
        btn_edit.clicked.connect(self.open_page_editor)
        btn_row.addWidget(btn_edit)

        btn_save = QPushButton("Salvar PDF")
        btn_save.setStyleSheet("font-weight: 600; padding: 8px;")
        btn_save.clicked.connect(self.merge_and_save)
        btn_row.addWidget(btn_save)

        layout.addLayout(btn_row)

        hint = QLabel("Fluxo: 1) PDFs na lista  2) “Editar páginas…” (opcional)  3) “Salvar PDF”.")
        hint.setStyleSheet("color: #555;")
        layout.addWidget(hint)

    def current_files(self) -> List[str]:
        return [self.list_files.item(i).text() for i in range(self.list_files.count())]

    def build_default_pages_sequence(self) -> List[PageRef]:
        seq: List[PageRef] = []
        for fp in self.current_files():
            name = os.path.basename(fp)
            reader = PdfReader(fp)
            for i in range(len(reader.pages)):
                seq.append(PageRef(fp, name, i, 0))
        return seq

    def add_files_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar PDFs", "", "PDF (*.pdf)")
        if not files:
            return
        existing = {self.list_files.item(i).text() for i in range(self.list_files.count())}
        added = 0
        for f in files:
            if is_pdf(f) and f not in existing:
                self.list_files.addItem(f)
                added += 1
        self.status.setText(f"Arquivos na lista: {self.list_files.count()} | Adicionados agora: {added}")
        self.pages_sequence = []

    def remove_selected_files(self):
        selected = self.list_files.selectedItems()
        if not selected:
            return
        for item in selected:
            self.list_files.takeItem(self.list_files.row(item))
        self.status.setText(f"Arquivos na lista: {self.list_files.count()}")
        self.pages_sequence = []

    def clear_files(self):
        self.list_files.clear()
        self.status.setText("Arquivos na lista: 0")
        self.pages_sequence = []

    def choose_output_file(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Salvar PDF juntado como…", "transcript_unificado.pdf", "PDF (*.pdf)"
        )
        if path:
            if not path.lower().endswith(".pdf"):
                path += ".pdf"
            self.output_path.setText(path)

    def open_page_editor(self):
        if self.list_files.count() == 0:
            QMessageBox.warning(self, "Sem arquivos", "Arraste/adicione ao menos 1 PDF na lista.")
            return

        try:
            if not self.pages_sequence:
                self.pages_sequence = self.build_default_pages_sequence()

            dlg = PageEditorDialog(self, self.pages_sequence, self.icon)
            if dlg.exec() == QDialog.DialogCode.Accepted:
                self.pages_sequence = dlg.get_pages()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não consegui preparar as páginas:\n{e}")

    def merge_and_save(self):
        if self.list_files.count() == 0:
            QMessageBox.warning(self, "Sem arquivos", "Arraste/adicione ao menos 1 PDF na lista.")
            return

        out = self.output_path.text().strip()
        if not out:
            QMessageBox.warning(self, "Arquivo final", "Escolha o nome/local do arquivo final.")
            return
        if not out.lower().endswith(".pdf"):
            out += ".pdf"
            self.output_path.setText(out)

        try:
            if not self.pages_sequence:
                self.pages_sequence = self.build_default_pages_sequence()

            writer = PdfWriter()
            readers = {}

            for pr in self.pages_sequence:
                if pr.file_path not in readers:
                    readers[pr.file_path] = PdfReader(pr.file_path)

                page = readers[pr.file_path].pages[pr.page_index]
                rot = pr.rotation % 360
                if rot:
                    page = page.rotate(rot)

                writer.add_page(page)

            out_dir = os.path.dirname(out)
            if out_dir and not os.path.exists(out_dir):
                os.makedirs(out_dir, exist_ok=True)

            with open(out, "wb") as f:
                writer.write(f)

            QMessageBox.information(self, "Sucesso", f"PDF gerado com sucesso:\n{out}")

        except Exception as e:
            QMessageBox.critical(self, "Erro ao salvar", f"Ocorreu um erro:\n{e}")


if __name__ == "__main__":
    # Ajuda o Windows a usar o ícone correto na barra de tarefas
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "com.rick.pdfmerger.editor.1"
        )
    except Exception:
        pass

    app = QApplication(sys.argv)

    # pdf_ico.ico está na mesma pasta do .py (no projeto). No exe, inclua via --add-data.
    icon_path = resource_path("pdf_ico.ico")
    icon = QIcon(icon_path)

    # Ícone global (barra de tarefas / alt-tab / diálogos)
    app.setWindowIcon(icon)

    w = PdfMergerApp(icon)
    w.show()
    sys.exit(app.exec())