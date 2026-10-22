"""Conversor PDF -> OFX / Excel (Linux e Windows)."""
from __future__ import annotations

import os
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

from core.exporters import to_excel, to_ofx
from core.parser import Statement, Transaction, parse_pdf

ACCENT, ACCENT_HOVER = "#3B5BDB", "#2F4AB8"
CARD = ("#FFFFFF", "#1B2236")
BG = ("#EEF2FA", "#10151F")


def open_folder(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("dark")
        self.title("Conversor PDF → OFX & Excel")
        self.geometry("920x720")
        self.minsize(820, 680)
        self.configure(fg_color=BG)
        ico = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "assets" / "icon.ico"
        if sys.platform == "win32" and ico.exists():
            self.after(200, lambda: self.iconbitmap(str(ico)))
        self.files: list[Path] = []
        self.out_dir: Path | None = None
        F = lambda size, bold=False: ctk.CTkFont(size=size, weight="bold" if bold else "normal")
        muted = ("#6B7690", "#8E9ABA")

        # cabeçalho
        header = ctk.CTkFrame(self, fg_color=("#1F2A44", "#141B2D"), corner_radius=0)
        header.pack(fill="x")
        hin = ctk.CTkFrame(header, fg_color="transparent")
        hin.pack(fill="x", padx=36, pady=26)
        logo = ctk.CTkLabel(hin, text="⇄", width=64, height=64, corner_radius=18,
                            fg_color=ACCENT, text_color="white", font=F(32, True))
        logo.pack(side="left")
        tbox = ctk.CTkFrame(hin, fg_color="transparent")
        tbox.pack(side="left", padx=18)
        ctk.CTkLabel(tbox, text="Conversor de Extratos", font=F(28, True),
                     text_color="white", anchor="w").pack(anchor="w")
        ctk.CTkLabel(tbox, text="PDF  →  OFX  &  Excel formatado, em segundos", font=F(14),
                     text_color="#AEB9D6", anchor="w").pack(anchor="w", pady=(2, 0))
        self.theme_btn = ctk.CTkSegmentedButton(hin, values=["Claro", "Escuro"], command=self.set_theme,
                                                selected_color=ACCENT, selected_hover_color=ACCENT_HOVER)
        self.theme_btn.set("Escuro")
        self.theme_btn.pack(side="right")

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=36, pady=(22, 26))

        # área de arquivos
        files_card = ctk.CTkFrame(body, fg_color=CARD, corner_radius=20, border_width=1,
                                  border_color=("#DCE3F2", "#2A3350"))
        files_card.pack(fill="both", expand=True)
        top = ctk.CTkFrame(files_card, fg_color="transparent")
        top.pack(fill="x", padx=24, pady=(20, 8))
        ctk.CTkLabel(top, text="1  ·  Seus arquivos", font=F(17, True)).pack(side="left")
        self.count = ctk.CTkLabel(top, text="Nenhum arquivo", text_color=muted, font=F(13))
        self.count.pack(side="right")

        self.listbox = ctk.CTkScrollableFrame(files_card, fg_color=("#F4F7FD", "#141A2B"), corner_radius=14)
        self.listbox.pack(fill="both", expand=True, padx=24, pady=6)
        self.empty = ctk.CTkLabel(self.listbox, text="📄\n\nNenhum PDF adicionado ainda\nClique em “Adicionar PDFs” para começar",
                                  text_color=muted, font=F(14), justify="center")

        btns = ctk.CTkFrame(files_card, fg_color="transparent")
        btns.pack(fill="x", padx=24, pady=(8, 20))
        ctk.CTkButton(btns, text="＋  Adicionar PDFs", height=44, corner_radius=12, font=F(14, True),
                      fg_color=ACCENT, hover_color=ACCENT_HOVER, command=self.add_files).pack(side="left")
        ctk.CTkButton(btns, text="Limpar lista", height=44, corner_radius=12, font=F(14),
                      fg_color="transparent", border_width=1, border_color=muted,
                      text_color=("#333", "#DDD"), hover_color=("#E6EBF7", "#232C45"),
                      command=self.clear).pack(side="left", padx=10)

        # opções
        opt_card = ctk.CTkFrame(body, fg_color=CARD, corner_radius=20, border_width=1,
                                border_color=("#DCE3F2", "#2A3350"))
        opt_card.pack(fill="x", pady=(16, 0))
        ctk.CTkLabel(opt_card, text="2  ·  Formato e destino", font=F(17, True)).pack(anchor="w", padx=24, pady=(18, 8))
        row = ctk.CTkFrame(opt_card, fg_color="transparent")
        row.pack(fill="x", padx=24)
        self.var_ofx = ctk.BooleanVar(value=True)
        self.var_xlsx = ctk.BooleanVar(value=True)
        self.var_merge = ctk.BooleanVar(value=False)
        for text, var in (("OFX  (contabilidade)", self.var_ofx), ("Excel  (.xlsx)", self.var_xlsx),
                          ("Unir em um só arquivo", self.var_merge)):
            ctk.CTkSwitch(row, text=text, variable=var, font=F(13), progress_color=ACCENT).pack(side="left", padx=(0, 26))
        out = ctk.CTkFrame(opt_card, fg_color="transparent")
        out.pack(fill="x", padx=24, pady=(14, 18))
        self.out_label = ctk.CTkLabel(out, text="📁  Salvar ao lado dos PDFs", text_color=muted, anchor="w", font=F(13))
        self.out_label.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(out, text="Escolher pasta", width=140, height=36, corner_radius=10,
                      fg_color="transparent", border_width=1, border_color=muted,
                      text_color=("#333", "#DDD"), hover_color=("#E6EBF7", "#232C45"),
                      command=self.pick_out).pack(side="right")

        # ação
        self.progress = ctk.CTkProgressBar(body, progress_color=ACCENT, height=8, corner_radius=4)
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(18, 6))
        self.status = ctk.CTkLabel(body, text="Pronto para converter.", text_color=muted, anchor="w", font=F(13))
        self.status.pack(fill="x")
        self.go = ctk.CTkButton(body, text="⚡  Converter agora", height=56, corner_radius=16,
                                font=F(18, True), fg_color=ACCENT, hover_color=ACCENT_HOVER,
                                command=self.start)
        self.go.pack(fill="x", pady=(12, 0))
        self.refresh()

    # ---- ações
    def set_theme(self, v: str) -> None:
        ctk.set_appearance_mode("light" if v == "Claro" else "dark")

    def add_files(self) -> None:
        for f in filedialog.askopenfilenames(title="Selecione os PDFs", filetypes=[("PDF", "*.pdf")]):
            p = Path(f)
            if p not in self.files:
                self.files.append(p)
        self.refresh()

    def clear(self) -> None:
        self.files.clear()
        self.refresh()

    def pick_out(self) -> None:
        d = filedialog.askdirectory()
        if d:
            self.out_dir = Path(d)
            self.out_label.configure(text=f"📁  {self.out_dir}")

    def refresh(self) -> None:
        for w in self.listbox.winfo_children():
            w.destroy()
        if not self.files:
            self.empty = ctk.CTkLabel(self.listbox, text="📄\n\nNenhum PDF adicionado ainda\nClique em “Adicionar PDFs” para começar",
                                      text_color=("#6B7690", "#8E9ABA"), justify="center", font=ctk.CTkFont(size=14))
            self.empty.pack(pady=40, expand=True)
        for p in self.files:
            chip = ctk.CTkFrame(self.listbox, fg_color=("#FFFFFF", "#1E2740"), corner_radius=10)
            chip.pack(fill="x", pady=3, padx=2)
            ctk.CTkLabel(chip, text=f"📄  {p.name}", anchor="w", font=ctk.CTkFont(size=13)).pack(side="left", padx=12, pady=8)
            ctk.CTkButton(chip, text="✕", width=30, height=28, corner_radius=8, fg_color="transparent",
                          text_color=("#C0392B", "#FF7B72"), hover_color=("#FBE3E0", "#3A2430"),
                          command=lambda x=p: self.remove(x)).pack(side="right", padx=8)
        self.count.configure(text=f"{len(self.files)} arquivo(s)" if self.files else "Nenhum arquivo")

    def remove(self, p: Path) -> None:
        self.files.remove(p)
        self.refresh()

    def say(self, text: str, progress: float | None = None) -> None:
        def _u() -> None:
            self.status.configure(text=text)
            if progress is not None:
                self.progress.set(progress)
        self.after(0, _u)

    def start(self) -> None:
        if not self.files:
            self.status.configure(text="Adicione ao menos um PDF.")
            return
        if not (self.var_ofx.get() or self.var_xlsx.get()):
            self.status.configure(text="Marque OFX e/ou Excel.")
            return
        self.go.configure(state="disabled")
        threading.Thread(target=self.run, daemon=True).start()

    def run(self) -> None:
        errors: list[str] = []
        try:
            merged = Statement()
            n = len(self.files)
            last_dir: Path = self.out_dir or self.files[0].parent
            for i, pdf in enumerate(self.files):
                self.say(f"Lendo {pdf.name}…", i / n)
                try:
                    st = parse_pdf(pdf)
                    if not st.transactions:
                        errors.append(f"{pdf.name}: nenhum lançamento reconhecido")
                        continue
                    if self.var_merge.get():
                        merged.transactions += st.transactions
                        merged.bank, merged.bank_id = merged.bank or st.bank, (
                            merged.bank_id if merged.bank else st.bank_id)
                        continue
                    last_dir = self.out_dir or pdf.parent
                    self.export(st, last_dir / pdf.stem)
                except Exception as e:  # noqa: BLE001  – um arquivo ruim não derruba os demais
                    errors.append(f"{pdf.name}: {e}")
            if self.var_merge.get() and merged.transactions:
                self.export(merged, last_dir / "extratos_unificados")
            msg = f"Concluído! Arquivos salvos em {last_dir}"
            if errors:
                msg += f"  ⚠ {len(errors)} com problema: " + " | ".join(errors)
            self.say(msg, 1.0)
            self.after(0, lambda: open_folder(last_dir))
        except Exception as e:  # noqa: BLE001
            self.say(f"Erro: {e}")
        finally:
            self.after(0, lambda: self.go.configure(state="normal"))

    def export(self, st: Statement, base: Path) -> None:
        # with_name (e não with_suffix): nomes como "extrato 01.10.2025" perderiam o final
        if self.var_ofx.get():
            to_ofx(st.transactions, base.with_name(base.name + ".ofx"), st.account, st.bank_id)
        if self.var_xlsx.get():
            to_excel(st.transactions, base.with_name(base.name + ".xlsx"))


if __name__ == "__main__":
    App().mainloop()
