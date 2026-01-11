import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from automacao_contabil import run_automation


def build_ui():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    app = ctk.CTk()
    app.title("Automação Contábil")
    app.geometry("800x480")
    app.minsize(480, 320)
    app.grid_rowconfigure(0, weight=1)
    app.grid_columnconfigure(0, weight=1)
    # Inicia em tela cheia e configura atalhos para alternar (F11) e sair (Escape)
    app.attributes("-fullscreen", True)
    fullscreen = {"state": True}

    def toggle_fullscreen(event=None):
        fullscreen["state"] = not fullscreen["state"]
        app.attributes("-fullscreen", fullscreen["state"])

    def exit_fullscreen(event=None):
        if fullscreen["state"]:
            fullscreen["state"] = False
            app.attributes("-fullscreen", False)

    app.bind("<F11>", toggle_fullscreen)
    app.bind("<Escape>", exit_fullscreen)

    ui_user = ctk.StringVar()
    ui_pass = ctk.StringVar()
    selected = {"path": None}

    login_frame = ctk.CTkFrame(app)
    file_frame = ctk.CTkFrame(app)
    # Centraliza conteúdo criando 3 colunas: [espaço][conteúdo][espaço]
    login_frame.grid(row=0, column=0, sticky="nsew")

    for frame in (login_frame, file_frame):
        # configura linhas para comportamento previsível ao redimensionar
        # agora temos linhas 0..11, com 0 e 11 como espaçadores para centralizar verticalmente
        for i in range(0, 12):
            frame.grid_rowconfigure(i, weight=0)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_rowconfigure(11, weight=1)
        # colunas: 0 e 2 flexíveis (preenchem espaço), 1 conteúdo centralizado
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=0)
        frame.grid_columnconfigure(2, weight=1)

    # calcula tamanhos iniciais responsivos com base na largura da tela
    screen_w = app.winfo_screenwidth()
    entry_width = max(360, int(screen_w * 0.35))
    file_wrap = int(screen_w * 0.5)

    # Login layout (responsive)
    # --- Login (coluna central = 1) ---
    # deslocamos tudo uma linha para baixo para criar espaçador superior
    ctk.CTkLabel(login_frame, text="Acesso", font=ctk.CTkFont(size=20, weight="bold")).grid(row=1, column=1, pady=(20, 10))
    ctk.CTkLabel(login_frame, text="E-mail").grid(row=2, column=1, pady=(8, 0))
    ctk.CTkEntry(login_frame, textvariable=ui_user, width=entry_width).grid(row=3, column=1, pady=6)
    ctk.CTkLabel(login_frame, text="Senha").grid(row=4, column=1, pady=(8, 0))
    ctk.CTkEntry(login_frame, textvariable=ui_pass, show="*", width=entry_width).grid(row=5, column=1, pady=6)

    def to_file_screen():
        login_frame.grid_forget()
        file_frame.grid(row=0, column=0, sticky="nsew")

    # botão centralizado (uma linha abaixo dos campos)
    ctk.CTkButton(login_frame, text="Avançar", command=to_file_screen).grid(row=6, column=1, pady=18)

    # File selection layout (responsive)
    # --- Seleção de arquivo (coluna central = 1) ---
    ctk.CTkLabel(file_frame, text="Selecionar planilha de lançamentos", font=ctk.CTkFont(size=16, weight="bold")).grid(row=1, column=1, pady=(20, 10))
    file_label = ctk.CTkLabel(file_frame, text="Nenhum arquivo selecionado", wraplength=file_wrap)
    file_label.grid(row=2, column=1, pady=(6, 10))

    def choose_file():
        path = filedialog.askopenfilename(title="Selecione a planilha de lançamentos", filetypes=[("Excel files", "*.xlsx")])
        if path:
            selected["path"] = path
            file_label.configure(text=path)

    ctk.CTkButton(file_frame, text="Selecionar arquivo", command=choose_file).grid(row=3, column=1, pady=(4, 8))

    # status e botão centralizados (com espaçador superior)
    status_label = ctk.CTkLabel(file_frame, text="Aguardando...")
    status_label.grid(row=4, column=1, pady=(6, 6))

    start_btn = ctk.CTkButton(file_frame, text="Iniciar Automação")
    start_btn.grid(row=5, column=1, pady=(10, 6))

    def start_action():
        path = selected.get("path")
        if not path:
            messagebox.showwarning("Aviso", "Selecione um arquivo .xlsx antes de iniciar.")
            return

        def target():
            try:
                run_automation(path)
                app.after(0, lambda: messagebox.showinfo("Sucesso", "Relatório incluído com sucesso."))
            except Exception as e:
                app.after(0, lambda: messagebox.showerror("Erro", f"Erro na automação: {e}"))
            finally:
                app.after(0, lambda: (start_btn.configure(state="normal"), status_label.configure(text="Concluído")))

        start_btn.configure(state="disabled")
        status_label.configure(text="Executando...")
        thread = threading.Thread(target=target, daemon=True)
        thread.start()

    start_btn.configure(command=start_action)

    app.mainloop()


if __name__ == "__main__":
    build_ui()
