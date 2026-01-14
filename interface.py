import customtkinter as ctk
import threading
from PIL import Image
import os
import sys

def recurso_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, para o PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# Importamos a função de processamento
try:
    from robo_notas import processar_arquivos
except ImportError:
    print("Erro: O arquivo 'robo_notas.py' não foi encontrado.")

class AppAutomacao(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- CONFIGURAÇÃO DA JANELA ---
        self.title("Sistema de Gestão de Notas - Desenvolvedor Gabriel Sonvezzo")
        self.geometry("620x700") # Altura ideal para caber tudo
        self.resizable(False, False) 
        ctk.set_appearance_mode("dark") 
        
        self.azul_perfipar = "#0044ff"
        self.azul_hover = "#0033cc"

        # --- LAYOUT OTIMIZADO ---
        
        # 1. Logo Centralizada (Espaço reduzido)
        try:
            diretorio_atual = os.path.dirname(os.path.abspath(__file__))
            caminho_logo = os.path.join(diretorio_atual, "logo.png")
            
            img_original = Image.open(caminho_logo)
            largura_orig, altura_orig = img_original.size
            
            # Ajustamos a largura para 250px e calculamos a altura proporcional
            largura_exibicao = 220 
            proporcao = largura_exibicao / largura_orig
            altura_exibicao = int(altura_orig * proporcao)

            self.img_logo = ctk.CTkImage(
                light_image=img_original, 
                dark_image=img_original, 
                size=(largura_exibicao, altura_exibicao)
            )
            
            # pady=(10, 5) remove o excesso de "branco" (espaço vazio) em cima e embaixo
            self.label_logo = ctk.CTkLabel(self, image=self.img_logo, text="")
            self.label_logo.pack(pady=(5, 2), anchor="center") 
        except Exception as e:
            print(f"Erro ao carregar logo: {e}")

        # 2. Títulos
        self.label_titulo = ctk.CTkLabel(self, text="AUTOMAÇÃO DE ENTRADA MP", font=("Arial", 20, "bold"))
        self.label_titulo.pack(pady=(5, 0))

        self.label_subtitulo = ctk.CTkLabel(self, text="Desenvolvedor de Sistemas: Gabriel", font=("Arial", 11), text_color="gray")
        self.label_subtitulo.pack(pady=(0, 10))

        # 3. Status e Barra de Progresso
        self.label_status = ctk.CTkLabel(self, text="Status: Pronto para iniciar", font=("Arial", 13, "italic"))
        self.label_status.pack(pady=2)

        self.barra = ctk.CTkProgressBar(self, width=500, height=12, mode="determinate", progress_color=self.azul_perfipar)
        self.barra.set(0)
        self.barra.pack(pady=5)

        # 4. Caixa de Logs (Reduzida para 240px de altura para o botão subir)
        self.caixa_log = ctk.CTkTextbox(self, width=540, height=240, font=("Consolas", 12), corner_radius=10, border_width=1, border_color="#333")
        self.caixa_log.pack(pady=10, padx=20)

        # 5. Botão de Execução (Agora visível!)
        self.botao = ctk.CTkButton(self, text="LANÇAR NOTAS NA PLANILHA", 
                                   command=self.iniciar_thread,
                                   fg_color=self.azul_perfipar, 
                                   hover_color=self.azul_hover,
                                   font=("Arial", 15, "bold"),
                                   height=45, width=400)
        self.botao.pack(pady=(5, 15))

    def atualizar_tela(self, mensagem):
        self.caixa_log.insert("end", f"{mensagem}\n")
        self.caixa_log.see("end")

    def iniciar_thread(self):
        self.botao.configure(state="disabled", text="PROCESSANDO...")
        self.caixa_log.delete("1.0", "end")
        self.barra.configure(mode="indeterminate")
        self.barra.start()
        self.label_status.configure(text="Status: Lendo arquivos e cruzando lotes...")
        
        thread = threading.Thread(target=self.executar_motor, daemon=True)
        thread.start()

    def executar_motor(self):
        try:
            processar_arquivos(callback=self.atualizar_tela)
            self.label_status.configure(text="Status: Concluído!")
            self.barra.configure(mode="determinate")
            self.barra.set(1)
        except Exception as e:
            self.atualizar_tela(f"❌ ERRO: {e}")
            self.label_status.configure(text="Status: Erro detectado")
            self.barra.set(0)
        
        self.barra.stop()
        self.botao.configure(state="normal", text="LANÇAR NOTAS NA PLANILHA")

if __name__ == "__main__":
    app = AppAutomacao()
    app.mainloop()