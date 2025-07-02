"""
Interface gráfica para controle da automação do SSW.
Este módulo implementa uma interface gráfica usando customtkinter para gerenciar
a automação de extração de relatórios do sistema SSW.

Principais funcionalidades:
- Interface gráfica amigável com tema escuro
- Sistema de agendamento de execuções
- Sistema de log em tempo real
- Minimização para bandeja do sistema (system tray)
- Controle de execução manual e agendada
"""

import customtkinter as ctk
import threading
import time
import json
import os
from datetime import datetime
import schedule
import sys
import pystray
from pystray import MenuItem as item
from PIL import Image

# Classe para redirecionar a saída do console para o textbox da interface
class TextboxRedirector:
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, text):
        self.textbox.log(text, end="")

    def flush(self):
        pass

# Tenta importar a função principal de automação. Se falhar, cria uma função mock para testes
try:
    from auto_455 import main as automacao_main
except ImportError:
    def automacao_main(stop_event=None):
        print("ERRO: O arquivo 'auto_455.py' não foi encontrado.")
        time.sleep(5)

# Janela de configuração dos agendamentos
class ScheduleWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.parent = parent
        self.schedule_file = "agendamentos.json"
        
        self.title("Configurações de Agendamento")
        self.geometry("500x400")
        self.resizable(False, False)
        
        self.transient(parent)
        self.grab_set()
        
        self.setup_ui()
        self.load_schedules()
        self.update_schedule_list()
        
    def setup_ui(self):
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        title_label = ctk.CTkLabel(main_frame, text="Gerenciar Agendamentos", 
                                 font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(pady=(10, 20))
        
        add_frame = ctk.CTkFrame(main_frame)
        add_frame.pack(fill="x", padx=10, pady=(0, 20))
        
        add_label = ctk.CTkLabel(add_frame, text="Adicionar Novo Horário:")
        add_label.pack(pady=(10, 5))
        
        input_frame = ctk.CTkFrame(add_frame, fg_color="transparent")
        input_frame.pack(pady=(0, 10))
        
        self.time_entry = ctk.CTkEntry(input_frame, placeholder_text="HH:MM", width=120)
        self.time_entry.pack(side="left", padx=(0, 10))
        
        add_button = ctk.CTkButton(input_frame, text="Adicionar", 
                                 command=self.add_schedule, width=100)
        add_button.pack(side="left")
        
        list_frame = ctk.CTkFrame(main_frame)
        list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 20))
        
        list_label = ctk.CTkLabel(list_frame, text="Agendamentos Ativos:")
        list_label.pack(pady=(10, 5))
        
        self.schedule_scroll = ctk.CTkScrollableFrame(list_frame, height=150)
        self.schedule_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        control_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        control_frame.pack(fill="x", padx=10)
        
        close_button = ctk.CTkButton(control_frame, text="Fechar", 
                                   command=self.destroy, width=100)
        close_button.pack(side="right", padx=(10, 0))
        
        clear_button = ctk.CTkButton(control_frame, text="Limpar Todos", 
                                   command=self.clear_all_schedules, width=100,
                                   hover_color="darkred")
        clear_button.pack(side="right")
    
    def load_schedules(self):
        try:
            if os.path.exists(self.schedule_file):
                with open(self.schedule_file, 'r', encoding='utf-8') as f:
                    self.schedules = json.load(f)
            else:
                self.schedules = []
        except Exception as e:
            self.schedules = []
            self.parent.log(f"Erro ao carregar agendamentos: {e}")
    
    def save_schedules(self):
        try:
            with open(self.schedule_file, 'w', encoding='utf-8') as f:
                json.dump(self.schedules, f, indent=2, ensure_ascii=False)
            self.parent.log(f"Agendamentos salvos: {len(self.schedules)} horários ativos")
        except Exception as e:
            self.parent.log(f"Erro ao salvar agendamentos: {e}")
    
    def add_schedule(self):
        time_str = self.time_entry.get().strip()
        
        if not time_str:
            self.show_error("Por favor, insira um horário")
            return
            
        try:
            datetime.strptime(time_str, "%H:%M")
        except ValueError:
            self.show_error("Formato inválido. Use HH:MM (ex: 09:30)")
            return
        
        if time_str in self.schedules:
            self.show_error("Este horário já está agendado")
            return
        
        self.schedules.append(time_str)
        self.schedules.sort()
        
        self.save_schedules()
        self.update_schedule_list()
        self.parent.update_schedules()
        
        self.time_entry.delete(0, "end")
        
        self.parent.log(f"Agendamento adicionado: {time_str}")
    
    def remove_schedule(self, time_str):
        if time_str in self.schedules:
            self.schedules.remove(time_str)
            self.save_schedules()
            self.update_schedule_list()
            self.parent.update_schedules()
            self.parent.log(f"Agendamento removido: {time_str}")
    
    def clear_all_schedules(self):
        if not self.schedules:
            self.show_error("Não há agendamentos para remover")
            return
            
        dialog = ctk.CTkInputDialog(text="Digite 'CONFIRMAR' para remover todos os agendamentos:",
                                  title="Confirmação")
        if dialog.get_input() == "CONFIRMAR":
            self.schedules.clear()
            self.save_schedules()
            self.update_schedule_list()
            self.parent.update_schedules()
            self.parent.log("Todos os agendamentos foram removidos")
    
    def update_schedule_list(self):
        for widget in self.schedule_scroll.winfo_children():
            widget.destroy()
        
        if not self.schedules:
            no_schedule_label = ctk.CTkLabel(self.schedule_scroll, 
                                           text="Nenhum agendamento configurado",
                                           text_color="gray")
            no_schedule_label.pack(pady=20)
            return
        
        for i, time_str in enumerate(self.schedules):
            schedule_frame = ctk.CTkFrame(self.schedule_scroll)
            schedule_frame.pack(fill="x", padx=5, pady=2)
            
            time_label = ctk.CTkLabel(schedule_frame, text=f"🕐 {time_str}", 
                                    font=ctk.CTkFont(size=14))
            time_label.pack(side="left", padx=10, pady=5)
            
            remove_button = ctk.CTkButton(schedule_frame, text="Remover", 
                                        command=lambda t=time_str: self.remove_schedule(t),
                                        width=80, height=25, 
                                        fg_color="darkred", hover_color="red")
            remove_button.pack(side="right", padx=10, pady=5)
    
    def show_error(self, message):
        error_dialog = ctk.CTkInputDialog(text=f"ERRO: {message}\n\nPressione OK para continuar",
                                        title="Erro")
        error_dialog.get_input()

# Classe principal da aplicação
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Variáveis de controle da automação e agendamento
        self.automation_thread = None  # Thread que executa a automação
        self.stop_event = None        # Evento para controle de parada da automação
        self.scheduler_thread = None   # Thread que gerencia os agendamentos
        self.is_schedule_running = False  # Estado do agendador
        self.schedule_window = None    # Referência para janela de agendamentos

        # Configuração da janela principal
        self.title("Automação 455")
        self.geometry("600x600")
        self.grid_columnconfigure(0, weight=1)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        main_frame = ctk.CTkFrame(self, corner_radius=15)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(main_frame, text="Automação 455 - SSW", 
                                 font=ctk.CTkFont(size=24, weight="bold"))
        title_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        action_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        action_frame.grid(row=1, column=0, pady=10)
        action_frame.grid_columnconfigure((0, 1), weight=1)

        self.start_button = ctk.CTkButton(action_frame, text="Iniciar Automação", 
                                        command=self.start_automation)
        self.start_button.grid(row=0, column=0, padx=5)

        self.stop_button = ctk.CTkButton(action_frame, text="Parar Automação", 
                                       command=self.stop_automation, state="disabled")
        self.stop_button.grid(row=0, column=1, padx=5)

        schedule_frame = ctk.CTkFrame(main_frame)
        schedule_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        schedule_frame.grid_columnconfigure(0, weight=1)

        schedule_label = ctk.CTkLabel(schedule_frame, text="Sistema de Agendamento", 
                                    font=ctk.CTkFont(size=16, weight="bold"))
        schedule_label.grid(row=0, column=0, padx=10, pady=(10, 5))
        
        self.schedule_status_label = ctk.CTkLabel(schedule_frame, text="Nenhum agendamento ativo", 
                                                text_color="gray")
        self.schedule_status_label.grid(row=1, column=0, padx=10, pady=5)

        schedule_buttons_frame = ctk.CTkFrame(schedule_frame, fg_color="transparent")
        schedule_buttons_frame.grid(row=2, column=0, pady=10)
        
        self.manage_button = ctk.CTkButton(schedule_buttons_frame, text="Gerenciar Agendamentos", 
                                         command=self.open_schedule_window, width=150)
        self.manage_button.grid(row=0, column=0, padx=5)

        self.toggle_schedule_button = ctk.CTkButton(schedule_buttons_frame, text="Iniciar Agendador", 
                                                  command=self.toggle_scheduler, width=150)
        self.toggle_schedule_button.grid(row=0, column=1, padx=5)
        
        self.status_textbox = ctk.CTkTextbox(main_frame, height=280, 
                                           font=ctk.CTkFont(family="Courier New", size=12))
        self.status_textbox.grid(row=3, column=0, padx=20, pady=20, sticky="nsew")

        sys.stdout = TextboxRedirector(self)

        self.log("Painel de controle iniciado. Pronto para receber comandos.")
        
        self.update_schedules()
    
    def log(self, message, end="\n"):
        """Adiciona uma mensagem ao log com timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}{end}" if not message.startswith("[") else f"{message}{end}"
        
        self.status_textbox.insert("end", formatted_message)
        self.status_textbox.see("end")
        self.update_idletasks()

    def update_button_states(self, is_running=False):
        """Atualiza o estado dos botões baseado no estado da automação"""
        self.start_button.configure(state="disabled" if is_running else "normal")
        self.stop_button.configure(state="normal" if is_running else "disabled")
        self.manage_button.configure(state="disabled" if is_running else "normal")

    def open_schedule_window(self):
        if self.schedule_window is None or not self.schedule_window.winfo_exists():
            self.schedule_window = ScheduleWindow(self)
        else:
            self.schedule_window.focus()

    def load_schedules_from_file(self):
        try:
            if os.path.exists("agendamentos.json"):
                with open("agendamentos.json", 'r', encoding='utf-8') as f:
                    return json.load(f)
            return []
        except Exception as e:
            self.log(f"Erro ao carregar agendamentos: {e}")
            return []

    def update_schedules(self):
        schedules = self.load_schedules_from_file()
        
        schedule.clear()
        
        if schedules:
            for time_str in schedules:
                schedule.every().day.at(time_str).do(self.start_automation)
            
            schedule_text = f"{len(schedules)} agendamento{'s' if len(schedules) > 1 else ''} ativo{'s' if len(schedules) > 1 else ''}: {', '.join(schedules)}"
            self.schedule_status_label.configure(text=schedule_text, text_color="lightgreen")
            
            if not self.is_schedule_running and schedules:
                self.start_scheduler()
        else:
            self.schedule_status_label.configure(text="Nenhum agendamento ativo", text_color="gray")
            if self.is_schedule_running:
                self.stop_scheduler()

    def toggle_scheduler(self):
        if self.is_schedule_running:
            self.stop_scheduler()
        else:
            schedules = self.load_schedules_from_file()
            if schedules:
                self.start_scheduler()
            else:
                self.log("Nenhum agendamento configurado. Configure pelo menos um horário primeiro.")

    def start_scheduler(self):
        if not self.is_schedule_running:
            schedules = self.load_schedules_from_file()
            if not schedules:
                self.log("Nenhum agendamento configurado. Configure pelo menos um horário primeiro.")
                return
                
            schedule.clear()
            for time_str in schedules:
                schedule.every().day.at(time_str).do(self.start_automation)
            
            self.is_schedule_running = True
            self.scheduler_thread = threading.Thread(target=self._scheduler_worker)
            self.scheduler_thread.daemon = True
            self.scheduler_thread.start()
            
            self.toggle_schedule_button.configure(text="Parar Agendador", fg_color="darkred", hover_color="red")
            self.log(f"Sistema de agendamento iniciado com {len(schedules)} horário(s).")

    def stop_scheduler(self):
        self.is_schedule_running = False
        schedule.clear()
        self.toggle_schedule_button.configure(text="Iniciar Agendador", fg_color=self.start_button.cget("fg_color"), hover_color=self.start_button.cget("hover_color"))
        self.log("Sistema de agendamento parado e agendamentos limpos.")

    def start_automation(self):
        """Inicia a execução da automação em uma thread separada"""
        if self.automation_thread and self.automation_thread.is_alive():
            self.log("A automação já está em execução.")
            return

        self.log("="*50)
        self.log("Iniciando a automação...")
        self.update_button_states(is_running=True)
        
        self.stop_event = threading.Event()
        self.automation_thread = threading.Thread(target=self._automation_worker)
        self.automation_thread.daemon = True
        self.automation_thread.start()

    def stop_automation(self):
        """Envia sinal para parar a automação de forma segura"""
        if self.stop_event:
            self.log("Sinal de parada enviado. Aguardando finalização da tarefa atual...")
            self.stop_event.set()
            self.stop_button.configure(state="disabled")

    def _automation_worker(self):
        try:
            automacao_main(self.stop_event)
            if self.stop_event.is_set():
                self.log("Automação interrompida com sucesso.")
            else:
                self.log("Automação concluída com sucesso!")
        except Exception as e:
            self.log(f"ERRO CRÍTICO NA AUTOMAÇÃO:\n{e}")
        finally:
            self.log("="*50)
            self.update_button_states(is_running=False)
            if self.is_schedule_running:
                self.log("Aguardando próximo agendamento...")

    def _scheduler_worker(self):
        while self.is_schedule_running:
            try:
                schedule.run_pending()
                time.sleep(1)
            except Exception as e:
                self.log(f"Erro no agendador: {e}")
                time.sleep(1)
        self.log("Thread do agendador finalizada.")

    def graceful_shutdown(self):
        self.log("Encerrando a aplicação...")
        if self.is_schedule_running:
            self.stop_scheduler()
        if self.automation_thread and self.automation_thread.is_alive():
            self.stop_automation()
        self.destroy()

if __name__ == "__main__":
    # Inicialização da aplicação e configuração do ícone na bandeja
    app = App()
    icon = None

    def mostrar_janela():
        """Restaura a janela principal quando clicado no ícone da bandeja"""
        app.after(0, app.deiconify)

    def esconder_janela():
        """Minimiza a janela principal para a bandeja"""
        app.withdraw()

    def sair_do_app():
        """Encerra a aplicação de forma segura"""
        if icon:
            icon.stop()
        app.graceful_shutdown()

    def setup_tray():
        global icon
        try:
            image = Image.open("icon.png")
        except FileNotFoundError:
            image = Image.new('RGB', (64, 64), 'blue')
        
        menu = (item('Mostrar Painel', mostrar_janela), item('Sair', sair_do_app))
        icon = pystray.Icon("automacao_455", image, "Automação 455", menu)
        icon.run()

    app.attributes('-toolwindow', True)
    app.protocol("WM_DELETE_WINDOW", esconder_janela)

    tray_thread = threading.Thread(target=setup_tray, daemon=True)
    tray_thread.start()

    app.withdraw()
    app.mainloop()