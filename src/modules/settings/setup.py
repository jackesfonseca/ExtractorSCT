import tkinter.messagebox
import tkinter.filedialog
import datetime as dt

# estrutura hora atual para apresentar na tela
def get_actual_hour():
    hora_atual = dt.datetime.now()
    return '[' + hora_atual.strftime('%H:%M:%S') + '] '

def get_actual_date_hour():
    data_atual = dt.date.today()
    hora_atual = dt.datetime.now()
    return '[' + data_atual.strftime('%d/%m/%Y') + ' ' + hora_atual.strftime('%H:%M:%S') + '] '
    
def clean_logs(listbox_log):
    ans = tkinter.messagebox.askquestion('Atenção', 'Deseja limpar a caixa de log?',icon='warning')

    if ans == 'yes':
        listbox_log.delete(0, tkinter.END)
    
    else:
        return
