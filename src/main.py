import tkinter as tkk
import time
from PIL import Image, ImageTk
import tkinter.messagebox
import tkinter.filedialog
from modules.samp import download_fornecimento as df
from modules.samp import download_receita as dr
from modules.samp import folders
from modules.samp import extract_cods
from modules.settings import setup
import os

''' Apresenta informações sobre o software
@param: None
@return: None
'''
def show_about():
    # cria nova janela
    win = tkinter.Toplevel()
    win.title('REF - Sobre')

    # seta configurações de largura e altura da tela da aplicação
    width = 960
    height = 650

    # obtém informações de largura e altura do monitor
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()

    # calcula posição x e y central do monitor 
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)

    win.geometry('%dx%d+%d+%d' % (width, height, x, y))

    # seta informações e icones a serem apresentados
    sobre = ('Software desenvolvido para realizações de tarefas automatizadas para o setor da Regulação Econômica e Financeira\n\n' +
             '1. Mercado\n' +
             '  a) Fornecimento Energia\n' +
             '      - Baixar\n' +
             '  b) Receita Uso\n' +
             '      - Baixar\n' +
             '  c) Extrair códigos SAMP (SIMULADOR)\n\n' +
             '2. Extrair\n' +
             '  a) Códigos SAMP\n' +
             '      - Arquivo\n' +
             '      - Pasta\n')

    tkinter.Label(win, text=sobre, width=100, height=30, font=('Lato', 12), anchor='w', justify='left').pack()

    tkinter.Button(win, text='Lido', font=('Lato', 12),command=win.destroy).pack(pady=10)

    tkinter.Label(win, text='Internal Use', height=10, font=('Lato', 12), anchor='w', justify='center').pack(pady=10)

''' Apresenta orientações gerais de uso do software
@param: None
@return: None
'''
def show_orientations():
    # cria nova janela
    win = tkinter.Toplevel()
    win.title('REF - Orientações')

    # seta configurações de largura e altura da tela da aplicação
    width = 990
    height = 650

    # obtém informações de largura e altura do monitor
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()

    # calcula posição x e y central do monitor
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)

    win.geometry('%dx%d+%d+%d' % (width, height, x, y))

    # seta informações e icones a serem apresentados
    orientacoes = ('1 - Mercado\n' +
                   '    - Todos os arquivos precisam estar na mesma pasta\n')

    tkinter.Label(win, text='\nHINTS\n', font=('Lato', 14), anchor='w', justify='center').pack()

    tkinter.Label(win, text=orientacoes, width=100, height=15, font=('Lato', 12), anchor='w', justify='left').pack(pady=20)

    tkinter.Button(win, text='Lido', font=('Lato', 12),command=win.destroy).pack(pady=20)

    tkinter.Label(win, text='Internal Use', height=5, font=('Lato', 12), anchor='w', justify='center').pack(pady=20)

''' Inicia UI principal
@param: None
@return: None
'''
def main_GUI():

    # cria objeto Tkinter que referencia tela principal
    root = tkk.Tk()

    # adiciona informações e seta configurações 
    root.title("SCT - Extrator")

    width = 590
    height = 650

    # obtém informações de largura e altura do monitor
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # calcula posição x e y central do monitor 
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)

    root.geometry('%dx%d+%d+%d' % (width, height, x, y))

    root.resizable(False,False)

    # adiciona logo da aplicação
    # abs_path = os.path.dirname(os.path.abspath(__file__))
    # img_path = '\\img\\icon\\neo.png'
    # full_path = os.path.join(abs_path, img_path)

    ico = Image.open('./img/icon/robo.png')
    neo_image = ImageTk.PhotoImage(ico)
    root.wm_iconphoto(False, neo_image)

    root.iconbitmap(default='./img/icon/robo.ico')

    # cria menu superior
    menu_bar = tkk.Menu(root)
    root.config(menu=menu_bar)

    # cria menu SAMP
    samp_menu = tkk.Menu(menu_bar, tearoff=False)

    # cria menu fornecimento
    fornecimento_menu_opt = tkk.Menu(menu_bar, tearoff=False)
    fornecimento_menu_opt.add_command(label='Baixar Fornecimento Energia', command=lambda:df.download(root))

    # cria menu receita
    receita_menu_opt = tkk.Menu(menu_bar, tearoff=False)
    receita_menu_opt.add_command(label='Baixar Receita Uso', command=lambda:dr.download(root))

    # cria menu extracao
    extract_menu = tkk.Menu(menu_bar, tearoff=False)

    # cria menu extracao
    extract_samp_cods_menu = tkk.Menu(menu_bar, tearoff=False)
    extract_samp_cods_menu.add_command(label='Arquivo', command=lambda:extract_cods.samp_file(listbox_log))
    extract_samp_cods_menu.add_command(label='Pasta', command=lambda:extract_cods.samp_folder(listbox_log))

    # cria menu pastas
    folders_menu = tkk.Menu(menu_bar, tearoff=False)
    folders_menu.add_command(label='Organizar', command=lambda:folders.separate_by_year_month(listbox_log))

    evolucao_menu = tkk.Menu(menu_bar, tearoff=0)
    evolucao_menu.add_command(label='Atualizar')

    # cria menu opções
    opcoes_menu = tkk.Menu(menu_bar, tearoff=False)
    opcoes_menu.add_command(label='Limpar tela', command=lambda:setup.clean_logs(listbox_log))

    # cria menu ajuda
    ajuda_menu = tkk.Menu(menu_bar, tearoff=False)
    ajuda_menu.add_command(label='Bem vindo')
    ajuda_menu.add_separator()
    ajuda_menu.add_command(label='Orientação', command=show_orientations)
    ajuda_menu.add_command(label='Sobre...', command=show_about)

    # cria menu ativos e passivos
    ativos_passivos_menu = tkk.Menu(menu_bar, tearoff=0)
    ativos_passivos_menu.add_command(label='Gerar')

    # adiciona menu em cascata
    samp_menu.add_cascade(label='Fornecimento Energia', menu=fornecimento_menu_opt, underline=0)
    samp_menu.add_cascade(label='Receita Uso', menu=receita_menu_opt, underline=0)
    extract_menu.add_cascade(label='Códigos SAMP', menu=extract_samp_cods_menu, underline=0)
    menu_bar.add_cascade(label='Mercado', menu=samp_menu, underline=0)
    menu_bar.add_cascade(label='Extrair', menu=extract_menu, underline=0)
    menu_bar.add_cascade(label='Arquivos mercado', menu=folders_menu, underline=0)
    menu_bar.add_cascade(label='Opções', menu=opcoes_menu, underline=0)
    menu_bar.add_cascade(label='Ajuda', menu=ajuda_menu, underline=0)
    menu_bar.add_command(label='Sair', command=root.destroy)

    # adiciona label caixa de log
    # box_log_label = tkk.Label(text='Caixa de Log', font=("Lato", 9), anchor=tkk.NW)
    # box_log_label.pack()

    # adiciona configurações de exibição
    frame_log = tkk.Frame()
    frame_log.pack()

    listbox_log = tkk.Listbox(frame_log, height=3, width=70, font=("Lato", 11), selectmode='multiple', bd=0)
    listbox_log.pack(side=tkk.LEFT, pady=5)

    scrollbar_log_y = tkk.Scrollbar(frame_log)
    scrollbar_log_y.config(command=listbox_log)

    scrollbar_log_y.pack(side=tkk.RIGHT, fill=tkk.Y)

    listbox_log.config(yscrollcommand=scrollbar_log_y.set)

    '''
    imagem aqui
    '''
    # label = tkk.Label(root, image = new_img)
    label = tkk.Label(root, text='SCT', font=('Lato', 12))
    label.pack(side=tkk.BOTTOM)

    # verifica senha do usuário
    root.mainloop()

if __name__=='__main__':
    main_GUI()
