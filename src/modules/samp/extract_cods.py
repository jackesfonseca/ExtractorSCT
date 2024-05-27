import tkinter.filedialog
import tkinter.messagebox
from threading import Thread
from modules.settings import setup
import pandas as pd
import os

def list_files(path_origin):
    files = os.listdir(path_origin)
    only_xls_files = []

    for file in files:
        if os.path.isfile(os.path.join(path_origin, file)) and file.split('.')[-1] == 'xls':
            only_xls_files.append(file)

    return only_xls_files

def get_file_date(df, file_type):
    if file_type == 'CATIVO':
        return str(df.iloc[4]['Unnamed: 13'])
        
    return str(df.iloc[4]['Unnamed: 15'])
    
def get_type_file(df):
    cols = len(df.axes[1])

    if cols == 15:
        return 'CATIVO'
    return 'LIVRE'

def verify_unique_date(df):
    pass

def folder_has_files(path):
    files = os.listdir(path)

    if len(files) > 0:
        return True

    return False

def delete_all_files(path):
    files = os.listdir(path)

    for file in files:
        os.remove(os.path.join(path, file))

def samp_folder(listbox_log):
    Thread(target=samp_folder_thread(listbox_log)).start()

def samp_folder_thread(listbox_log):
    path_origin = tkinter.filedialog.askdirectory()
    path_dest = ''

    if path_origin == '':
        return

    save_same_folder = tkinter.messagebox.askquestion('Atenção!', 'Deseja salvar os arquivos no mesmo diretório?')

    if save_same_folder == 'no':
        tkinter.messagebox.showinfo('Atenção!', 'Selecione a pasta onde os arquivos serão salvos')

        while path_dest == '':
            path_dest = tkinter.filedialog.askdirectory()

        if folder_has_files(path_dest):
            delete_files = tkinter.messagebox.askquestion('Atenção!', 'A pasta selecionada possui arquivos salvos!\n Deseja excluir todos os arquivos?')

            if delete_files == 'yes':
                delete_all_files(path_dest)

    else:
        path_dest = path_origin

    files = list_files(path_origin)

    if len(files) == 0:
        tkinter.messagebox.showwarning('Atenção!', 'Não foram identificados arquivos na pasta informada!\n Verifique se a pasta correta foi selecionada, se existem arquivos ou se os arquivos estão corretos!')
        return

    else:
        ans_two = tkinter.messagebox.askquestion('Atenção!', 'Deseja excluir os arquivos originais?')

        msg = setup.get_actual_hour() + 'Iniciando extração\n'
        listbox_log.insert(tkinter.END, msg)

        for file in files:
            file_ext = file.split('.')[-1]
            if os.path.getsize(os.path.join(path_origin, file)) < 30000:
                tkinter.messagebox.showwarning('Atenção!', f'O tamanho do arquivo {file} não corresponde! Verifique se o arquivo selecionado possui dados e/ou se os dados estão completos!')
                msg = setup.get_actual_hour() + 'Arquivo ' + file + ' inválido!'
                listbox_log.insert(tkinter.END, msg)
                continue

            try:
                msg = setup.get_actual_hour() + 'Fazendo leitura do arquivo ' + file
                listbox_log.insert(tkinter.END, msg)
                df = pd.read_excel(os.path.join(path_origin, file))

                try:
                    file_type = get_type_file(df)
                    file_date = get_file_date(df, file_type)

                    if file_type == 'CATIVO':
                        columns_to_drop = ['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 5', 'Unnamed: 7', 'Unnamed: 9', 'Unnamed: 11','Unnamed: 13']
                    else:
                        columns_to_drop = ['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 7', 'Unnamed: 9', 'Unnamed: 11', 'Unnamed: 13', 'Unnamed: 15']
                        
                    df.drop(columns=columns_to_drop, inplace=True)

                    try:
                        if file_type == 'CATIVO':
                            df['Unnamed: 0'] = df['Unnamed: 4'].astype(str) + ';' + df['Unnamed: 6'].astype(str) + ';' + df['Unnamed: 8'].astype(str) + ';' + df['Unnamed: 10'].astype(str) + ';' + df['Unnamed: 12'].astype(str) + ';' + df['Unnamed: 14'].astype(str)
                            columns_to_drop = ['Unnamed: 4', 'Unnamed: 6', 'Unnamed: 8', 'Unnamed: 10', 'Unnamed: 12', 'Unnamed: 14']
                            df.drop(columns=columns_to_drop, inplace=True)
                            df = df.iloc[2:, 0].to_frame()

                        else:
                            df['Unnamed: 3'] = df['Unnamed: 3'].fillna(0)
                            df['Unnamed: 0'] = df['Unnamed: 3'].astype(str) + ';' + df['Unnamed: 6'].astype(str) + ';' + df['Unnamed: 8'].astype(str) + ';' + df['Unnamed: 10'].astype(str) + ';' + df['Unnamed: 12'].astype(str) + ';' + df['Unnamed: 14'].astype(str) + ";" + df['Unnamed: 16'].astype(str)
                            columns_to_drop = ['Unnamed: 3', 'Unnamed: 6', 'Unnamed: 8', 'Unnamed: 10', 'Unnamed: 12', 'Unnamed: 14', 'Unnamed: 16']
                            df.drop(columns=columns_to_drop, inplace=True)
                            df = df.iloc[1:, 0].to_frame()

                        txt_file = file.replace(file_ext, 'txt')
                        msg = setup.get_actual_hour() + 'Salvando arquivo ' + 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1]
                        listbox_log.insert(tkinter.END, msg)

                        txt_file = file.replace(file_ext, 'txt')
                        if os.path.isfile(os.path.join(path_origin, 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1])):
                            msg = 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1]
                            ans = tkinter.messagebox.askquestion('Atenção!', f'O arquivo {msg} já existe, deseja substituir?')

                            if ans == 'yes':
                                txt_file = file.replace(file_ext, 'txt')
                                df.to_csv(os.path.join(path_dest, 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1]), sep='\t',
                                        header=False, index=False)
                                msg = setup.get_actual_hour() + 'Arquivo ' + file.replace(file_ext, 'txt') + ' substituído!'
                                listbox_log.insert(tkinter.END, msg)
                            else:
                                msg = setup.get_actual_hour() + 'Arquivo ' + file.replace(file_ext, 'txt') + ' cancelado!'
                                listbox_log.insert(tkinter.END, msg)
            
                                continue
                        else:
                            txt_file = file.replace(file_ext, 'txt')
                            df.to_csv(os.path.join(path_dest, 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1]), sep='\t', header=False,
                                    index=False)
                            msg = setup.get_actual_hour() + 'Arquivo ' + file.replace(file_ext, 'txt') + ' salvo!'
                            listbox_log.insert(tkinter.END, msg)

                            if ans_two == 'yes':
                                msg = setup.get_actual_hour() + 'Arquivo original ' + file + ' excluído!'
                                listbox_log.insert(tkinter.END, msg)
                                os.remove(os.path.join(path_origin, file))
                            else:
                                continue
                        
    

                    except Exception as e:
        
                        tkinter.messagebox.showerror('Atenção!', f'Ocorreu um erro ao salvar o arquivo!\n {e}')

                except Exception as e:
    
                    tkinter.messagebox.showerror('Atenção!', f'Ocorreu um erro inesperado!\n {e}')

            except Exception as e:

                tkinter.messagebox.showerror('Atenção!', f'Erro ao tentar abrir arquivo!\n {e}')
        
        msg = setup.get_actual_hour() + 'Concluído!'
        listbox_log.insert(tkinter.END, msg)
        tkinter.messagebox.showinfo('Atenção!', 'Extração realizada com sucesso')

def samp_file(listbox_log):
    Thread(target=samp_file_thread(listbox_log)).start()

def samp_file_thread(listbox_log):
    path_origin = tkinter.filedialog.askopenfilename()
    file = path_origin.split('/')[-1]
    only_path_origin = path_origin.replace(file, '')
    file_ext = file.split('.')[-1]
    path_dest = ''

    if path_origin == '':
        return

    elif file_ext != 'xls' and file_ext != 'xlsx':
        tkinter.messagebox.showwarning('Atenção!', 'Extenção de arquivo não corresponde!\n Insira um arquivo xls ou xlsx')
        return

    if os.path.getsize(path_origin) < 30000:
        tkinter.messagebox.showwarning('Atenção!',f'O tamanho do arquivo {file} não corresponde! Verifique se o arquivo selecionado possui dados e/ou se os dados estão completos!')
        return

    save_same_folder = tkinter.messagebox.askquestion('Atenção!', 'Deseja salvar os arquivos no mesmo diretório?')

    if save_same_folder == 'no':
        tkinter.messagebox.showinfo('Atenção!', 'Selecione a pasta onde os arquivos serão salvos')

        while path_dest == '':
            path_dest = tkinter.filedialog.askdirectory()

    else:
        path_dest = only_path_origin

    ans_two = tkinter.messagebox.askquestion('Atenção!', 'Deseja excluir o arquivo original?')

    msg = setup.get_actual_hour() + 'Iniciando extração\n'
    listbox_log.insert(tkinter.END, msg)

    try:
        df = pd.read_excel(path_origin)

        try:
            file_type = get_type_file(df)
            file_date = get_file_date(df, file_type)
            
            msg = setup.get_actual_hour() + 'Fazendo leitura do arquivo ' + file
            listbox_log.insert(tkinter.END, msg)
            
            if file_type == 'CATIVO':
                columns_to_drop = ['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 5', 'Unnamed: 7', 'Unnamed: 9', 'Unnamed: 11','Unnamed: 13']
            else:
                columns_to_drop = ['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 7', 'Unnamed: 9', 'Unnamed: 11', 'Unnamed: 13', 'Unnamed: 15']
            df.drop(columns=columns_to_drop, inplace=True)            

            try:
                if file_type == 'CATIVO':
                    df['Unnamed: 0'] = df['Unnamed: 4'].astype(str) + ';' + df['Unnamed: 6'].astype(str) + ';' + df['Unnamed: 8'].astype(str) + ';' + df['Unnamed: 10'].astype(str) + ';' + df['Unnamed: 12'].astype(str) + ';' + df['Unnamed: 14'].astype(str)
                    
                    columns_to_drop = ['Unnamed: 4', 'Unnamed: 6', 'Unnamed: 8', 'Unnamed: 10', 'Unnamed: 12', 'Unnamed: 14']
                    df.drop(columns=columns_to_drop, inplace=True)
                    df = df.iloc[2:, 0].to_frame()

                else:
                    df['Unnamed: 3'] = df['Unnamed: 3'].fillna(0)
                    df['Unnamed: 0'] = df['Unnamed: 3'].astype(str) + ';' + df['Unnamed: 6'].astype(str) + ';' + df['Unnamed: 8'].astype(str) + ';' + df['Unnamed: 10'].astype(str) + ';' + df['Unnamed: 12'].astype(str) + ';' + df['Unnamed: 14'].astype(str) + ";" + df['Unnamed: 16'].astype(str)
                    columns_to_drop = ['Unnamed: 6', 'Unnamed: 8', 'Unnamed: 10', 'Unnamed: 12', 'Unnamed: 14', 'Unnamed: 16']
                    df.drop(columns=columns_to_drop, inplace=True)
                    df = df.iloc[1:, 0].to_frame()

                txt_file = file.replace(file_ext, 'txt')

                if os.path.isfile(os.path.join(only_path_origin, 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1])):
                    msg = 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1]
                    ans = tkinter.messagebox.askquestion('Atenção!', f'O arquivo {msg} já existe, deseja substituir?')

                    if ans == 'yes':
                        msg = setup.get_actual_hour() + 'Salvando arquivo ' + 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1]                     
                        listbox_log.insert(tkinter.END, msg)

                        df.to_csv(os.path.join(path_dest, 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1]), sep='\t', header=False, index=False)
                        msg = setup.get_actual_hour() + 'Arquivo ' + file.replace(file_ext, 'txt') + ' salvo!'
                        listbox_log.insert(tkinter.END, msg)
    
                        tkinter.messagebox.showinfo('Atenção!', 'Extração realizada com sucesso!')

                        if ans_two == 'yes':
                            os.remove(path_origin)
                    
                    else:
                        msg = setup.get_actual_hour() + 'Arquivo ' + file.replace(file_ext, 'txt') + ' cancelado!'
                        listbox_log.insert(tkinter.END, msg)
    
                else:
                    msg = setup.get_actual_hour() + 'Salvando arquivo ' + 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1]                   
                    listbox_log.insert(tkinter.END, msg)
                    df.to_csv(os.path.join(path_dest, 'P.GCO.NEWSAMPB.M' + file_date.split('-')[0] + file_date.split('-')[1] + '.' + file_type + '.' + txt_file.split('.')[-1]), sep='\t', header=False,index=False)
                    tkinter.messagebox.showinfo('Atenção!', 'Extração realizada com sucesso')

                    msg = setup.get_actual_hour() + 'Arquivo ' + file.replace(file_ext, 'txt') + ' salvo!'
                    listbox_log.insert(tkinter.END, msg)


                    if ans_two == 'yes':
                        os.remove(path_origin)
            except Exception as e:
                tkinter.messagebox.showerror('Atenção!', f'Ocorreu um erro ao salvar o arquivo!\n {e}')

        except Exception as e:
            tkinter.messagebox.showerror('Atenção!', f'Ocorreu um erro inesperado!\n {e}')

    except Exception as e:
        tkinter.messagebox.showerror('Atenção!', f'Erro ao tentar abrir arquivo!\n {e}')

if __name__=='__main__':
    samp_folder('')