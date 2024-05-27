import tkinter.messagebox
import tkinter.filedialog
from modules.settings import setup
import shutil
import os

def list_files(origin_path):
    files = os.listdir(origin_path)
    txt_files = []

    if len(files) == 0:
        return files
    for file in files:
        if file.split('.')[-1] == 'txt':
            txt_files.append(file)

    return txt_files

def verify_file_format(file):
    if len(file.split('.')) == 6:
        return True
    return False

def separate_by_year_month(listbox_log):
    tkinter.messagebox.showinfo('Atenção!', 'Selecione o diretório onde estão os arquivos de mercado')

    origin_path = tkinter.filedialog.askdirectory()

    if origin_path == '':
        return

    tkinter.messagebox.showinfo('Atenção!', 'Selecione o diretório do SCT')

    dest_path = tkinter.filedialog.askdirectory()

    if dest_path == '':
        return

    if not (os.path.isdir(os.path.join(dest_path, 'SAMP_LIVRE-IMPORTAR')) or os.path.isdir(os.path.join(dest_path, 'SAMP-IMPORTAR'))):
        tkinter.messagebox.showerror('Atenção!', 'A pasta do SCT selecionada não corresponde!\n Certifique-se de selecionar a pasta raiz (contém as pastas SAMP_LIVRE-IMPORTAR e SAMP-IMPORTAR)')
        return
    
    files = list_files(origin_path)

    if len(files) == 0:
        tkinter.messagebox.showerror('Atenção!', 'A pasta selecionada está vazia ou não posusi arquivos TXT')
        return

    for file in files:
        is_correct = verify_file_format(file)

        if is_correct:
            file_date = file.split('.')[3]
            file_date = file_date.replace('M', '')
            file_date_start = file_date[:2]
            file_date_end = file_date[2:]
            file_date = file_date_end + file_date_start

            if file.split('.')[4] == 'LIVRE':
                if os.path.isdir(os.path.join(dest_path, 'SAMP_LIVRE-IMPORTAR', file_date)):
                    if os.path.isfile(os.path.join(dest_path, 'SAMP_LIVRE-IMPORTAR', file_date, file)):
                        folder = os.path.join('SAMP-IMPORTAR', file_date)
                        ans = tkinter.messagebox.askquestion('Atenção!', f'O arquivo {file} já existe na pasta {folder}. Deseja substituí-lo?')

                        if ans == 'yes':
                            os.remove(os.path.join(dest_path, 'SAMP_LIVRE-IMPORTAR', file_date, file))
                            shutil.move(os.path.join(origin_path, file), os.path.join(dest_path, 'SAMP_LIVRE-IMPORTAR', file_date))
                        else:
                            continue
                    else:
                        shutil.move(os.path.join(origin_path, file), os.path.join(dest_path, 'SAMP_LIVRE-IMPORTAR', file_date))

                elif os.path.isdir(os.path.join(dest_path, 'SAMP_LIVRE-IMPORTAR', file_date + '-ok')):
                    tkinter.messagebox.showwarning('Atenção!', f'A pasta {file_date} já foi importada para a planilha do SCT!\n')
                    tkinter.messagebox.showwarning('Atenção!', f'Para adicionar o arquivo {file} estorne os dados do mês!\n')
                    continue

                else:
                    os.mkdir(os.path.join(dest_path, 'SAMP_LIVRE-IMPORTAR', file_date))
                    shutil.move(os.path.join(origin_path, file), os.path.join(dest_path, 'SAMP_LIVRE-IMPORTAR', file_date))
            else:
                if os.path.isdir(os.path.join(dest_path, 'SAMP-IMPORTAR', file_date)):
                    if os.path.isfile(os.path.join(dest_path, 'SAMP_LIVRE-IMPORTAR', file_date, file)):
                        folder = os.path.join('SAMP-IMPORTAR', file_date)
                        ans = tkinter.messagebox.askquestion('Atenção!', f'O arquivo {file} já existe na pasta {folder}. Deseja substituí-lo?')
                        
                        if ans == 'yes':
                            os.remove(os.path.join(dest_path, 'SAMP-IMPORTAR', file_date, file))
                            shutil.move(os.path.join(origin_path, file), os.path.join(dest_path, 'SAMP-IMPORTAR', file_date))
                        else:
                            continue
                    else:
                        shutil.move(os.path.join(origin_path, file), os.path.join(dest_path, 'SAMP-IMPORTAR', file_date))

                elif os.path.isdir(os.path.join(dest_path, 'SAMP-IMPORTAR', file_date + '-ok')):
                    tkinter.messagebox.showwarning('Atenção!', f'A pasta {file_date} já foi importada para a planilha do SCT!\n')
                    tkinter.messagebox.showwarning('Atenção!', f'Para adicionar o arquivo {file} estorne os dados do mês!\n')
                    continue
                else:
                    os.mkdir(os.path.join(dest_path, 'SAMP-IMPORTAR', file_date))
                    shutil.move(os.path.join(origin_path, file), os.path.join(dest_path, 'SAMP-IMPORTAR', file_date))
                    
        else:
            listbox_log.insert(tkinter.END, f'{setup.get_actual_hour()}O arquivo {file} não está no formato correto!') 

    if is_correct:
        tkinter.messagebox.showinfo('Atenção!', 'Operação finalizada!')