# Extrator SCT
* **REF - regulação Econômica e Financeira**
* **Autor:** Jackes Tiago Ferreira da Fonseca

## Compilação

### Renomear caminho
**OBS.:**  Antes de gerar o .exe lembrar de substituir ./ por ../ em

* src/main.py
    - GUI() x3

* src/modules/samp/download_fornecimento.py
    - setup() x1
    
* src/modules/samp/download_receita.py
    - setup() x1

### Executar comando
Dentro da pasta REF rodar
python3 -m PyInstaller --noconsole --icon=img/icon/neo.png --onefile src/main.py

copiar a pasta **build** e o arquivo **.exe** para a pasta ARQUIVOS_REF/_ExtratorSCT