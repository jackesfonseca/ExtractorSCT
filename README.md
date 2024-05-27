# Extrator SCT
* **REF - regulação Econômica e Financeira**
* **Autor:** Jackes Tiago Ferreira da Fonseca

## SCT
![image](https://github.com/jackesfonseca/ExtractorSCT/assets/53023400/5a1e2543-ed65-4e14-a0cf-c7f912dc29ac)

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
python3 -m PyInstaller --noconsole --icon=img/icon/robo.png --onefile src/main.py

copiar a pasta **build** e o arquivo **.exe** para a pasta ARQUIVOS_REF/_ExtratorSCT
