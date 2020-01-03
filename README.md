# Power-Point-Presentation

Mini-projeto para gerar automaticamente apresentações power-point a partir da letra de música, projeto para o ministério de louvor da IPND

<p align="center">
    <img src="https://raw.githubusercontent.com/JonasFreireAlcantara/Power-Point-Presentation/master/screenshot_slide.png" width="550">
</p>

#### Configuração
Este script requer ter o python3 e pip instalados no seu sistema
será necessário executar o seguinte comando para instalar as dependências necessárias:
```shell script
  pip install -r requirements.txt
```

#### Arquivos fonte:
* ml.py
* slides.py


#### Utilização

A sintaxe para gerar as apresentações é a seguinte:
```shell script
    python slides.py <<pasta_letras>> <<pasta_slides>> <<imagem_de_fundo>> <<imagem_logo>>\n')

 pasta_letras: caminho para a pasta que contém as letras da músicas
 pasta_slides: caminho para a pasta que conterá os slides gerados 
 imagem_de_fundo: caminho para a imagem que será utilizada como background
 imagem_logo: caminho para a imagem que será utilizada como logo da IPND
```
O diretório python-pptx-tutorial possui alguns arquivos utilizados para estudo da biblioteca python-pptx

-----------------------------------------

Mini-project to automatically generating power-point presentations from music letter, project to IPND praise minister

#### Setup
This script requires python3 and pip installed in your system
it will be necessary to run this command for dependencies installation:
```shell script
  pip install -r requirements.txt
```

#### Source files
* ml.py
* slides.py

#### Usage

The syntax to generate the presentations:
```shell script
    python slides.py <<pasta_letras>> <<pasta_slides>> <<imagem_de_fundo>> <<imagem_logo>>\n')

 pasta_letras: path to folder containin music letters
 pasta_slides: path to folder which will hold the generated slides 
 imagem_de_fundo: path to image which will be used as background
 imagem_logo: path to image which will be used as IPND's logo
```

The python-pptx-tutorial directory contains some file utilized for python-pptx library study
