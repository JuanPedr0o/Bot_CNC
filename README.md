Scripts de Automação Industrial
Processamento de arquivos CNC (.xxl) e geração de relatórios/planilhas.

1. GerarNormais.py
Função: Processa arquivos .xxl padrão.

Ações:

Ignora arquivos com DZ=25.5.

Remove blocos ;Vertical Milling com T=101 (se tiverem menos de 11 linhas).

Mantém estrutura de pastas original.

2. GerarOsDe25.py
Função: Focado em arquivos com DZ=25.5.

Ações:

Remove blocos com T=101.

Substitui T=102 por T=112.

Mantém subpastas.

3. GerarRetalhos.py
Função: Processa arquivos de retalhos (nomes começando com R).

Ações:

Modifica comandos -AD para -I ou -A (baseado em DY).

Remove linhas P_LOAD/P_UNLOAD.

Adiciona linha final N X=3600 Y=46.

4. start.py (Script Principal)
Função: Interface gráfica e fluxo principal.

Ações:

Cria backup da pasta selecionada.

Gera planilhas (Lote.xlsx, LISTA-RETALHOS.xlsm, etc.).

Processa arquivos .TXT e .xxl.

Renomeia arquivos com base no conteúdo.

Executa os demais scripts na ordem:
GerarNormais.py → GerarOsDe25.py → GerarRetalhos.py → Extras.py.

5. Extras.py
Função: Ajustes finais em arquivos .xxl.

Ações:

Define BZ=15.

Corrige valores de Z baseados em DZ.

Insere XGIN em blocos específicos.

Dependências
Python 3.x

Bibliotecas: pandas, openpyxl, tkinter, unidecode, pathlib.

Configuração
Um arquivo config_dir.txt armazena o diretório de trabalho.

Importante: Execute start.py primeiro para configurar automaticamente.
