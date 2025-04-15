from pathlib import Path
import os
import re

# Função para verificar se o arquivo contém DZ=25.5 na primeira linha
def verificar_dz_primeira_linha(linhas):
    if linhas and "DZ=25.5" in linhas[0]:
        return True
    return False

# Função para processar o arquivo
def processar_arquivo_dz(input_file, output_file):
    with open(input_file, 'r') as file:
        lines = file.readlines()

    # Verifica se DZ=25.5 está na primeira linha
    if not verificar_dz_primeira_linha(lines):
        print(f"Arquivo {input_file} não contém DZ=25.5 na primeira linha. Ignorando...")
        return

    linhas_processadas = []
    bloco_atual = []

    for linha in lines:
        if linha.startswith(";Vertical Milling"):
            # Processar o bloco anterior
            if bloco_atual:
                resultado = processar_bloco(bloco_atual)
                if resultado:
                    linhas_processadas.extend(resultado)
                bloco_atual = []
            bloco_atual.append(linha)  # Começa um novo bloco
        else:
            bloco_atual.append(linha)

    # Processar o último bloco
    if bloco_atual:
        resultado = processar_bloco(bloco_atual)
        if resultado:
            linhas_processadas.extend(resultado)

    # Salva o arquivo processado
    if linhas_processadas:
        with open(output_file, 'w') as file:
            file.writelines(linhas_processadas)
        print(f"Arquivo processado e salvo em: {output_file}")
    else:
        print(f"Arquivo {input_file} não gerou saída válida.")

# Função para processar blocos e apagar blocos com T=101
def processar_bloco(bloco):
    if not bloco:
        return []

    # Verifica se o bloco contém T=101
    for linha in bloco:
        if "T=101" in linha:
            return []  # Apaga o bloco inteiro

    # Substituir T=102 por T=112
    novo_bloco = []
    for linha in bloco:
        if "T=102" in linha:
            linha = linha.replace("T=102", "T=112")
        novo_bloco.append(linha)

    return novo_bloco

# Função principal modificada para processar subpastas
def processar_pastas_dz(input_folder, output_folder):
    # Processa recursivamente todas as subpastas
    for root, dirs, files in os.walk(input_folder):
        for arquivo in files:
            if arquivo.endswith('.xxl'):
                caminho_entrada = os.path.join(root, arquivo)
                
                # Cria estrutura de subpastas correspondente na saída
                relative_path = os.path.relpath(root, input_folder)
                caminho_saida_root = os.path.join(output_folder, relative_path)
                os.makedirs(caminho_saida_root, exist_ok=True)
                
                caminho_saída = os.path.join(caminho_saida_root, arquivo)

                # Ignora se o arquivo já existe (opcional)
                # if os.path.exists(caminho_saída):
                #     continue

                print(f"Processando arquivo: {caminho_entrada}")
                processar_arquivo_dz(caminho_entrada, caminho_saída)

    print("Todos os arquivos foram processados com sucesso!")

# Configuração de caminhos
config_path = 'config_dir.txt'
if not os.path.exists(config_path):
    raise FileNotFoundError("Arquivo config_dir.txt não encontrado!")

with open(config_path, 'r') as f:
    config_dir = Path(f.read().strip())

input_folder = config_dir
output_folder = config_dir

# Executar o processamento
processar_pastas_dz(input_folder, output_folder)

print("Executando GerarRetalhos.py...")