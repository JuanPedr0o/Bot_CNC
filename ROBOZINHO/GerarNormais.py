from pathlib import Path
import os
import re

# Função para extrair o valor de DZ
def extrair_valor_dz(linhas):
    for linha in linhas:
        correspondencia = re.search(r'DZ=([\d\.]+)', linha)
        if correspondencia:
            return float(correspondencia.group(1))
    return None

# Função para verificar se o arquivo deve ser ignorado
def deve_ignorar_arquivo(linhas, nome_arquivo):
    valor_dz = extrair_valor_dz(linhas)
    if valor_dz == 25.5:
        print(f"Arquivo {nome_arquivo} ignorado devido ao valor DZ=25.5.")
        return True
    return False

def process_cnc_file(input_file, output_file):
    with open(input_file, 'r') as file:
        lines = file.readlines()

    # Extrai o valor de DZ
    valor_dz = extrair_valor_dz(lines)
    
    filtered_lines = []
    current_block = []
    block_start_line = None  # Armazena a linha inicial do bloco ";Vertical Milling"

    for line in lines:
        if line.startswith(";Vertical Milling"):
            # Se houver um bloco em processamento, verifica se deve manter ou apagar
            if current_block:
                result = process_block(current_block, valor_dz)  # Passa o valor de DZ
                if result:
                    if block_start_line:  # Adiciona a linha inicial apenas se o bloco for mantido
                        filtered_lines.append(block_start_line)
                    filtered_lines.extend(result)
                block_start_line = None  # Reseta a linha inicial do bloco
                current_block = []
            # Armazena a linha inicial do novo bloco
            block_start_line = line
        else:
            current_block.append(line)

    # Processa o último bloco após o último ";Vertical Milling"
    if current_block:
        result = process_block(current_block, valor_dz)  # Passa o valor de DZ
        if result:
            if block_start_line:  # Adiciona a linha inicial apenas se o bloco for mantido
                filtered_lines.append(block_start_line)
            filtered_lines.extend(result)

    with open(output_file, 'w') as file:
        file.writelines(filtered_lines)

def process_block(block_lines, valor_dz):
    # Conta o número de linhas do bloco
    block_length = len(block_lines)
    contains_t101 = any("T=101" in line for line in block_lines)

    if block_length < 11:
        # Blocos com menos de 10 linhas são apagados se contiverem T=101
        if contains_t101:
            return []  # Apaga o bloco
        else:
            return block_lines  # Mantém o bloco

    # Blocos com 10 ou mais linhas são mantidos automaticamente
    return block_lines


config_path = 'config_dir.txt'
if not os.path.exists(config_path):
    raise FileNotFoundError("Arquivo config_dir.txt não encontrado!")

with open(config_path, 'r') as f:
    config_dir = Path(f.read().strip())

input_folder = config_dir  # Processa arquivos diretamente aqui
output_folder = config_dir  # Salva os resultados no mesmo lugar

os.makedirs(output_folder, exist_ok=True)

# Processamento recursivo com os.walk()
for root, dirs, files in os.walk(input_folder):
    for filename in files:
        if filename.endswith('.xxl'):
            input_path = os.path.join(root, filename)
            
            # Cria estrutura de pastas equivalente na saída
            relative_path = os.path.relpath(root, input_folder)
            output_path_dir = os.path.join(output_folder, relative_path)
            os.makedirs(output_path_dir, exist_ok=True)
            
            output_path = os.path.join(output_path_dir, filename)

            # Verificação de arquivo existente (opcional)
            # if os.path.exists(output_path):
            #     continue

            with open(input_path, 'r') as file:
                lines = file.readlines()

            if deve_ignorar_arquivo(lines, filename):
                continue

            process_cnc_file(input_path, output_path)
            print(f"Arquivo processado: {input_path} -> {output_path}")

print(f"Todos os arquivos foram salvos na pasta '{output_folder}'.")