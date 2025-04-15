from pathlib import Path
import os
import re

def process_cnc_file(input_folder, output_folder):
    # Processa recursivamente todas as subpastas
    for root, dirs, files in os.walk(input_folder):
        for filename in files:
            if filename.endswith('.xxl'):
                input_path = os.path.join(root, filename)
                
                # Cria estrutura de pastas equivalente na saída
                relative_path = os.path.relpath(root, input_folder)
                output_path_dir = os.path.join(output_folder, relative_path)
                os.makedirs(output_path_dir, exist_ok=True)
                
                output_path = os.path.join(output_path_dir, filename)

                # Processamento do arquivo
                try:
                    with open(input_path, 'r') as file:
                        lines = file.readlines()
                except Exception as e:
                    print(f"Erro ao ler {input_path}: {e}")
                    continue

                modified_lines = []
                current_block = []
                dz_value = None

                for line in lines:
                    # Captura DZ do cabeçalho
                    if dz_value is None:
                        match_dz = re.search(r'DZ=([\d\.]+)', line)
                        if match_dz:
                            dz_value = float(match_dz.group(1))
                    
                    if line.startswith(";Vertical Milling"):
                        if current_block:
                            modified_lines.extend(process_block(current_block, dz_value))
                        current_block = [line]
                    else:
                        current_block.append(line)

                # Processa último bloco
                if current_block:
                    modified_lines.extend(process_block(current_block, dz_value))

                # Salva o arquivo modificado
                try:
                    with open(output_path, 'w') as file:
                        file.writelines(modified_lines)
                    print(f"Arquivo processado: {input_path} -> {output_path}")
                except Exception as e:
                    print(f"Erro ao salvar {output_path}: {e}")

def process_block(block, dz_value):
    modified_block = []
    insert_xgin = False
    bloco_tem_xgin = any("XGIN" in line for line in block)

    # Aplica as modificações
    for i, line in enumerate(block):
        # Ajusta valores de Z
        if dz_value is not None:
            line = re.sub(
                r'Z=([\d\.]+)',
                lambda m: f"Z={dz_value}" if float(m.group(1)) > dz_value else m.group(0),
                line
            )
        
        # Força BZ=15
        line = re.sub(r'BZ=([\d\.]+)', 'BZ=15', line)
        
        # Verifica necessidade de inserir XGIN
        if not bloco_tem_xgin:
            condicoes = (
                (any("T=102" in l for l in block) and len(block) <= 11) or
                (any("T=112" in l for l in block) and len(block) <= 10) or
                len(block) > 14
            )
            if condicoes and i == 2:
                modified_block.append("XGIN Q=1 R=5 G=1\n")
                insert_xgin = True
        
        modified_block.append(line)
    
    return modified_block

# Configuração de caminhos
config_path = 'config_dir.txt'
if not os.path.exists(config_path):
    raise FileNotFoundError("Arquivo config_dir.txt não encontrado!")

with open(config_path, 'r') as f:
    config_dir = Path(f.read().strip())

input_folder = config_dir
output_folder = config_dir

# Executa o processamento
process_cnc_file(input_folder, output_folder)

print(f"Todos os arquivos foram processados e salvos na pasta '{output_folder}'.")