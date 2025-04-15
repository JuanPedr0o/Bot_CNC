from pathlib import Path
import os
import re

def modificar_ad(linhas):
    for i, linha in enumerate(linhas):
        if "-AD" in linha:
            match = re.search(r'DY=([\d\.]+)', linha)
            if match:
                dy = float(match.group(1))
                if dy < 850:
                    linhas[i] = linha.replace("-AD", "-I")
                else:
                    linhas[i] = linha.replace("-AD", "-A")
    return linhas

def process_r_files(input_folder, output_folder):
    # Processa recursivamente todas as subpastas
    for root, dirs, files in os.walk(input_folder):
        for filename in files:
            if filename.startswith('R') and filename.endswith('.xxl'):
                input_path = os.path.join(root, filename)
                
                # Cria estrutura de subpastas equivalente na saída
                relative_path = os.path.relpath(root, input_folder)
                output_path_dir = os.path.join(output_folder, relative_path)
                os.makedirs(output_path_dir, exist_ok=True)
                
                output_path = os.path.join(output_path_dir, filename)

                try:
                    with open(input_path, 'r') as file:
                        lines = file.readlines()
                except Exception as e:
                    print(f"Erro ao ler {input_path}: {e}")
                    continue

                # Aplica as modificações
                modified_lines = modificar_ad(lines)
                
                # Remove linhas com P_LOAD/P_UNLOAD e adiciona linha final
                final_lines = []
                for line in modified_lines:
                    if "P_LOAD" in line or "P_UNLOAD" in line:
                        continue
                    final_lines.append(line)
                final_lines.append("N X=3600 Y=46\n")

                try:
                    with open(output_path, 'w') as file:
                        file.writelines(final_lines)
                    print(f"Arquivo processado: {input_path} -> {output_path}")
                except Exception as e:
                    print(f"Erro ao salvar {output_path}: {e}")

# Configuração de caminhos
config_path = 'config_dir.txt'
if not os.path.exists(config_path):
    raise FileNotFoundError("Arquivo config_dir.txt não encontrado!")

with open(config_path, 'r') as f:
    config_dir = Path(f.read().strip())

input_folder = config_dir
output_folder = config_dir

# Executa o processamento
process_r_files(input_folder, output_folder)

print(f"Todos os arquivos foram processados e salvos na pasta '{output_folder}'.")
print("Executando Extras.py...")