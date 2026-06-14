import os
import shutil
import datetime
import json
import logging
import re 

def configurar_logger():
    #Configura o log para salvar em arquivo e mostrar na tela.
    logger = logging.getLogger("OrganizadorDePastas")
    logger.setLevel(logging.INFO)
    
    # Evita duplicar se a função for chamada mais de uma vez
    if not logger.handlers:
        #Define o arquivo onde o histórico será salvo
        file_handler = logging.FileHandler("historico_movimentacoes.log", encoding="utf-8")
        #Define a saída para a tela
        console_handler = logging.StreamHandler()
        
        #Formatação da mensagem de log: [Data/Hora] - [NÍVEL] - Mensagem
        formato = logging.Formatter('[%(asctime)s] - %(levelname)s - %(message)s', datefmt='%d/%m/%Y %H:%M:%S')
        file_handler.setFormatter(formato)
        console_handler.setFormatter(formato)
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

    return logger

def carregar_configuracoes(caminho_json='config.json'):
    try:
        with open(caminho_json, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Erro ao carregar JSON: {e}")
        return None

try:
    from win32api import GetFileVersionInfo, LOWORD, HIWORD
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

def get_version_info(file_path):
    if not WIN32_AVAILABLE: return None
    try:
        info = GetFileVersionInfo(file_path, "\\")
        ms, ls = info['FileVersionMS'], info['FileVersionLS']
        return f"{HIWORD(ms)}.{LOWORD(ms)}.{HIWORD(ls)}"
    except: return None

def organizar_entrega():
    log = configurar_logger()
    log.info("- Iniciando o processo de organização de pastas -")

    # Carrega os caminhos do JSON
    config = carregar_configuracoes()
    if not config: 
        log.error("Processo abortado: Falha ao carregar configurações")
        return

    origem = config['diretorio_origem']
    entrega = config['diretorio_base_entrega']
    pasta_lib_fixa = config['pasta_liberados_fixa']
    extensoes = config['extensoes_permitidas']
    
    data_hoje = datetime.datetime.now().strftime("%m%d")

    # Caminhos fixos
    p3_liberados = os.path.join(entrega, pasta_lib_fixa)
    p4_volta = os.path.join(entrega, "_VOLTA_VERSAO_", data_hoje)

    #1 - Backup
    if os.path.exists(p3_liberados):
        os.makedirs(p4_volta, exist_ok=True)
        log.info(f"Realizando backup completo de '{pasta_lib_fixa}' para a pasta Volta Versao...")
        try:
            arquivos_backup = 0
            for item in os.listdir(p3_liberados):
                caminho_item = os.path.join(p3_liberados, item)
                if os.path.isfile(caminho_item):
                    shutil.copy2(caminho_item, os.path.join(p4_volta, item))
                    arquivos_backup += 1
            log.info(f"Backup de Volta Versão concluído com sucesso ({arquivos_backup} arquivos salvos)\n")
        except Exception as e:
            log.error(f"Erro durante o backup de Volta Versão: {e}\n")
    else:
        log.warning(f"A pasta '{pasta_lib_fixa}' não existe ainda. Não há arquivos para realizar o backup prévio.\n")
        os.makedirs(p3_liberados, exist_ok=True)


    #2 - busca a maior versão para organizar os arquivos
    arquivos_validos = []
    maiores_versoes = {} # Dicionário para guardar a maior versão de cada 'nome_limpo'

    for arquivo in os.listdir(origem):
        caminho_full = os.path.join(origem, arquivo)
        
        #Valida extensão baseada no JSON
        if not os.path.isfile(caminho_full) or not any(arquivo.lower().endswith(ext) for ext in extensoes):
            continue

        versao_tripla = get_version_info(caminho_full)
        if not versao_tripla: 
            log.warning(f"Ignorado: {arquivo} (Não foi possível obter a versão)")
            continue

        #Extrai o nome limpo cortando antes do primeiro '_' seguido de número
        nome_base = os.path.splitext(arquivo)[0]
        extensao = os.path.splitext(arquivo)[1]
        nome_base_limpo = re.split(r'_\d+', nome_base)[0]
        nome_limpo = nome_base_limpo + extensao

        # Converte a string de versão (ex: '25.8.69') para uma tupla (25, 8, 69) para comparação correta
        try:
            tupla_versao = tuple(map(int, versao_tripla.split('.')))
        except ValueError:
            tupla_versao = (0, 0, 0)

        # Guarda os dados do arquivo na lista
        arquivos_validos.append({
            'arquivo': arquivo,
            'caminho_full': caminho_full,
            'versao_tripla': versao_tripla,
            'tupla_versao': tupla_versao,
            'nome_limpo': nome_limpo
        })

        # Atualiza o dicionário se esta for a maior versão encontrada para este executável até agora
        if nome_limpo not in maiores_versoes or tupla_versao > maiores_versoes[nome_limpo]:
            maiores_versoes[nome_limpo] = tupla_versao


    #3 - processamento e distribuição dos executáveis
    for item in arquivos_validos:
        arquivo = item['arquivo']
        caminho_full = item['caminho_full']
        versao_tripla = item['versao_tripla']
        nome_limpo = item['nome_limpo']
        tupla_versao = item['tupla_versao']

        # Mapeamento dos caminhos
        p1_atualizador = os.path.join(entrega, "_Executaveis_Liberados_Atualizador", versao_tripla)
        p2_espelho = os.path.join(entrega, "espelho_entrega", data_hoje)

        for p in [p1_atualizador, p2_espelho]:
            os.makedirs(p, exist_ok=True)

        try:
            log.info(f"Processando arquivo: {arquivo} (Versão: {versao_tripla})")

            #1 - Distribuição para o Atualizador (Sempre copia)
            shutil.copy2(caminho_full, os.path.join(p1_atualizador, nome_limpo))
            log.info(f"  -> Copiado: '{nome_limpo}' para o Atualizador.")

            #2 - Distribuição para Liberados Fixa (Apenas se for a maior versão)
            if tupla_versao == maiores_versoes[nome_limpo]:
                shutil.copy2(caminho_full, os.path.join(p3_liberados, nome_limpo))
                log.info(f"  -> Copiado: '{nome_limpo}' para Liberados Fixa (Maior versão do lote).")
            else:
                log.info(f"  -> Ignorado na Liberados Fixa: '{nome_limpo}' (Uma versão superior foi encontrada neste lote).")

            #3 - Mover o arquivo original para o Espelho de Entrega (Sempre mover para rastreabilidade)
            shutil.move(caminho_full, os.path.join(p2_espelho, arquivo))
            log.info(f"  -> Movido: Original '{arquivo}' para o Espelho de Entrega.")
            log.info(f"Sucesso: Fluxo do arquivo '{nome_limpo}' concluído.\n")

        except Exception as e:
            log.error(f"Erro crítico ao processar o arquivo '{arquivo}': {e}\n")
            
    log.info("--- Processo de organização finalizado ---\n")

if __name__ == "__main__":
    organizar_entrega()