import os
import csv
import re
import json 

try:
    from win32api import GetFileVersionInfo, LOWORD, HIWORD
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

def get_versao_arquivo(caminho_arquivo):
    #Verifica a versão interna completa do arquivo (ex: 25.8.69.814)
    if not WIN32_AVAILABLE: return "Sem versão"
    try:
        info = GetFileVersionInfo(caminho_arquivo, "\\")
        ms, ls = info['FileVersionMS'], info['FileVersionLS']
        # HIWORD(ms) = Major, LOWORD(ms) = Minor = ms most significant
        # HIWORD(ls) = Build, LOWORD(ls) = Revision = ls least significant
        return f"{HIWORD(ms)}.{LOWORD(ms)}.{HIWORD(ls)}.{LOWORD(ls)}"
    except: 
        return "Sem versão"

def carregar_modulos_oficiais(arquivo_modulos_json):
    #Verifica o arquivo JSON e retorna uma lista de módulos oficiais em maiúsculo
    if not os.path.exists(arquivo_modulos_json):
        print(f"Arquivo '{arquivo_modulos_json}' não encontrado. A validação de módulos não será feita")
        return set()
    
    try:
        with open(arquivo_modulos_json, 'r', encoding='utf-8') as f:
            dados = json.load(f)
            return {item['MODULO'].upper() for item in dados if 'MODULO' in item}
    except Exception as e:
        print(f"Erro ao ler {arquivo_modulos_json}: {e}")
        return set()

def validar_modulo_json(modulo_extraido_nome, lista_modulos_oficiais):
    #Faz o de-para do módulo do arquivo com o JSON oficial
    if not lista_modulos_oficiais:
        return modulo_extraido_nome.upper(), "Sem Validação (JSON ausente)"

    modulo_maiusculo = modulo_extraido_nome.upper()

    #Correspondência exata (Ex: APO030 -> APO030 ou CML_TELAS -> CML_TELAS)
    if modulo_maiusculo in lista_modulos_oficiais:
        return modulo_maiusculo, "Sim"

    #Tenta adicionar '010' comum nos módulos para padronizar com o de-para (Ex: PCP -> PCP010)
    if f"{modulo_maiusculo}010" in lista_modulos_oficiais:
        return f"{modulo_maiusculo}010", "Sim (Ajustado com 010)"

    return modulo_maiusculo, "Não (Módulo não consta no JSON)"


def gerar_extracao_dados(
    pasta_executaveis,
    caminho_json_modulos,
    arquivo_saida_csv='dados_arquivos_bi.csv'
):
    #Faz a varredura na pasta e gera um CSV com os metadados dos arquivos

    print(f"Iniciando leitura na pasta: {pasta_executaveis}")

    #Carrega a lista oficial de módulos
    lista_modulos_oficiais = carregar_modulos_oficiais(caminho_json_modulos)

    #Linhas que serão gravadas no CSV
    linhas_relatorio = []
    padrao_nome_arquivo = re.compile(
        r"""
        ^(?P<modulo>.*?)_
        (?P<primeiro_numero>\d+)
        (?:_(?P<segundo_numero>\d+))?
        (?:_(?P<status>[^_]+))?
        (?:_(?P<info_extra>.+))?
        \.(?P<extensao>[^.]+)$
        """,
        re.VERBOSE
    )

    if not os.path.exists(pasta_executaveis):
        print(f"Erro: A pasta '{pasta_executaveis}' não foi encontrada")
        return

    for pasta_atual, subpastas, lista_arquivos in os.walk(pasta_executaveis):

        for arquivo in lista_arquivos:
            caminho_completo_arquivo = os.path.join(pasta_atual, arquivo)
            versao_executavel = get_versao_arquivo(caminho_completo_arquivo)
            arquivo_identificado = padrao_nome_arquivo.match(arquivo)

            if arquivo_identificado:
                modulo_identificado = arquivo_identificado.group("modulo")
                primeiro_numero = arquivo_identificado.group("primeiro_numero")
                segundo_numero = arquivo_identificado.group("segundo_numero")

                status = (arquivo_identificado.group("status") or "Sem Status")
                info_extra = (arquivo_identificado.group("info_extra") or "")
                tipo_arquivo = arquivo_identificado.group("extensao")

                # Se existe segundo número: # MODULO_LOTE_TICKET
                if segundo_numero:
                    lote_semana = primeiro_numero
                    ticket = segundo_numero

                # Senão: # MODULO_TICKET
                else:
                    lote_semana = "Sem Lote"
                    ticket = primeiro_numero

                modulo_oficial, status_validacao = validar_modulo_json(
                    modulo_identificado,
                    lista_modulos_oficiais);
                
                linhas_relatorio.append([
                    pasta_atual,
                    arquivo,
                    modulo_identificado,
                    modulo_oficial,
                    status_validacao,
                    versao_executavel,
                    lote_semana,
                    ticket,
                    status,
                    info_extra,
                    tipo_arquivo,
                    "Padrão Reconhecido"
                ])

            else:
                linhas_relatorio.append([
                    pasta_atual,
                    arquivo,
                    "Desconhecido",
                    "Desconhecido",
                    "Não",
                    versao_executavel,
                    "N/A",
                    "N/A",
                    "N/A",
                    "",
                    os.path.splitext(arquivo)[1].replace('.', ''),
                    "Fora do Padrão"
                ])
    try:
        with open(
            arquivo_saida_csv,
            mode='w',
            newline='',
            encoding='utf-8-sig'
        ) as arquivo_csv:
            gravador_csv = csv.writer(
                arquivo_csv,
                delimiter=';'
            )

            gravador_csv.writerow([
                'Pasta de Origem',
                'Nome do Arquivo',
                'Módulo Extraído',
                'Módulo Oficial (De-Para)',
                'Módulo Validado no JSON?',
                'Versão do Arquivo',
                'Lote/Ordem (Nome)',
                'Ticket',
                'Status',
                'Informação Extra',
                'Extensão',
                'Observação'
            ])

            gravador_csv.writerows(linhas_relatorio)

        print(f"Sucesso! Relatório gerado em: "
              f"{os.path.abspath(arquivo_saida_csv)}")

        print(f"Total de {len(linhas_relatorio)} arquivos catalogados.")

    except Exception as e:
        print(f"Erro ao salvar o CSV: {e}")

if __name__ == "__main__":

    #Pasta que será analisada
    pasta_para_ler = r"C:\\Users\\andre\\Desktop\\Tiffany\\Pessoal\\tcc\\ExtratorPastas"

    #Arquivo JSON contendo os módulos oficiais
    arquivo_modulos_json = r"Modulos_prg.json"

    #Nome do CSV que será gerado
    caminho_csv_saida = "relatorio_tickets.csv"

    gerar_extracao_dados(
        pasta_para_ler,
        arquivo_modulos_json,
        caminho_csv_saida
    )