import pandas as pd
import glob
import os

caminho_pasta = input(r"Cole o caminho da pasta: ")
cabecalho = ["Ticket", "Descricao","Data Criacao", "Excluir", "Empresa", "Excluir2", "Operador", "Data Interacao", "Excluir3", "Status"] 

padrao = os.path.join(caminho_pasta, "*.xlsx")
arquivos = glob.glob(padrao)

if not arquivos:
    print("Nenhum arquivo encontrado.")
else:
    lista_df = []
    
    for f in arquivos:
        try:
            df_temporario = pd.read_excel(f, header=None, names=cabecalho)
            
            lista_df.append(df_temporario)
            print(f"Lido (sem cabeçalho original): {os.path.basename(f)}")
            
        except Exception as e:
            print(f"Erro ao processar {os.path.basename(f)}: {e}")

    if lista_df:
        df_final = pd.concat(lista_df, ignore_index=True) # Mescla todos os arquivos
         
        df_final.to_excel("relatorio_mesclado.xlsx", index=False) # Salva o resultado final
        print(f"\nSucesso! {len(lista_df)} arquivos mesclados com novos cabeçalhos.")