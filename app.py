import pandas as pd
import os
import logging

# config log
logging.basicConfig(
    filename='log.txt',
    filemode='a', 
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    datefmt='%d/%m/%Y %H:%M:%S'
)

caminho_origem = r"H:\LOGISTICA FATURAMENTO POS VENDAS\PÓS VENDAS\PÓS VENDAS 2026.xlsx"
caminho_destino = r"H:\TI\POWER BI\LOGISTICA\AUTOMAÇÃO FECHAMENTO MENSAL\Fechamento Mensal.xlsx"

try:
    base_bruta = pd.read_excel(caminho_origem)
    # print(base_bruta.info())

    estados_map = {
        'AC': 'Acre', 'AL': 'Alagoas', 'AP': 'Amapá', 'AM': 'Amazonas', 
        'BA': 'Bahia', 'CE': 'Ceará', 'DF': 'Distrito Federal', 'ES': 'Espírito Santo',
        'GO': 'Goiás', 'MA': 'Maranhão', 'MT': 'Mato Grosso', 'MS': 'Mato Grosso do Sul', 
        'MG': 'Minas Gerais', 'PA': 'Pará', 'PB': 'Paraíba', 'PR': 'Paraná',
        'PE': 'Pernambuco', 'PI': 'Piauí', 'RJ': 'Rio de Janeiro', 'RN': 'Rio Grande do Norte',
        'RS': 'Rio Grande do Sul', 'RO': 'Rondônia', 'RR': 'Roraima', 'SC': 'Santa Catarina',
        'SP': 'São Paulo', 'SE': 'Sergipe', 'TO': 'Tocantins'
    }

    hoje = pd.Timestamp.now().normalize()
    data_limite = pd.Timestamp('2026-03-01')
    
    base_bruta = base_bruta[base_bruta['DATA DE ENTREGA'] >= data_limite]
    base_bruta = base_bruta.dropna(subset=['DATA DE ENTREGA'])
    
    base_bruta['mês'] = base_bruta['DATA DE ENTREGA'].dt.strftime('01/%m/%Y')
    base_bruta['estado'] = base_bruta['UF'].map(estados_map)

        # Se "DIAS P/ ENTREGA" for maior ou igual a 0, é considerado no prazo (retorna 1, senão 0)
    base_bruta['no_prazo'] = (base_bruta['DIAS P/ ENTREGA'] >= 0).astype(int)

    # Se "DIAS P/ ENTREGA" for menor que 0 (negativo), é considerado atraso (retorna 1, senão 0)
    base_bruta['atraso'] = (base_bruta['DIAS P/ ENTREGA'] < 0).astype(int)

    resumo_por_nf = base_bruta.groupby(['mês','UF', 'estado'])[['VALOR NF','VALOR CTE', 'no_prazo', 'atraso']].sum()
    resumo_por_nf['total_entregue'] = resumo_por_nf['no_prazo'] + resumo_por_nf['atraso']
    resumo_por_nf['delta_frete'] = (resumo_por_nf['VALOR CTE'] / resumo_por_nf['VALOR NF']) * 100
    resumo_por_nf['kpi ns'] = (resumo_por_nf['no_prazo'] / resumo_por_nf['total_entregue']) * 100

    resumo_final = resumo_por_nf.reset_index()

    with pd.ExcelWriter(caminho_destino, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        resumo_final.to_excel(writer, sheet_name='POR UF 2026', index=False)

    # Registro no log
    nome_arquivo = os.path.basename(caminho_destino)
    
    mensagem = f"Os dados foram consolidados com sucesso na sua planilha {nome_arquivo}."
    logging.info(mensagem)
    print(mensagem)

except Exception as e:
    logging.error(f"Erro na automação: {e}")
    print(f"Ocorreu um erro. Verifique o arquivo de log.")