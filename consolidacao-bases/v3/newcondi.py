import pandas as pd
import os

# === CONFIGURAÇÕES ===

# Caminho da pasta onde estão os arquivos .xlsx (caso queira alterar aqui de acordo com a data ou projeto)
pasta_bases = r'C:\Users\roberto.junior\Documents\Remarcados\11.08'

# Valor mínimo do estoque (condição ajustável) <<<<<<<<<<<<<<
estoque_minimo_remarcado = 1   # Altere aqui o valor mínimo para REMARCADO
estoque_minimo_nao_remarcado = 8  # Altere aqui o valor mínimo para NÃO REMARCADO

# Lista automaticamente todos os arquivos .xlsx da pasta (NÂO ALTERAR)
arquivos = [f for f in os.listdir(pasta_bases) if f.endswith('.xlsx')]

try:
    print("🔍 Etapa 1: Iniciando leitura das bases...")

    # === LEITURA DAS BASES ===
    bases = []
    for nome in arquivos:
        caminho = os.path.join(pasta_bases, nome)
        print(f"📄 Lendo arquivo: {nome}")
        df = pd.read_excel(caminho, engine='openpyxl')

        # Verifica se as colunas existem
        if 'Status Remarcado' not in df.columns:
            raise ValueError(f"❌ Coluna 'Status Remarcado' não encontrada no arquivo: {nome}")
        if 'Estoque Atual' not in df.columns:
            raise ValueError(f"❌ Coluna 'Estoque Atual' não encontrada no arquivo: {nome}")

        # Padroniza os valores da coluna Status Remarcado
        df['Status Remarcado'] = df['Status Remarcado'].astype(str).str.strip()

        # Remove linhas com 'Aumento de Preço'
        df = df[df['Status Remarcado'] != 'Aumento de Preço']

        bases.append(df)

    print("🔗 Etapa 2: Unificando todas as bases...")
    df_total = pd.concat(bases, ignore_index=True)

    # === APLICAÇÃO DAS REGRAS ===

    print("📌 Etapa 3: Aplicando regras de filtragem...")

    # Regra 1: Status Remarcado = "Remarcado" e Estoque Atual >= estoque_minimo_remarcado
    regra1 = df_total[
        (df_total['Status Remarcado'] == 'Remarcado') &
        (df_total['Estoque Atual'] >= estoque_minimo_remarcado)
    ]

    # Regra 2: Status Remarcado = "Não Remarcado" e Estoque Atual >= estoque_minimo_nao_remarcado
    regra2 = df_total[
        (df_total['Status Remarcado'] == 'Não Remarcado') &
        (df_total['Estoque Atual'] >= estoque_minimo_nao_remarcado)
    ]

    # Junta os dois resultados
    df_final = pd.concat([regra1, regra2], ignore_index=True)

    print(f"📊 Total de linhas finais: {len(df_final)}")

    # === EXPORTA A BASE FINAL ===

    print("💾 Etapa 4: Salvando arquivo final...")
    saida = os.path.join(pasta_bases, 'base_final_consolidada.xlsx')
    df_final.to_excel(saida, index=False)

    print(f'\n✅ Base final consolidada salva em:\n{saida}')

except Exception as e:
    print("❌ Ocorreu um erro durante a execução:")
    print(str(e))
