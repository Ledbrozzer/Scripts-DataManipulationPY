# pip install fuzzywuzzy
#pip install python-Levenshtein  # Optional, to Increment Performance
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

# Paths t/t-Excel Files
caminho_valor_contrato = "Valor Contrato.xlsx"
caminho_valor_faturado = "Valor Faturado.xlsx"

# Read both ExcelFiles
contrato_df = pd.read_excel(caminho_valor_contrato)
faturado_df = pd.read_excel(caminho_valor_faturado)

# List t/Armazen Results
resultados = []

# Iterar on each cliente in "Valor Contrato"
for _, linha_contrato in contrato_df.iterrows():
    cliente_contrato = linha_contrato['Cliente']
    valor_contrato = linha_contrato['Valor de Contrato']
    
    # Find best match t/t-cliente on ExcelSheet "Valor Faturado"
    melhor_match = process.extractOne(cliente_contrato, faturado_df['Cliente'], scorer=fuzz.token_sort_ratio)
    
    if melhor_match and melhor_match[1] >= 75:  # Similarity of (AtLeast) 75%
        cliente_faturado = melhor_match[0]  # Name match from faturado
        valor_faturado = faturado_df.loc[faturado_df['Cliente'] == cliente_faturado, 'Valor Faturado'].values[0]
        
        # Calc dif between-values
        diferenca = valor_faturado - valor_contrato
        
        # Add infos t/result
        resultados.append({
            'Cliente': cliente_contrato,
            'Diferença': diferenca,
            'Valor Faturado': valor_faturado,
            'Valor de Contrato': valor_contrato
        })

# Create DataFrame w/results
resultado_final = pd.DataFrame(resultados)

# Reorganize columns t/t-order wanted
resultado_final = resultado_final[['Cliente', 'Diferença', 'Valor Faturado', 'Valor de Contrato']]

# Order data by column "Diferença" in DESC
resultado_final.sort_values(by='Diferença', ascending=False, inplace=True)

# Save results in a new Excel Sheet
#caminho_saida = r"#:\#\#\#\Result_Dif_Order.xlsx"
caminho_saida = "Result_Dif_Order.xlsx"
resultado_final.to_excel(caminho_saida, index=False)

print(f"Arquivo criado com sucesso: {caminho_saida}")
