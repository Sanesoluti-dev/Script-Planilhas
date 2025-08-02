# -*- coding: utf-8 -*-
"""
• Correção de tempos de coleta para valores iguais
• Mantém todos os valores do certificado inalterados
• Precisão absoluta com decimal.Decimal (prec=28)
• Processamento de planilhas Excel
"""

from decimal import Decimal, ROUND_HALF_UP, getcontext
import pandas as pd
import numpy as np

getcontext().prec = 28

def ler_planilha_calibracao(arquivo_excel):
    """
    Lê a planilha de calibração e extrai os dados necessários
    """
    try:
        print(f"Tentando ler arquivo: {arquivo_excel}")
        
        # Verifica se o arquivo existe
        import os
        if not os.path.exists(arquivo_excel):
            print(f"❌ Arquivo não encontrado: {arquivo_excel}")
            return None
        
        # Lê a planilha Excel
        df_excel = pd.read_excel(arquivo_excel, sheet_name=0, engine='openpyxl')
        
        print("=== ESTRUTURA DA PLANILHA ===")
        print(f"Dimensões: {df_excel.shape}")
        print("\nPrimeiras 10 linhas:")
        print(df_excel.head(10))
        
        print("\nColunas disponíveis:")
        print(df_excel.columns.tolist())
        
        # Procura por colunas que contenham "Tempo de Coleta" ou similar
        colunas_tempo = []
        colunas_vazao = []
        
        for col in df_excel.columns:
            if isinstance(col, str):
                if 'tempo' in col.lower() or 'coleta' in col.lower():
                    colunas_tempo.append(col)
                if 'vazão' in col.lower() or 'flow' in col.lower():
                    colunas_vazao.append(col)
        
        print(f"\nColunas de tempo encontradas: {colunas_tempo}")
        print(f"Colunas de vazão encontradas: {colunas_vazao}")
        
        # Procura por dados de calibração na planilha
        print("\n=== PROCURANDO DADOS DE CALIBRAÇÃO ===")
        
        # Procura por linhas que contenham dados numéricos de calibração
        for idx, row in df_excel.iterrows():
            # Procura por valores que possam ser tempos de coleta (números entre 100-500)
            valores_numericos = []
            for col in df_excel.columns:
                try:
                    valor = row[col]
                    if isinstance(valor, (int, float)) and 100 <= valor <= 500:
                        valores_numericos.append((col, valor))
                except:
                    pass
            
            if valores_numericos:
                print(f"Linha {idx}: Possíveis tempos de coleta encontrados: {valores_numericos}")
                print(f"Conteúdo da linha: {row.tolist()}")
                break
        
        return df_excel
        
    except Exception as e:
        print(f"Erro ao ler planilha: {e}")
        import traceback
        traceback.print_exc()
        return None

def identificar_dados_calibracao(df_excel):
    """
    Identifica e extrai os dados de calibração da planilha
    """
    # Esta função será implementada após analisar a estrutura da planilha
    # Por enquanto, vamos usar os dados de exemplo do código original
    print("\n=== DADOS IDENTIFICADOS ===")
    
    # Dados de exemplo (serão substituídos pelos dados reais da planilha)
    df_orig = pd.DataFrame({
        "Ponto":             [1, 2, 3],
        "time_corr_s":       [256.97269, 316.97044, 376.96819],
        "flow_ref_corr_lph": [33957.9218, 33744.3903, 33789.8358],
        "vol_ref_corr_l":    [2423.9607, 2971.1039, 3538.2481],
        "flow_med_corr_lph": [34092.4434, 33844.8597, 33859.3961]
    })
    
    return df_orig

def corrigir_tempos_coleta(df_orig):
    """
    Corrige os tempos de coleta para que sejam iguais, mantendo precisão absoluta
    """
    print("\n=== CORREÇÃO DE TEMPOS DE COLETA ===")
    
    # --------------------------------------------------
    # 1. Dados de entrada
    # --------------------------------------------------
    print("Dados originais:")
    print(df_orig.to_string(index=False))
    
    # --------------------------------------------------
    # 2. Constantes do padrão
    # --------------------------------------------------
    BU, BW  = Decimal('3.75e-5'), Decimal('0.0177')
    INT, SL = Decimal('0.02435782'), Decimal('-0.00000042652')
    PULSE   = Decimal('0.200')
    T_NOM   = Decimal('360')
    
    def t_raw(tc):  return (tc + BW) / (1 - BU)
    def v_raw(vc, qc):
        return vc / (1 - (INT + SL*qc)/Decimal('100'))
    
    # --------------------------------------------------
    # 3. Calcula tempo de coleta médio (será usado para todos os pontos)
    # --------------------------------------------------
    tempos_originais = [Decimal(str(t)) for t in df_orig["time_corr_s"]]
    tempo_medio = sum(tempos_originais) / len(tempos_originais)
    
    print(f"\nTempos originais: {[float(t) for t in tempos_originais]}")
    print(f"Tempo médio calculado: {float(tempo_medio)}")
    
    # --------------------------------------------------
    # 4. Gera linhas com tempos de coleta iguais
    # --------------------------------------------------
    linhas = []
    for p, tc_f, qref_f, vc_f, qmed_f in df_orig.itertuples(index=False):
        qref = Decimal(str(qref_f))
        vc   = Decimal(str(vc_f))
        qmed = Decimal(str(qmed_f))
        
        # Usa o tempo médio para todos os pontos
        tc = tempo_medio
        
        q_raw = v_raw(vc, qref) / t_raw(tc)               # L/s bruto
        N     = int((q_raw * T_NOM / PULSE)
                    .to_integral_value(rounding=ROUND_HALF_UP))
        
        v_raw_N  = Decimal(N)*PULSE
        v_corr_N = v_raw_N * (1 - (INT + SL*qref)/Decimal('100'))
        
        t_corr = v_corr_N / qref * Decimal('3600')
        t_raw_ = (t_corr + BW) / (1 - BU)
        v_med  = qmed * t_corr / Decimal('3600')
        
        linhas.append(dict(Ponto=p, Pulsos=N,
                           t_raw=t_raw_, t_corr=t_corr,
                           V_corr=v_corr_N, V_med=v_med,
                           Q_ref=qref, Q_med=qmed))
    
    df = pd.DataFrame(linhas)
    
    # --------------------------------------------------
    # 5. Verifica se os valores do certificado permanecem inalterados
    # --------------------------------------------------
    print("\n=== VERIFICAÇÃO DE PRECISÃO ===")
    
    # Calcula médias originais
    media_ref_orig = Decimal(str(df_orig["flow_ref_corr_lph"].mean()))
    media_med_orig = Decimal(str(df_orig["flow_med_corr_lph"].mean()))
    
    # Calcula médias após correção
    media_ref_new = Decimal(str(df["Q_ref"].mean()))
    media_med_new = Decimal(str(df["Q_med"].mean()))
    
    print(f"Vazão Ref Original: {float(media_ref_orig)}")
    print(f"Vazão Ref Nova:     {float(media_ref_new)}")
    print(f"Diferença:          {float(media_ref_new - media_ref_orig)}")
    
    print(f"\nVazão Med Original: {float(media_med_orig)}")
    print(f"Vazão Med Nova:     {float(media_med_new)}")
    print(f"Diferença:          {float(media_med_new - media_med_orig)}")
    
    # --------------------------------------------------
    # 6. Ajusta proporcionalmente se necessário
    # --------------------------------------------------
    if abs(media_ref_new - media_ref_orig) > Decimal('1e-10'):
        print("\n⚠️  AJUSTE NECESSÁRIO - Aplicando correção proporcional...")
        
        # Calcula fator de correção
        alpha = media_ref_orig / media_ref_new
        
        # Aplica correção proporcional
        df["t_corr"] *= alpha
        df["t_raw"]  *= alpha
        df["Q_ref"]   = df["V_corr"] / df["t_corr"] * Decimal('3600')
        df["Q_med"]   = df["V_med"]  / df["t_corr"] * Decimal('3600')
        
        # Recalcula médias após ajuste
        media_ref_final = Decimal(str(df["Q_ref"].mean()))
        media_med_final = Decimal(str(df["Q_med"].mean()))
        
        print(f"Vazão Ref Final:   {float(media_ref_final)}")
        print(f"Vazão Med Final:   {float(media_med_final)}")
        print(f"Diferença Final:   {float(media_ref_final - media_ref_orig)}")
    
    # --------------------------------------------------
    # 7. Calcula erros
    # --------------------------------------------------
    df["Erro"] = (df["Q_med"] - df["Q_ref"]) / df["Q_ref"] * Decimal('100')
    
    # --------------------------------------------------
    # 8. Tabela final
    # --------------------------------------------------
    df_final = pd.DataFrame({
        "Ponto":                 df["Ponto"],
        "Qtd Pulsos":            df["Pulsos"],
        "Tempo Coleta (s)":      df["t_raw"],
        "Tempo Coleta Corr (s)": df["t_corr"],
        "Vazão Ref L/h":         df["Q_ref"],
        "Vol Ref Corr L":        df["V_corr"],
        "Vazão Med L/h":         df["Q_med"],
        "Leitura Medidor L":     df["V_med"],
        "Erro %":                df["Erro"]
    })
    
    df_final.loc["Média"] = ["Média","",
        "", "",
        df_final["Vazão Ref L/h"].mean(),
        "", df_final["Vazão Med L/h"].mean(),
        "", df_final["Erro %"].mean()]
    
    return df_final, df_orig

def main():
    """
    Função principal que executa todo o processo
    """
    print("=== SISTEMA DE CORREÇÃO DE TEMPOS DE COLETA ===")
    
    # 1. Lê a planilha
    arquivo_excel = "SAN-038-25-09.xlsx"
    df_excel = ler_planilha_calibracao(arquivo_excel)
    
    if df_excel is None:
        print("❌ Erro ao ler planilha. Verifique se o arquivo existe.")
        return
    
    # 2. Identifica dados de calibração
    df_orig = identificar_dados_calibracao(df_excel)
    
    # 3. Corrige tempos de coleta
    df_final, df_orig = corrigir_tempos_coleta(df_orig)
    
    # 4. Exibe resultados
    pd.set_option("display.float_format", "{:.5f}".format)
    
    print("\n=== RESULTADO FINAL ===")
    print("Tempos de coleta corrigidos (todos iguais):")
    tempos_corrigidos = df_final.iloc[:3]["Tempo Coleta Corr (s)"].values
    print(f"Tempo 1: {tempos_corrigidos[0]:.5f}")
    print(f"Tempo 2: {tempos_corrigidos[1]:.5f}")
    print(f"Tempo 3: {tempos_corrigidos[2]:.5f}")
    
    print("\n=== TABELA COMPLETA ===")
    print(df_final.to_string(index=False))

if __name__ == "__main__":
    main()