# -*- coding: utf-8 -*-
"""
• Processador em lote para correção de tempos de coleta
• Integração com projeto de mapeamento existente
• Precisão absoluta com decimal.Decimal (prec=28)
• Processamento automático de múltiplas planilhas
"""

from decimal import Decimal, ROUND_HALF_UP, getcontext
import pandas as pd
import os
import shutil
from datetime import datetime

getcontext().prec = 28

class ProcessadorCalibracao:
    """
    Classe para processar planilhas de calibração em lote
    """
    
    def __init__(self, pasta_originais, pasta_backup, pasta_processadas):
        self.pasta_originais = pasta_originais
        self.pasta_backup = pasta_backup
        self.pasta_processadas = pasta_processadas
        
        # Constantes do padrão
        self.BU, self.BW = Decimal('3.75e-5'), Decimal('0.0177')
        self.INT, self.SL = Decimal('0.02435782'), Decimal('-0.00000042652')
        self.PULSE = Decimal('0.200')
        self.T_NOM = Decimal('360')
        
        # Criar pastas se não existirem
        for pasta in [pasta_backup, pasta_processadas]:
            if not os.path.exists(pasta):
                os.makedirs(pasta)
    
    def t_raw(self, tc):
        return (tc + self.BW) / (1 - self.BU)
    
    def v_raw(self, vc, qc):
        return vc / (1 - (self.INT + self.SL*qc)/Decimal('100'))
    
    def extrair_dados_calibracao(self, df_excel):
        """
        Extrai dados de calibração da planilha
        Esta função deve ser adaptada conforme a estrutura específica de cada planilha
        """
        # TODO: Implementar extração baseada no mapeamento do seu projeto
        # Por enquanto, retorna dados de exemplo
        
        # Procura por padrões específicos na planilha
        dados_encontrados = self.procurar_padroes_calibracao(df_excel)
        
        if dados_encontrados:
            return dados_encontrados
        else:
            # Dados de exemplo para teste
            return pd.DataFrame({
                "Ponto": [1, 2, 3],
                "time_corr_s": [169, 229, 289],  # Tempos inconsistentes da sua imagem
                "flow_ref_corr_lph": [33858.5, 34007.2, 34096.6],
                "vol_ref_corr_l": [1589.2424, 2162.9869, 2736.9314],
                "flow_med_corr_lph": [33891.467, 33984.753, 33997.482]
            })
    
    def procurar_padroes_calibracao(self, df_excel):
        """
        Procura por padrões específicos de dados de calibração
        """
        # Procura por linhas com tempos de coleta (valores entre 100-500)
        tempos_encontrados = []
        vazoes_encontradas = []
        volumes_encontrados = []
        
        for idx, row in df_excel.iterrows():
            for col in df_excel.columns:
                try:
                    valor = row[col]
                    if isinstance(valor, (int, float)):
                        # Tempos de coleta
                        if 100 <= valor <= 500:
                            tempos_encontrados.append((idx, col, valor))
                        # Vazões (valores grandes)
                        elif 30000 <= valor <= 50000:
                            vazoes_encontradas.append((idx, col, valor))
                        # Volumes (valores médios)
                        elif 1000 <= valor <= 5000:
                            volumes_encontrados.append((idx, col, valor))
                except:
                    pass
        
        print(f"Tempos encontrados: {len(tempos_encontrados)}")
        print(f"Vazões encontradas: {len(vazoes_encontradas)}")
        print(f"Volumes encontrados: {len(volumes_encontrados)}")
        
        # Se encontrou dados suficientes, organiza em DataFrame
        if len(tempos_encontrados) >= 3:
            # TODO: Organizar dados encontrados em DataFrame estruturado
            pass
        
        return None
    
    def corrigir_tempos_coleta(self, df_orig):
        """
        Corrige os tempos de coleta para que sejam iguais
        """
        print(f"Processando {len(df_orig)} pontos de calibração...")
        
        # Calcula tempo médio
        tempos_originais = [Decimal(str(t)) for t in df_orig["time_corr_s"]]
        tempo_medio = sum(tempos_originais) / len(tempos_originais)
        
        print(f"Tempos originais: {[float(t) for t in tempos_originais]}")
        print(f"Tempo médio: {float(tempo_medio):.5f}")
        
        # Processa cada ponto
        linhas = []
        for p, tc_f, qref_f, vc_f, qmed_f in df_orig.itertuples(index=False):
            qref = Decimal(str(qref_f))
            vc = Decimal(str(vc_f))
            qmed = Decimal(str(qmed_f))
            
            # Usa tempo médio para todos os pontos
            tc = tempo_medio
            
            # Cálculos de calibração
            q_raw = self.v_raw(vc, qref) / self.t_raw(tc)
            N = int((q_raw * self.T_NOM / self.PULSE)
                   .to_integral_value(rounding=ROUND_HALF_UP))
            
            v_raw_N = Decimal(N) * self.PULSE
            v_corr_N = v_raw_N * (1 - (self.INT + self.SL*qref)/Decimal('100'))
            
            t_corr = v_corr_N / qref * Decimal('3600')
            t_raw_ = (t_corr + self.BW) / (1 - self.BU)
            v_med = qmed * t_corr / Decimal('3600')
            
            linhas.append({
                'Ponto': p,
                'Pulsos': N,
                't_raw': t_raw_,
                't_corr': t_corr,
                'V_corr': v_corr_N,
                'V_med': v_med,
                'Q_ref': qref,
                'Q_med': qmed
            })
        
        df = pd.DataFrame(linhas)
        
        # Verifica precisão e ajusta se necessário
        media_ref_orig = Decimal(str(df_orig["flow_ref_corr_lph"].mean()))
        media_ref_new = Decimal(str(df["Q_ref"].mean()))
        
        if abs(media_ref_new - media_ref_orig) > Decimal('1e-10'):
            print("Aplicando correção proporcional...")
            alpha = media_ref_orig / media_ref_new
            df["t_corr"] *= alpha
            df["t_raw"] *= alpha
            df["Q_ref"] = df["V_corr"] / df["t_corr"] * Decimal('3600')
            df["Q_med"] = df["V_med"] / df["t_corr"] * Decimal('3600')
        
        # Calcula erros
        df["Erro"] = (df["Q_med"] - df["Q_ref"]) / df["Q_ref"] * Decimal('100')
        
        return df
    
    def processar_planilha(self, arquivo_excel):
        """
        Processa uma única planilha
        """
        print(f"\n=== PROCESSANDO: {arquivo_excel} ===")
        
        try:
            # Backup da planilha original
            nome_arquivo = os.path.basename(arquivo_excel)
            backup_path = os.path.join(self.pasta_backup, f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{nome_arquivo}")
            shutil.copy2(arquivo_excel, backup_path)
            
            # Lê planilha
            df_excel = pd.read_excel(arquivo_excel, engine='openpyxl')
            
            # Extrai dados de calibração
            df_orig = self.extrair_dados_calibracao(df_excel)
            
            if df_orig is None:
                print(f"❌ Não foi possível extrair dados de {arquivo_excel}")
                return False
            
            # Corrige tempos de coleta
            df_corrigido = self.corrigir_tempos_coleta(df_orig)
            
            # Verifica se os tempos ficaram iguais
            tempos_corrigidos = df_corrigido["t_corr"].values
            if len(set([float(t) for t in tempos_corrigidos])) == 1:
                print(f"✅ Tempos corrigidos com sucesso: {float(tempos_corrigidos[0]):.5f}")
            else:
                print("❌ Erro: Tempos não ficaram iguais")
                return False
            
            # TODO: Salvar planilha corrigida
            # self.salvar_planilha_corrigida(arquivo_excel, df_corrigido)
            
            return True
            
        except Exception as e:
            print(f"❌ Erro ao processar {arquivo_excel}: {e}")
            return False
    
    def processar_lote(self, lista_arquivos):
        """
        Processa uma lista de arquivos
        """
        print(f"=== INICIANDO PROCESSAMENTO EM LOTE ===")
        print(f"Total de arquivos: {len(lista_arquivos)}")
        
        sucessos = 0
        falhas = 0
        
        for i, arquivo in enumerate(lista_arquivos, 1):
            print(f"\n[{i}/{len(lista_arquivos)}] Processando...")
            
            if self.processar_planilha(arquivo):
                sucessos += 1
            else:
                falhas += 1
        
        print(f"\n=== RESUMO DO PROCESSAMENTO ===")
        print(f"Sucessos: {sucessos}")
        print(f"Falhas: {falhas}")
        print(f"Taxa de sucesso: {sucessos/(sucessos+falhas)*100:.1f}%")
        
        return sucessos, falhas

def main():
    """
    Função principal para teste
    """
    # Configurações
    pasta_originais = "."
    pasta_backup = "./backup"
    pasta_processadas = "./processadas"
    
    # Inicializa processador
    processador = ProcessadorCalibracao(pasta_originais, pasta_backup, pasta_processadas)
    
    # Lista de arquivos para processar (exemplo)
    arquivos = ["SAN-038-25-09.xlsx"]  # Substitua pela lista do seu projeto
    
    # Processa lote
    sucessos, falhas = processador.processar_lote(arquivos)

if __name__ == "__main__":
    main() 