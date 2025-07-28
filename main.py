# -*- coding: utf-8 -*-
"""
MAIN - ORQUESTRADOR DOS ALGORITMOS DE CORREÇÃO
===============================================

Este script executa os três algoritmos em sequência:
1. extrator_dados_certificado.py - Extrai dados do certificado (NÃO ALTERAR)
2. extrator_pontos_calibracao.py - Extrai pontos de calibração (PODEM SER ALTERADOS)
3. script_corrigido.py - Corrige parâmetros alteráveis mantendo o certificado

Cada script passa informações para o próximo através de arquivos JSON.
"""

import subprocess
import sys
import os
from datetime import datetime

def executar_script(script_name, descricao):
    """Executa um script Python e retorna True se bem-sucedido"""
    print(f"\n{'='*60}")
    print(f"EXECUTANDO: {script_name}")
    print(f"DESCRIÇÃO: {descricao}")
    print(f"{'='*60}")
    
    try:
        # Executa o script
        resultado = subprocess.run([sys.executable, script_name], 
                                 capture_output=True, text=True, encoding='latin-1')
        
        # Exibe a saída
        if resultado.stdout:
            print(resultado.stdout)
        
        if resultado.stderr:
            print("ERROS:")
            print(resultado.stderr)
        
        # Verifica se foi bem-sucedido
        if resultado.returncode == 0:
            print(f"SUCESSO: {script_name} executado com SUCESSO!")
            return True
        else:
            print(f"FALHA: {script_name} falhou com código de saída: {resultado.returncode}")
            return False
            
    except Exception as e:
        print(f"ERRO: Erro ao executar {script_name}: {e}")
        return False

def verificar_arquivos_json():
    """Verifica se os arquivos JSON necessários foram criados"""
    arquivos_necessarios = [
        "certificado.json",
        "pontos_calibracao.json", 
        "resultados_corrigidos.json"
    ]
    
    print(f"\n{'='*60}")
    print("VERIFICAÇÃO DOS ARQUIVOS JSON")
    print(f"{'='*60}")
    
    todos_existem = True
    for arquivo in arquivos_necessarios:
        if os.path.exists(arquivo):
            tamanho = os.path.getsize(arquivo)
            print(f"OK: {arquivo} - {tamanho} bytes")
        else:
            print(f"ERRO: {arquivo} - NAO ENCONTRADO")
            todos_existem = False
    
    return todos_existem

def main():
    """Função principal que orquestra a execução dos três algoritmos"""
    
    print("SISTEMA DE CORRECAO DE PARAMETROS ALTERAVEIS")
    print("="*60)
    print("ORQUESTRADOR DOS TRÊS ALGORITMOS")
    print("="*60)
    print(f"Início: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    # Lista dos scripts a serem executados em ordem
    scripts = [
        {
            "nome": "extrator_dados_certificado.py",
            "descricao": "Extrai dados da aba 'Emissão do Certificado' - VALORES QUE NÃO PODEM SER ALTERADOS"
        },
        {
            "nome": "extrator_pontos_calibracao.py", 
            "descricao": "Extrai pontos de calibração da aba 'Coleta de Dados' - VALORES QUE PODEM SER ALTERADOS"
        },
        {
            "nome": "script_corrigido.py",
            "descricao": "Corrige parâmetros alteráveis mantendo o certificado intacto"
        }
    ]
    
    # Executa cada script em sequência
    sucessos = 0
    for i, script in enumerate(scripts, 1):
        print(f"\n📋 PASSO {i}/3: {script['nome']}")
        
        if executar_script(script['nome'], script['descricao']):
            sucessos += 1
        else:
            print(f"\nFALHA NO PASSO {i}. Interrompendo execução.")
            break
    
    # Verifica se todos os scripts foram executados com sucesso
    if sucessos == 3:
        print(f"\n{'='*60}")
        print("SUCESSO: TODOS OS ALGORITMOS EXECUTADOS COM SUCESSO!")
        print(f"{'='*60}")
        
        # Verifica se os arquivos JSON foram criados
        if verificar_arquivos_json():
            print(f"\n📊 RESUMO FINAL:")
            print(f"   • 3/3 scripts executados com sucesso")
            print(f"   • 3/3 arquivos JSON criados")
            print(f"   • Sistema de correção completo")
            print(f"   • Certificado preservado")
            print(f"   • Parâmetros alteráveis corrigidos")
            print(f"   • Tempos de coleta igualados")
            
            print(f"\n📁 ARQUIVOS GERADOS:")
            print(f"   • certificado.json - Dados do certificado (não alterados)")
            print(f"   • pontos_calibracao.json - Dados de calibração (alteráveis)")
            print(f"   • resultados_corrigidos.json - Resultados finais corrigidos")
            
            print(f"\nFim: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
            print(f"SUCESSO: PROCESSO COMPLETO FINALIZADO COM SUCESSO!")
            
        else:
            print(f"\n⚠️  ATENÇÃO: Nem todos os arquivos JSON foram criados.")
            print(f"   Verifique se há erros nos scripts individuais.")
            
    else:
        print(f"\nFALHA NO PROCESSO")
        print(f"   • {sucessos}/3 scripts executados com sucesso")
        print(f"   • Verifique os erros acima e tente novamente")

if __name__ == "__main__":
    main() 