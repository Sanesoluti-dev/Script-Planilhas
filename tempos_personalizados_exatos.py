# -*- coding: utf-8 -*-
"""
SCRIPT PARA USAR TEMPOS PERSONALIZADOS EXATOS
=============================================

Este script usa os valores exatos de tempo de coleta que o usuário encontrou manualmente
para garantir que a planilha seja ajustada exatamente como ele fez manualmente.
"""

from ajustador_tempo_coleta import main

# Valores exatos fornecidos pelo usuário para o primeiro ponto
# Formato: [tempo_leitura_1, tempo_leitura_2, tempo_leitura_3]
tempos_personalizados = {
    'ponto_1': [
        359.8001,      # Primeira leitura
        359.851939,    # Segunda leitura  
        359.9921874    # Terceira leitura
    ]
}

print("🎯 USANDO TEMPOS PERSONALIZADOS EXATOS")
print("=" * 50)
print("Tempos fornecidos pelo usuário:")
for ponto, tempos in tempos_personalizados.items():
    print(f"   {ponto}:")
    for i, tempo in enumerate(tempos, 1):
        print(f"      Leitura {i}: {tempo} s")

print("\n" + "=" * 50)
print("Executando ajuste com tempos personalizados...")
print("=" * 50)

# Executa o ajuste com os tempos personalizados
main(tempos_personalizados) 