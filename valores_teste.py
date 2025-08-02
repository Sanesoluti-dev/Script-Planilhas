from decimal import Decimal

# Valores fixos e determinísticos para testes
# ESTRATÉGIA HÍBRIDA: Valores principais + Fallback para casos extremos

# VALORES PRINCIPAIS: 239.800000 - 240.200000 (400 valores)
valores_principais = []

# Valores menores que 240.000 (239.800 a 239.999)
for i in range(200):
    valor = Decimal('239.800000') + Decimal(str(i * 0.001000))
    valores_principais.append(valor)

# Valores maiores que 240.000 (240.001 a 240.200)
for i in range(200):
    valor = Decimal('240.001000') + Decimal(str(i * 0.001000))
    valores_principais.append(valor)

# VALORES DE FALLBACK: 239.600000 - 239.800000 (200 valores) - para casos extremos
valores_fallback = []

# Valores de fallback (239.600 a 239.799)
for i in range(200):
    valor = Decimal('239.600000') + Decimal(str(i * 0.001000))
    valores_fallback.append(valor)

# Combina todos os valores
valores_base = valores_principais + valores_fallback

# Ordena para facilitar a busca
valores_base.sort()

print(f"✅ Valores de teste carregados: {len(valores_base)} valores")
print(f"   • Valores principais: {len(valores_principais)} (239.800-240.200)")
print(f"   • Valores fallback: {len(valores_fallback)} (239.600-239.800)")
print(f"   • Range total: {float(min(valores_base)):.6f} a {float(max(valores_base)):.6f}")
print(f"   • Incremento: 0.001000")
print(f"   • Valores < 240.000: {len([v for v in valores_base if v < Decimal('240.000000')])}")
print(f"   • Valores > 240.000: {len([v for v in valores_base if v > Decimal('240.000000')])}")
print(f"   • Precisão: 6 casas decimais significativas")
print(f"   • Estratégia: Híbrida (principais + fallback)") 