"""
Visualización de la Crisis de los Tulipanes (1634-1638)
Genera gráficos interactivos del dataset histórico
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# Configuración estilo
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['figure.figsize'] = (14, 10)
plt.rcParams['font.size'] = 11
plt.rcParams['axes.titlesize'] = 14
plt.rcParams['axes.labelsize'] = 12

# Colores por fase económica
COLORES_FASE = {
    'incipiente': '#2ecc71',   # verde
    'boom': '#3498db',         # azul
    'pico': '#e74c3c',         # rojo
    'pico_max': '#c0392b',     # rojo oscuro
    'crash_inicio': '#e67e22',  # naranja
    'crash': '#e74c3c',        # rojo
    'recovery': '#9b59b6',    # morado
    'estabilizado': '#1abc9c',   # turquesa
}

# Cargar datos
df = pd.read_csv('crisis_tulipanes.csv')

# Crear columna fecha combinada (año-mes)
df['fecha'] = df['año'] + (df['mes'] - 1) / 12

# ============================================================
# GRÁFICO 1: Precio vs Tiempo con fases
# ============================================================
fig, ax = plt.subplots(figsize=(14, 7))

# Scatter plot con colores por fase
for fase, color in COLORES_FASE.items():
    mask = df['estado_economico'] == fase
    if mask.any():
        ax.scatter(df[mask]['fecha'], df[mask]['precio_guilders'], 
                   c=color, s=100, label=fase.replace('_', ' ').title(), 
                   alpha=0.8, edgecolors='white', linewidth=1)

# Línea de tendencia
ax.plot(df['fecha'], df['precio_guilders'], 
        color='gray', alpha=0.5, linestyle='--', linewidth=1)

# Marcar eventos clave
ax.annotate('Precio pico: 4000 guilders', 
            xy=(1637.12, 4000), xytext=(1636.5, 4200),
            arrowprops=dict(arrowstyle='->', color='red'),
            fontsize=10, color='red')

ax.annotate('CRASH: -95%', 
            xy=(1637.2, 500), xytext=(1635.8, 800),
            arrowprops=dict(arrowstyle='->', color='orange'),
            fontsize=10, color='orange')

ax.set_xlabel('Año')
ax.set_ylabel('Precio (guilders)')
ax.set_title('Crisis de los Tulipanes: Precio vs Tiempo (1634-1638)', fontweight='bold')
ax.legend(loc='upper left', fontsize=10)
ax.set_xlim(1633.5, 1638.5)
ax.set_ylim(0, 4500)

# Añadir líneas divisorias de fases
ax.axvline(x=1635, color='gray', linestyle=':', alpha=0.5)
ax.axvline(x=1636, color='gray', linestyle=':', alpha=0.5)
ax.axvline(x=1637.1, color='gray', linestyle=':', alpha=0.5)

plt.tight_layout()
plt.savefig('grafico1_precio_vs_tiempo.png', dpi=150, bbox_inches='tight')
plt.close()

print("[OK] Grafico 1: Precio vs Tiempo guardado")

# ============================================================
# GRÁFICO 2: Volumen por Variedad
# ============================================================
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

# Por volumen total
volumen_variedad = df.groupby('variedad')['volumen_transado'].sum()
colors = ['#e74c3c', '#3498db', '#2ecc71']
axes[0].pie(volumen_variedad, labels=volumen_variedad.index, 
            autopct='%1.1f%%', colors=colors, startangle=90,
            explode=(0.05, 0, 0))
axes[0].set_title('Volumen Total por Variedad', fontweight='bold')

# Por precio promedio
precio_variedad = df.groupby('variedad')['precio_guilders'].mean()
axes[1].bar(precio_variedad.index, precio_variedad.values, color=colors)
axes[1].set_xlabel('Variedad')
axes[1].set_ylabel('Precio Promedio (guilders)')
axes[1].set_title('Precio Promedio por Variedad', fontweight='bold')

for i, v in enumerate(precio_variedad.values):
    axes[1].text(i, v + 50, f'{v:.0f}', ha='center', fontweight='bold')

plt.tight_layout()
plt.savefig('grafico2_variedad.png', dpi=150, bbox_inches='tight')
plt.close()

print("[OK] Grafico 2: Variedad guardado")

# ============================================================
# GRÁFICO 3: Evolución por Fase Económica
# ============================================================
fig, axes = plt.subplots(2, 2, figsize=(14, 10))

# 3a: Precio promedio por fase
fase_order = ['incipiente', 'boom', 'pico', 'pico_max', 'crash_inicio', 'crash', 'recovery', 'estabilizado']
precio_fase = df.groupby('estado_economico')['precio_guilders'].mean()
precio_fase = precio_fase.reindex([f for f in fase_order if f in precio_fase.index])
colors_fase = [COLORES_FASE[f] for f in precio_fase.index]
axes[0, 0].bar(range(len(precio_fase)), precio_fase.values, color=colors_fase)
axes[0, 0].set_xticks(range(len(precio_fase)))
axes[0, 0].set_xticklabels([f.replace('_', '\n') for f in precio_fase.index], rotation=45, ha='right')
axes[0, 0].set_ylabel('Precio Promedio (guilders)')
axes[0, 0].set_title('Precio Promedio por Fase', fontweight='bold')

# 3b: Volumen por fase
volumen_fase = df.groupby('estado_economico')['volumen_transado'].sum()
volumen_fase = volumen_fase.reindex([f for f in fase_order if f in volumen_fase.index])
axes[0, 1].bar(range(len(volumen_fase)), volumen_fase.values, color=colors_fase)
axes[0, 1].set_xticks(range(len(volumen_fase)))
axes[0, 1].set_xticklabels([f.replace('_', '\n') for f in volumen_fase.index], rotation=45, ha='right')
axes[0, 1].set_ylabel('Volumen Total')
axes[0, 1].set_title('Volumen Total por Fase', fontweight='bold')

# 3c: Tipo de comprador por fase
comprador_fase = df.groupby(['estado_economico', 'tipo_comprador']).size().unstack(fill_value=0)
comprador_fase = comprador_fase.reindex([f for f in fase_order if f in comprador_fase.index])
comprador_fase.plot(kind='bar', stacked=True, ax=axes[1, 0], 
                   colormap='Set2')
axes[1, 0].set_xticklabels([f.replace('_', '\n') for f in fase_order if f in comprador_fase.index], 
                          rotation=45, ha='right')
axes[1, 0].set_ylabel('Cantidad de Transacciones')
axes[1, 0].set_title(' Tipo de Comprador por Fase', fontweight='bold')
axes[1, 0].legend(title='Comprador', loc='upper right')

# 3d: Índice de precio (base 100)
indice = df.groupby('fecha')['indice_precio'].first()
axes[1, 1].fill_between(indice.index, indice.values, alpha=0.3, color='blue')
axes[1, 1].plot(indice.index, indice.values, color='blue', linewidth=2)
axes[1, 1].set_xlabel('Año')
axes[1, 1].set_ylabel('Índice (Base 100 = 10 guilders)')
axes[1, 1].set_title('Índice de Precio en el Tiempo', fontweight='bold')
axes[1, 1].set_yscale('log')
axes[1, 1].set_ylim(50, 50000)

plt.tight_layout()
plt.savefig('grafico3_fases.png', dpi=150, bbox_inches='tight')
plt.close()

print("[OK] Grafico 3: Fases guardado")

# ============================================================
# GRÁFICO 4: Mapa de Calor - Precio x Mes x Año
# ============================================================
fig, ax = plt.subplots(figsize=(12, 8))

# Pivot table: precio promedio por año y mes
pivot = df.pivot_table(values='precio_guilders', index='mes', columns='año', aggfunc='mean')
pivot = pivot.reindex(range(1, 13))

# heatmap
im = ax.imshow(pivot.values, cmap='YlOrRd', aspect='auto')

# Labels
ax.set_xticks(range(len(pivot.columns)))
ax.set_xticklabels(pivot.columns)
ax.set_yticks(range(len(pivot.index)))
ax.set_yticklabels([f'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'])
ax.set_xlabel('Año')
ax.set_ylabel('Mes')
ax.set_title('Mapa de Calor: Precio por Mes y Año', fontweight='bold')

# Colorbar
cbar = plt.colorbar(im, ax=ax)
cbar.set_label('Precio (guilders)')

# Anotar valores significativos
for i in range(len(pivot.index)):
    for j in range(len(pivot.columns)):
        val = pivot.values[i, j]
        if not np.isnan(val):
            color = 'white' if val > 1500 else 'black'
            ax.text(j, i, f'{val:.0f}', ha='center', va='center', 
                    color=color, fontsize=8)

plt.tight_layout()
plt.savefig('grafico4_calor.png', dpi=150, bbox_inches='tight')
plt.close()

print("[OK] Grafico 4: Mapa de Calor guardado")

# ============================================================
# RESUMEN ESTADÍSTICO
# ============================================================
print("\n" + "="*60)
print("RESUMEN DE VISUALIZACIONES GENERADAS")
print("="*60)

print("""
GRAFICOS CREADOS:

1. grafico1_precio_vs_tiempo.png
   => Evolucion del precio en el tiempo con colores por fase
   => Muestra el Boom (1635), Pico (1636) y Crash (1637)

2. grafico2_variedad.png
   => Pie: Volumen por variedad (Semper Augustus domina)
   => Bar: Precio promedio por variedad

3. grafico3_fases.png
   => 4 subgraficos:
     - Precio promedio por fase
     - Volumen por fase
     - Tipo de comprador por fase
     - Indice de precio (log scale)

4. grafico4_calor.png
   => Heatmap precio vs (mes x ano)
   => Muestra la concentracion del boom en 1636-1637
""")

print("\nESTADISTICAS CLAVE:")
print(f"   Precio máximo: {df['precio_guilders'].max():,} guilders")
print(f"   Volumen total: {df['volumen_transado'].sum():,} bulbos")
print(f"   Aumento máximo: {df['indice_precio'].max() / df['indice_precio'].min():.0f}x")
print(f"   Caída máxima: {-((df['precio_guilders'].max() - df['precio_guilders'].min()) / df['precio_guilders'].max() * 100):.0f}%")