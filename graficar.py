import pandas as pd
import matplotlib.pyplot as plt
import os

# Ruta al archivo Excel
archivo_excel = os.path.join('Datos', 'ejemplo.xlsx')

# Crear carpeta para guardar las gráficas si no existe
os.makedirs('Generados', exist_ok=True)

# Nombres reales de las hojas para barras y pastel
hojas_barras = ['Barras1', 'Barras2', 'Barras3']
hojas_pastel = ['Pastel1', 'Pastel2', 'Pastel3']    

# Graficar barras
for nombre_hoja in hojas_barras:
    df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja, engine='openpyxl')
    eje_x = df.iloc[:, 1]
    eje_y = df.iloc[:, 2]
    fig, ax = plt.subplots()
    ax.bar(eje_x, eje_y, color="green", label=f"Datos de {nombre_hoja}")
    ax.set_xlabel("Eje X")
    ax.set_ylabel("Eje Y")
    ax.set_title(f"Gráfica de Barras - {nombre_hoja}")
    ax.legend()
    plt.xticks(rotation=45)
    fig.savefig(f"Generados/grafica_barras_{nombre_hoja}.png", bbox_inches="tight")
    plt.close(fig)
    print(f"Gráfica de barras de {nombre_hoja} guardada en 'Generados/grafica_barras_{nombre_hoja}.png'")

# Graficar pastel
for nombre_hoja in hojas_pastel:
    df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja, engine='openpyxl')
    etiquetas = df.iloc[:, 1]
    valores = df.iloc[:, 2]
    fig, ax = plt.subplots()
    ax.pie(valores, labels=etiquetas, autopct='%1.1f%%', startangle=90)
    ax.set_title(f"Gráfica de Pastel - {nombre_hoja}")
    fig.savefig(f"Generados/grafica_pastel_{nombre_hoja}.png", bbox_inches="tight")
    plt.close(fig)
    print(f"Gráfica de pastel de {nombre_hoja} guardada en 'Generados/grafica_pastel_{nombre_hoja}.png'")