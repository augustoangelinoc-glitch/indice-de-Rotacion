# 🔄 Índice de Rotación de Inventarios

App en Streamlit para calcular el **Índice de Rotación de Inventarios** por
almacén y por material a partir de tus archivos de Inventario Inicial,
Inventario Final y Salidas.

## ¿Qué calcula?

- **Índice de Rotación (periodo)** = Valor de Salidas / Valor Promedio de Inventario
- **Valor Promedio** = (Valor Inicial + Valor Final) / 2
- **Rotación Anualizada** = Índice de Rotación × (365 / días del periodo)
- **Días de Inventario** = días del periodo / Índice de Rotación
- **Materiales Sin Rotación**: stock con inventario pero sin ninguna salida (dinero inmovilizado)
- **Clasificación ABC** de materiales por valor de salidas (Pareto 80/95 por defecto, ajustable)

Todo se calcula a nivel (Almacén, Código), porque un mismo código puede tener
comportamientos de rotación distintos en cada almacén.

## 1. Subir este proyecto a GitHub

```bash
# Dentro de la carpeta del proyecto (donde está este README)
git init
git add .
git commit -m "Índice de Rotación de Inventarios - primera versión"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/TU_REPOSITORIO.git
git push -u origin main
```

Si ya tienes un repositorio (por ejemplo el de "inventario"), puedes copiar
esta carpeta completa dentro de él como una subcarpeta, ej. `rotacion/`, y
subirla con los mismos comandos `git add` / `commit` / `push`.

## 2. Ejecutarlo en tu computadora

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abrirá automáticamente en tu navegador (normalmente en `http://localhost:8501`).

## 3. Ejecutarlo gratis en la nube (Streamlit Community Cloud)

1. Entra a https://share.streamlit.io/ e inicia sesión con tu cuenta de GitHub.
2. Clic en **"New app"**.
3. Elige tu repositorio, la rama (`main`) y como **Main file path** escribe `app.py`
   (o `rotacion/app.py` si lo pusiste en una subcarpeta).
4. Clic en **Deploy**. En 1-2 minutos tendrás un link público para compartir.

## 4. Archivos que debes subir en la app

La app pide 3 archivos obligatorios y 1 opcional. Dentro de la app, en el
panel **"📋 Formato esperado de cada archivo"**, hay ejemplos y un botón para
descargar las 4 plantillas en un solo Excel. También están guardadas sueltas
en la carpeta [`plantillas/`](plantillas) de este repositorio:

| Archivo | Obligatorio | Columnas mínimas |
|---|---|---|
| `Plantilla_Inventario_Inicial.xlsx` | Sí | Almacén, Código, Sistema (Qty) |
| `Plantilla_Inventario_Final.xlsx` | Sí | Almacén, Código, Sistema (Qty) |
| `Plantilla_Salidas.xlsx` | Sí | Almacén, Material/Código, Fecha, Unidades |
| `Plantilla_Costos_Unitarios.xlsx` | No | Código, Costo Unitario Kardex |

La app **detecta automáticamente** los nombres de columna aunque no coincidan
exactamente con la plantilla (ej. reconoce "Cod.", "SKU", "Código" o "Item"
como la misma columna de código).

## 5. Detalles importantes (cosas que se corrigieron/mejoraron)

- **Los códigos nunca pierden ceros a la izquierda.** El código lee siempre
  la columna de Código/Material como texto, así que `00854` se mantiene como
  `00854` y no se convierte en `854`. Aun así, si tu Excel de origen ya
  guarda esa celda como número, revisa que la columna "Código" esté
  formateada como **Texto** antes de exportarla, porque ahí el dato ya se
  pierde antes de llegar a la app.
- **La descripción del material siempre aparece.** Si un código no tiene
  descripción en un archivo (por ejemplo en Inventario Final), la app la
  busca en los otros archivos cargados (Inventario Inicial, Salidas) antes
  de mostrar el reporte. Si no la encuentra en ninguno, muestra
  "(Sin descripción registrada)" en vez de dejarlo vacío.
- **Cruce correcto entre Código de Almacén y Nombre de Almacén.** Si tu
  archivo de Salidas trae tanto el código del almacén ("A1421") como su
  nombre completo ("PRINCIPAL PLUZ - SJL") en columnas separadas, la app usa
  siempre el **código** para cruzar contra Inventario (que es el mismo
  formato ahí), y solo usa el nombre para que las tablas se vean más claras.
  Antes esto se cruzaba por el nombre y duplicaba almacenes sin cruzar bien
  los datos — ya está corregido.
- **Costo unitario automático.** Si no subes el archivo de Costos
  Unitarios, la app toma el costo Kardex del Inventario Final, y si el
  material no aparece ahí, del Inventario Inicial.
- **Movimientos sin costo.** Si algún movimiento de salida no tiene costo
  disponible en ningún archivo, se excluye del valor de salidas (no del
  conteo de unidades) y la app te avisa cuántos fueron, para que sepas que
  el índice de esos almacenes podría estar levemente subestimado.

## 6. Estructura del proyecto

```
├── app.py            # Interfaz Streamlit (lo que ve el usuario)
├── logic.py           # Toda la lógica de cálculo (sin nada de interfaz)
├── requirements.txt   # Dependencias
├── plantillas/         # Plantillas de ejemplo listas para descargar
└── README.md
```

`logic.py` está separado de `app.py` a propósito: así puedes probar los
cálculos con `python3` directamente (sin abrir Streamlit) y reutilizar las
funciones si más adelante agregas otro reporte.
