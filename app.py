# -*- coding: utf-8 -*-
"""
app.py
------
Interfaz Streamlit del "Índice de Rotación de Inventarios PLUZ".

Cómo ejecutarlo localmente:
    pip install -r requirements.txt
    streamlit run app.py

Cómo subirlo a GitHub y correrlo en Streamlit Community Cloud:
    Ver instrucciones detalladas en README.md.
"""

import os
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

from logic import (
    load_inventario_file,
    load_salidas_file,
    load_costos_file,
    build_master_lookup,
    build_costos_unitarios,
    compute_rotacion,
    compute_materiales_sin_rotacion,
    compute_abc,
    build_resumen,
    build_export_excel,
)

st.set_page_config(page_title="Índice de Rotación de Inventarios PLUZ", page_icon="🔄", layout="wide")

# ---------------------------------------------------------------------------
# Estado de sesión: guarda los resultados entre clics para no recalcular
# ---------------------------------------------------------------------------
if "resultado" not in st.session_state:
    st.session_state.resultado = None

st.title("🔄 Índice de Rotación de Inventarios — PLUZ")
st.caption(
    "Sube Inventario Inicial, Inventario Final y Salidas (Costos Unitarios es opcional) "
    "y la app calcula el índice de rotación por almacén y por material."
)

# ---------------------------------------------------------------------------
# Sidebar — parámetros configurables
# ---------------------------------------------------------------------------
st.sidebar.title("⚙️ Parámetros del cálculo")
umbral_rotacion_alta = st.sidebar.number_input(
    "Rotación anualizada mínima para 'Rotación alta' (veces/año)",
    min_value=0.0, value=4.0, step=0.5,
    help="Un material con rotación anualizada por debajo de este valor se clasifica como 'Rotación baja'.",
)
st.sidebar.markdown("**Clasificación ABC (Pareto) de salidas**")
umbral_a = st.sidebar.slider("Corte para Clase A (% acumulado)", 50, 95, 80) / 100
umbral_b = st.sidebar.slider("Corte para Clase B (% acumulado)", int(umbral_a * 100) + 1, 99, 95) / 100

st.sidebar.markdown("---")
st.sidebar.caption(
    "Los códigos con ceros a la izquierda (ej. '00854') se conservan tal cual: "
    "se leen siempre como texto para que no se pierdan al abrir el Excel."
)

# ---------------------------------------------------------------------------
# 1) Formato esperado + plantillas de ejemplo
# ---------------------------------------------------------------------------
with st.expander("📋 Formato esperado de cada archivo (haz clic para ver ejemplos y descargar plantillas)"):
    t1, t2, t3, t4 = st.tabs(["Inventario Inicial", "Inventario Final", "Salidas", "Costos Unitarios (opcional)"])

    ejemplo_inv = pd.DataFrame({
        "Periodo": ["Inicial", "Inicial"],
        "Almacen": ["A1421", "A1421"],
        "Codigo": ["150501", "00854"],
        "Descripcion": ["SECC.TRIP.HORZ.220V.250A.P.FUS.NH", "CURVA PVC SAP 3/4 X 90"],
        "U.Medida": ["UND", "UND"],
        "Sistema (Qty)": [4, 145],
        "Familia": ["MATERIAL CONSIGNACION", "MATERIAL CONSIGNACION"],
        "Costo Kardex (S/)": [443.17, 12.35],
        "Valor Inventario (S/)": [1772.69, 1790.75],
    })
    with t1:
        st.write("Una fila por código de material en el almacén, al **inicio** del periodo.")
        st.dataframe(ejemplo_inv, hide_index=True, use_container_width=True)
        st.caption("Obligatorias: Almacén, Código, Sistema (Qty). El resto se completa automáticamente si falta.")
    with t2:
        st.write("Mismas columnas que Inventario Inicial, pero con el saldo al **final** del periodo.")
        st.dataframe(ejemplo_inv.assign(Periodo="Final"), hide_index=True, use_container_width=True)

    ejemplo_sal = pd.DataFrame({
        "Fecha": ["2025-09-30", "2025-10-02"],
        "Cod.Almacen": ["A1421", "A1421"],
        "Almacen": ["A1421", "A1421"],
        "Material": ["130495", "00854"],
        "Descripcion": ["INT.TERMG.2P 50A 4.5KA", "CURVA PVC SAP 3/4 X 90"],
        "Grupo Material": ["MATERIAL CONSIGNACION", "MATERIAL CONSIGNACION"],
        "U.M.": ["UND", "UND"],
        "Unidades": [10, 25],
        "Costo Unitario (S/)": [0, 12.35],
        "Valor Salida (S/)": [0, 308.75],
    })
    with t3:
        st.write("Una fila por movimiento de salida (despacho/consumo).")
        st.dataframe(ejemplo_sal, hide_index=True, use_container_width=True)
        st.caption(
            "Obligatorias: Almacén, Material/Código, Fecha, Unidades. "
            "Si 'Valor Salida' viene vacío o en 0, se calcula automáticamente con el Costo Unitario "
            "(del mismo archivo o del archivo de Costos Unitarios / Inventarios)."
        )

    ejemplo_costos = pd.DataFrame({
        "Codigo": ["150501", "00854"],
        "Costo Unitario Kardex (S/)": [443.17, 12.35],
        "Fuente": ["Inv. Final", "Inv. Final"],
    })
    with t4:
        st.write(
            "Opcional. Si no lo subes, el costo unitario de cada código se toma automáticamente "
            "del Inventario Final (y del Inicial si el código no aparece en el Final)."
        )
        st.dataframe(ejemplo_costos, hide_index=True, use_container_width=True)

    # Botón para descargar las 4 plantillas juntas en un solo Excel de ejemplo
    plantilla_path = "/tmp/plantillas_indice_rotacion.xlsx"
    with pd.ExcelWriter(plantilla_path, engine="openpyxl") as writer:
        ejemplo_inv.to_excel(writer, sheet_name="Inventario_Inicial", index=False)
        ejemplo_inv.assign(Periodo="Final").to_excel(writer, sheet_name="Inventario_Final", index=False)
        ejemplo_sal.to_excel(writer, sheet_name="Salidas", index=False)
        ejemplo_costos.to_excel(writer, sheet_name="Costos_Unitarios", index=False)
    with open(plantilla_path, "rb") as f:
        st.download_button(
            "📥 Descargar las 4 plantillas de ejemplo (Excel)",
            data=f.read(),
            file_name="plantillas_indice_rotacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---------------------------------------------------------------------------
# 2) Carga de archivos
# ---------------------------------------------------------------------------
st.subheader("1️⃣ Sube tus archivos")
c1, c2 = st.columns(2)
with c1:
    inv_inicial_file = st.file_uploader("Inventario Inicial (.xlsx)", type=["xlsx", "xls"], key="inv_ini")
    salidas_file = st.file_uploader("Salidas (.xlsx)", type=["xlsx", "xls"], key="sal")
with c2:
    inv_final_file = st.file_uploader("Inventario Final (.xlsx)", type=["xlsx", "xls"], key="inv_fin")
    costos_file = st.file_uploader("Costos Unitarios (.xlsx) — opcional", type=["xlsx", "xls"], key="costos")

st.subheader("2️⃣ Periodo analizado")
p1, p2 = st.columns(2)
usar_fechas_salidas = st.checkbox(
    "Calcular el periodo automáticamente a partir de las fechas de Salidas (recomendado)", value=True
)
if not usar_fechas_salidas:
    fecha_inicio_manual = p1.date_input("Fecha inicio del periodo")
    fecha_fin_manual = p2.date_input("Fecha fin del periodo")

calcular = st.button("🚀 3️⃣ Calcular Índice de Rotación", type="primary", use_container_width=False)

# ---------------------------------------------------------------------------
# 3) Procesamiento
# ---------------------------------------------------------------------------
if calcular:
    if inv_inicial_file is None or inv_final_file is None or salidas_file is None:
        st.error("Debes subir al menos Inventario Inicial, Inventario Final y Salidas.")
    else:
        try:
            with st.spinner("Leyendo archivos (los códigos se conservan como texto, con ceros a la izquierda si los tienen)..."):
                inv_inicial = load_inventario_file(inv_inicial_file, "Inicial")
                inv_final = load_inventario_file(inv_final_file, "Final")
                salidas = load_salidas_file(salidas_file)
                costos_manual = load_costos_file(costos_file) if costos_file is not None else None

            with st.spinner("Completando descripciones, familias y costos faltantes..."):
                maestro = build_master_lookup(inv_inicial, inv_final, salidas)
                costos = build_costos_unitarios(inv_final, inv_inicial, costos_manual)

            if usar_fechas_salidas:
                fechas_validas = salidas["fecha"].dropna()
                if fechas_validas.empty:
                    st.error("No se encontraron fechas válidas en el archivo de Salidas; ingresa el periodo manualmente.")
                    st.stop()
                fecha_inicio = fechas_validas.min().date()
                fecha_fin = fechas_validas.max().date()
            else:
                fecha_inicio, fecha_fin = fecha_inicio_manual, fecha_fin_manual

            dias_periodo = max((fecha_fin - fecha_inicio).days, 1)

            with st.spinner("Calculando índice de rotación por almacén y por material..."):
                rotacion_df = compute_rotacion(
                    inv_inicial, inv_final, salidas, costos, maestro,
                    dias_periodo=dias_periodo, umbral_rotacion_alta=umbral_rotacion_alta,
                )
                sin_rotacion_df = compute_materiales_sin_rotacion(rotacion_df)
                abc_df = compute_abc(rotacion_df, umbral_a=umbral_a, umbral_b=umbral_b)
                resumen = build_resumen(rotacion_df, salidas, fecha_inicio, fecha_fin, dias_periodo)

            st.session_state.resultado = {
                "rotacion_df": rotacion_df,
                "sin_rotacion_df": sin_rotacion_df,
                "abc_df": abc_df,
                "resumen": resumen,
            }
            st.success("Cálculo completado correctamente.")
        except ValueError as e:
            st.error(f"Error en el formato de los archivos: {e}")
        except Exception as e:
            st.error(f"Ocurrió un error procesando los archivos: {e}")

# ---------------------------------------------------------------------------
# 4) Resultados
# ---------------------------------------------------------------------------
resultado = st.session_state.resultado
if resultado is not None:
    rotacion_df = resultado["rotacion_df"]
    sin_rotacion_df = resultado["sin_rotacion_df"]
    abc_df = resultado["abc_df"]
    resumen = resultado["resumen"]

    st.markdown("---")
    st.subheader("📊 Resumen General")
    st.caption(f"Periodo analizado: {resumen['fecha_inicio']} a {resumen['fecha_fin']} ({resumen['dias_periodo']} días)")

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Inventario promedio total (S/.)", f"{resumen['total_inv_promedio']:,.0f}")
    k2.metric("Valor de salidas total (S/.)", f"{resumen['total_valor_salidas']:,.0f}")
    k3.metric("Índice de rotación (periodo)", f"{resumen['indice_total']:.2f}" if pd.notna(resumen["indice_total"]) else "N/D")
    k4.metric("Rotación anualizada", f"{resumen['rotacion_anualizada_total']:.2f}" if pd.notna(resumen["rotacion_anualizada_total"]) else "N/D")

    st.markdown("**Por almacén**")
    por_almacen_display = resumen["por_almacen"].rename(columns={
        "almacen": "Cód. Almacén", "almacen_nombre": "Almacén", "inv_inicial": "Inv. Inicial (S/.)", "inv_final": "Inv. Final (S/.)",
        "inv_promedio": "Inv. Promedio (S/.)", "valor_salidas": "Valor Salidas (S/.)",
        "num_movimientos": "N° Movimientos", "movimientos_sin_costo": "Mov. sin Costo",
        "indice_rotacion": "Índice Rotación (periodo)", "rotacion_anualizada": "Rotación Anualizada",
        "dias_inventario": "Días Inventario",
    })
    st.dataframe(por_almacen_display.round(2), hide_index=True, use_container_width=True)

    movs_sin_costo = int(resumen["por_almacen"]["movimientos_sin_costo"].sum())
    if movs_sin_costo > 0:
        st.warning(
            f"⚠️ {movs_sin_costo} movimiento(s) de salida no tienen costo unitario disponible en ningún "
            "archivo (ni en Salidas, ni en Costos Unitarios, ni en los Inventarios). Se excluyeron del "
            "valor de salidas, por lo que el índice de rotación podría estar levemente subestimado."
        )

    # --- Gráficos generales ---
    g1, g2 = st.columns(2)
    with g1:
        fig = px.bar(
            resumen["por_almacen"], x="almacen_nombre", y=["inv_promedio", "valor_salidas"], barmode="group",
            title="Inventario Promedio vs. Valor de Salidas por Almacén",
            labels={"almacen_nombre": "Almacén", "value": "S/.", "variable": "Indicador"},
        )
        st.plotly_chart(fig, use_container_width=True)
    with g2:
        dist = rotacion_df["clasificacion"].value_counts().reset_index()
        dist.columns = ["clasificacion", "cantidad"]
        fig = px.pie(dist, names="clasificacion", values="cantidad", title="Distribución de Materiales por Clasificación")
        st.plotly_chart(fig, use_container_width=True)

    # --- Tabla principal con filtros ---
    st.markdown("---")
    st.subheader("📑 Rotación por Material")
    f1, f2, f3 = st.columns(3)
    filtro_almacen = f1.multiselect("Almacén", sorted(rotacion_df["almacen_nombre"].dropna().unique()))
    filtro_clasif = f2.multiselect("Clasificación", sorted(rotacion_df["clasificacion"].dropna().unique()))
    busq_codigo = f3.text_input("Buscar por Código o Descripción")

    tabla = rotacion_df.copy()
    if filtro_almacen:
        tabla = tabla[tabla["almacen_nombre"].isin(filtro_almacen)]
    if filtro_clasif:
        tabla = tabla[tabla["clasificacion"].isin(filtro_clasif)]
    if busq_codigo:
        mask = tabla["codigo"].str.contains(busq_codigo, case=False, na=False) | \
               tabla["descripcion"].str.contains(busq_codigo, case=False, na=False)
        tabla = tabla[mask]

    tabla_display = tabla[[
        "almacen", "almacen_nombre", "codigo", "descripcion", "familia", "unidad", "cantidad_inicial", "cantidad_final",
        "cantidad_salidas", "valor_promedio", "valor_salidas", "indice_rotacion", "rotacion_anualizada",
        "dias_inventario", "clasificacion",
    ]].rename(columns={
        "almacen": "Cód. Almacén", "almacen_nombre": "Almacén", "codigo": "Código", "descripcion": "Descripción", "familia": "Familia",
        "unidad": "U.M.", "cantidad_inicial": "Cant. Inicial", "cantidad_final": "Cant. Final",
        "cantidad_salidas": "Cant. Salidas", "valor_promedio": "Valor Promedio (S/.)",
        "valor_salidas": "Valor Salidas (S/.)", "indice_rotacion": "Índice Rotación",
        "rotacion_anualizada": "Rotación Anualizada", "dias_inventario": "Días Inventario",
        "clasificacion": "Clasificación",
    })
    st.dataframe(tabla_display.round(2), hide_index=True, use_container_width=True, height=420)

    # --- Materiales sin rotación ---
    st.markdown("---")
    st.subheader("🐌 Materiales Sin Rotación (stock inmovilizado)")
    st.caption(f"{len(sin_rotacion_df)} material(es) con stock pero sin ninguna salida en el periodo.")
    sin_rot_display = sin_rotacion_df.rename(columns={
        "almacen": "Cód. Almacén", "almacen_nombre": "Almacén", "codigo": "Código", "descripcion": "Descripción", "familia": "Familia",
        "unidad": "U.M.", "cantidad_inicial": "Cant. Inicial", "cantidad_final": "Cant. Final",
        "valor_promedio": "Valor Promedio Inmovilizado (S/.)", "pct_del_total_inmovilizado": "% del Total Inmovilizado",
    })
    sin_rot_display["% del Total Inmovilizado"] = (sin_rot_display["% del Total Inmovilizado"] * 100).round(2)
    st.dataframe(sin_rot_display.round(2), hide_index=True, use_container_width=True, height=300)

    # --- ABC de salidas ---
    st.markdown("---")
    st.subheader("🏷️ Clasificación ABC de Salidas (Pareto)")
    abc_dist = abc_df["clase_abc"].value_counts().reset_index()
    abc_dist.columns = ["clase_abc", "cantidad"]
    ga, gb = st.columns(2)
    with ga:
        fig = px.bar(abc_dist.sort_values("clase_abc"), x="clase_abc", y="cantidad", title="N° de Materiales por Clase ABC")
        st.plotly_chart(fig, use_container_width=True)
    with gb:
        top15 = abc_df.nsmallest(15, "ranking")
        fig = px.bar(top15.sort_values("valor_salidas"), x="valor_salidas", y="descripcion", orientation="h",
                     title="Top 15 Materiales por Valor de Salidas", labels={"valor_salidas": "Valor Salidas (S/.)", "descripcion": "Material"})
        st.plotly_chart(fig, use_container_width=True)

    abc_display = abc_df.rename(columns={
        "ranking": "Ranking", "almacen": "Cód. Almacén", "almacen_nombre": "Almacén", "codigo": "Código", "descripcion": "Descripción",
        "valor_salidas": "Valor Salidas (S/.)", "pct_individual": "% Individual",
        "pct_acumulado": "% Acumulado", "clase_abc": "Clase ABC",
    })
    abc_display["% Individual"] = (abc_display["% Individual"] * 100).round(2)
    abc_display["% Acumulado"] = (abc_display["% Acumulado"] * 100).round(2)
    st.dataframe(abc_display, hide_index=True, use_container_width=True, height=350)

    # --- Exportación ---
    st.markdown("---")
    st.subheader("⬇️ 4️⃣ Descargar Resultado en Excel")
    if st.button("Generar archivo Excel"):
        with st.spinner("Generando archivo Excel con todas las hojas..."):
            output_path = "/tmp/resultado_indice_rotacion.xlsx"
            build_export_excel(rotacion_df, sin_rotacion_df, abc_df, resumen, output_path)
            with open(output_path, "rb") as f:
                excel_bytes = f.read()
        st.download_button(
            label="📥 Descargar Excel de Resultado",
            data=excel_bytes,
            file_name=f"indice_rotacion_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Sube los 3 archivos obligatorios (Inventario Inicial, Inventario Final y Salidas) y presiona **Calcular**.")
