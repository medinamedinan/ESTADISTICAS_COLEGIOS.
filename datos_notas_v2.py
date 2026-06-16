import pandas as pd
from playwright.sync_api import sync_playwright
from io import StringIO
import re

print("1. Cargando el archivo CSV original...")
df = pd.read_csv("centros_educativos.csv", sep=';', encoding='utf-8')

# ==========================================
# FILTROS INTELIGENTES (LÓGICA DEL USUARIO)
# ==========================================
print("2. Aplicando filtros: IES (Públicos) de Getafe y Leganés...")

# --- Filtro 1: CIUDAD ---
if 'MUNICIPIO' in df.columns:
    columna_ciudad = 'MUNICIPIO'
elif 'LOCALIDAD' in df.columns:
    columna_ciudad = 'LOCALIDAD'
else:
    print(f"❌ ERROR: No encuentro la columna de la ciudad. Tus columnas son: {df.columns.tolist()}")
    exit()

filtro_ciudad = df[columna_ciudad].astype(str).str.contains('GETAFE|LEGAN', case=False, na=False)

# --- Filtro 2: ES UN IES (Instituto de Educación/Enseñanza Secundaria) ---
# Escaneamos toda la fila buscando estas palabras exactas (lo que garantiza que es de la red pública)
expresion_ies = r'INSTITUTO DE EDUCACI[OÓ]N SECUNDARIA|INSTITUTO DE ENSE[ÑN]ANZA SECUNDARIA|\bIES\b'
filtro_ies = df.apply(lambda fila: fila.astype(str).str.contains(expresion_ies, case=False, regex=True).any(), axis=1)

# Aplicamos los filtros
df_filtrado = df[filtro_ciudad & filtro_ies].copy()

if df_filtrado.empty:
    print("❌ No encontré ningún colegio que cumpla las condiciones.")
    exit()

total_colegios = len(df_filtrado)
print(f"✅ ¡Filtro completado! Se han encontrado {total_colegios} Institutos de Educación Secundaria.\n")

# ==========================================
# PREPARACIÓN DE COLUMNAS PARA EL EXCEL
# ==========================================
df_filtrado['TITULACION_ESO'] = "Sin datos"
df_filtrado['NOTA_FASE_GENERAL_EVAU'] = "Sin datos"
df_filtrado['ALUMNOS_PRESENTADOS_EVAU'] = "Sin datos"
df_filtrado['PORCENTAJE_APTOS_EVAU'] = "Sin datos"
df_filtrado['NOTA_ALUMNOS_APTOS_EVAU'] = "Sin datos"

# ==========================================
# MOTOR DE EXTRACCIÓN MASIVA
# ==========================================
with sync_playwright() as p:
    print("3. Encendiendo el robot de extracción...")
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()

    contador = 1
    for index, row in df_filtrado.iterrows():
        codigo = row['CODIGO']
        nombre = row['CENTRO']
        url = f"https://gestiona.comunidad.madrid/wpad_pub/run/j/MostrarFichaCentro.icm?cdCentro={codigo}"

        print("-" * 70)
        print(f"📍 [{contador}/{total_colegios}] Analizando: {nombre} (Código: {codigo})")
        contador += 1

        try:
            page.goto(url, timeout=30000)
            page.wait_for_timeout(2000)
        except Exception as e:
            print(f"   ⚠️ Error de red al cargar la página. Saltando a este colegio...")
            continue

        # Entrar a RESULTADOS ACADÉMICOS
        try:
            page.evaluate("""() => {
                let tabs = Array.from(document.querySelectorAll('a, span, div'));
                let resTab = tabs.find(t => t.innerText && t.innerText.includes('RESULTADOS ACADÉMICOS'));
                if(resTab) resTab.click();
            }""")
            page.wait_for_timeout(3000)
        except:
            print("   ⚠️ Pestaña de Resultados no encontrada. Saltando...")
            continue

        # --- EXTRACCIÓN 1: % TITULACIÓN ESO ---
        try:
            page.evaluate("""() => {
                let labels = Array.from(document.querySelectorAll('label'));
                let esoBtn = labels.find(l => l.innerText.trim() === 'ESO');
                if(esoBtn) esoBtn.click();
            }""")
            page.wait_for_timeout(2000)
        except:
            pass

        dato_eso = "Sin datos"
        try:
            tablas_eso = pd.read_html(StringIO(page.content()), decimal=',', thousands='.')
            for tabla in tablas_eso:
                texto_tabla = tabla.to_string().lower()
                if "function" not in texto_tabla and ("secundaria" in texto_tabla or "titulación" in texto_tabla):
                    fila_centro = tabla[tabla.iloc[:, 0].astype(str).str.strip().str.lower() == 'centro']
                    if not fila_centro.empty:
                        for valor_bruto in reversed(fila_centro.iloc[0].astype(str).str.strip().tolist()):
                            if valor_bruto.lower() not in ['nan', 'none', '', 'centro']:
                                try:
                                    f = float(valor_bruto.replace(',', '.').replace('%', '').strip())
                                    while f > 100: f = f / 10
                                    dato_eso = f"{f:.2f}%"
                                    break
                                except ValueError:
                                    continue
        except:
            pass

        df_filtrado.at[index, 'TITULACION_ESO'] = dato_eso

        # --- EXTRACCIÓN 2: LAS 4 OPCIONES DE LA EvAU ---
        datos_evau = {
            "Nota_Fase_General": "Sin datos",
            "Alumnos_Presentados": "Sin datos",
            "Porcentaje_Aptos": "Sin datos",
            "Nota_Alumnos_Aptos": "Sin datos"
        }

        botones_evau = [
            {"label": "fase general", "clave": "Nota_Fase_General", "tipo": "nota"},
            {"label": "presentados", "clave": "Alumnos_Presentados", "tipo": "entero"},
            {"label": "% de alumnos aptos", "clave": "Porcentaje_Aptos", "tipo": "porcentaje"},
            {"label": "nota media de alumnos aptos", "clave": "Nota_Alumnos_Aptos", "tipo": "nota"}
        ]

        try:
            page.evaluate("""() => {
                let labels = Array.from(document.querySelectorAll('label'));
                let pauBtn = labels.find(l => l.innerText.trim() === 'PAU' || l.innerText.trim() === 'EVAU');
                if(pauBtn) pauBtn.click();
            }""")
            page.wait_for_timeout(1500)
        except:
            pass

        for boton in botones_evau:
            texto_buscado = boton["label"]
            clave = boton["clave"]
            tipo = boton["tipo"]

            try:
                page.evaluate(f"""(texto) => {{
                    let labels = Array.from(document.querySelectorAll('label'));
                    let targetBtn = labels.find(l => l.innerText.toLowerCase().includes(texto));
                    if(targetBtn) targetBtn.click();
                }}""", texto_buscado)

                page.wait_for_timeout(1500)

                tablas_pau = pd.read_html(StringIO(page.content()), decimal=',', thousands='.')
                for tabla in tablas_pau:
                    texto_tabla = tabla.to_string().lower()
                    if "function" not in texto_tabla and ("pau" in texto_tabla or "evau" in texto_tabla):
                        fila_centro = tabla[tabla.iloc[:, 0].astype(str).str.strip().str.lower() == 'centro']

                        if not fila_centro.empty:
                            valores_fila = fila_centro.iloc[0].astype(str).str.strip().tolist()

                            for valor_bruto in reversed(valores_fila):
                                if valor_bruto.lower() not in ['nan', 'none', '', 'centro']:
                                    try:
                                        if tipo == "nota":
                                            f = float(valor_bruto.replace(',', '.').strip())
                                            while f > 14: f = f / 10
                                            datos_evau[clave] = f"{f:.2f}"
                                        elif tipo == "porcentaje":
                                            f = float(valor_bruto.replace(',', '.').replace('%', '').strip())
                                            while f > 100: f = f / 10
                                            datos_evau[clave] = f"{f:.2f}%"
                                        elif tipo == "entero":
                                            v = valor_bruto.replace(',', '').replace('.', '').strip()
                                            datos_evau[clave] = str(int(float(v)))
                                        break
                                    except ValueError:
                                        continue
                            if datos_evau[clave] != "Sin datos":
                                break
            except:
                pass

        df_filtrado.at[index, 'NOTA_FASE_GENERAL_EVAU'] = datos_evau['Nota_Fase_General']
        df_filtrado.at[index, 'ALUMNOS_PRESENTADOS_EVAU'] = datos_evau['Alumnos_Presentados']
        df_filtrado.at[index, 'PORCENTAJE_APTOS_EVAU'] = datos_evau['Porcentaje_Aptos']
        df_filtrado.at[index, 'NOTA_ALUMNOS_APTOS_EVAU'] = datos_evau['Nota_Alumnos_Aptos']

        print(
            f"   📊 Resultados -> ESO: {dato_eso} | Presentados EvAU: {datos_evau['Alumnos_Presentados']} | Fase General: {datos_evau['Nota_Fase_General']}")

    browser.close()

# ==========================================
# GUARDADO DEL ARCHIVO FINAL
# ==========================================
print("\n" + "=" * 70)
print("💾 PROCESO COMPLETADO. GENERANDO ARCHIVO FINAL...")

nombre_archivo_salida = "resultados_ies_getafe_leganes.csv"
df_filtrado.to_csv(nombre_archivo_salida, sep=';', encoding='utf-8', index=False)

print(f"🎉 ¡Éxito total! Se ha generado tu nueva base de datos en: {nombre_archivo_salida}")
print("=" * 70 + "\n")