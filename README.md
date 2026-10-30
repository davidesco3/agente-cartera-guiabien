# 🛡️ GuíaBien Asesores en Seguros — Depurador Operativo de Cartera

Aplicación web interna de alta precisión construida en **Streamlit** y **Pandas** para procesar, auditar y depurar semanalmente los informes operativos de cartera e integraciones con **Seguros SURA**, diseñada con la identidad visual corporativa de **[guiabien.com](https://guiabien.com)**.

---

## 🚀 Inicio Rápido

### 1. Requisitos Previos
Asegúrate de contar con Python 3.10+ y las dependencias instaladas:
```bash
pip install -r requirements.txt
```
*(O directamente: `pip install pandas streamlit xlsxwriter openpyxl requests`)*

### 2. Ejecutar la Aplicación
```bash
python -m streamlit run app.py
```
Abre en tu navegador: `http://localhost:8501`

---

## 🧩 Arquitectura del Proyecto

```
Agente de cartera/
├── app.py                      # Interfaz gráfica Streamlit con look & feel GuíaBien
├── cartera_processor.py        # Motor de procesamiento de datos, normalización y lógica de negocio
├── excel_exporter.py           # Generación de Excel con XlsxWriter y formato condicional
├── config.py                   # Constantes de negocio, paleta de colores y rutas
├── requirements.txt            # Dependencias del proyecto
├── data/
│   ├── clientes_elite.csv      # Base persistente local de Clientes Élite
│   ├── no_contactar_cache.csv  # Respaldo local de lista No Contactar (en caso de caída/privacidad de Sheets)
│   └── ejemplos/               # Archivos de prueba operativos (.xlsx)
├── assets/
│   └── logo_guiabien.png       # Logotipo institucional descargado de guiabien.com
└── tests/
    └── test_cartera.py         # Suite completa de pruebas automatizadas
```

---

## ⚙️ Reglas de Negocio Implementadas

1. **Normalización Rigurosa**:
   - Conversión obligatoria de `Nº ident.fis.1` y `Póliza` a texto limpio (eliminando `.0` de floats, espacios y caracteres especiales) para garantizar cruces 100% confiables.
2. **Paso 1: Separar Devoluciones**:
   - `Operación == "Devolución"` se aísla en la hoja `Devoluciones`.
   - Únicamente los registros con `Operación == "Cobro"` continúan al flujo principal (`df_base`).
3. **Paso 2: Exclusión de Débitos Prematuros**:
   - `Medio Pago == "Débito"` y `Dias Cartera <= 16` son retirados del análisis (cobros en curso regular de ciclo bancario).
4. **Paso 3: Débitos Fallidos**:
   - `Medio Pago == "Débito"` y `Dias Cartera > 16` se trasladan a `Débitos fallidos`.
   - Se redacta automáticamente el mensaje de WhatsApp/SMS en la columna `Mensaje_Contacto`:
     > *"Hola {Nombre Tomador}, te escribe Paula de GuíaBien... te estamos contactando porque nos notifica SURA que el débito de tu póliza de {Solución} número {Póliza} fue rechazado por el banco por {Resp.Cobro}. Por favor revisa con tu banco para evitar que perdamos cobertura."*
     *(Si `Resp.Cobro` está vacío o nulo, se reemplaza por "motivos no especificados")*.
5. **Paso 4: Clientes de Alto Impacto (Urgentes)**:
   - Se extraen a `Alto impacto` los registros que cumplan cualquiera de estas tres condiciones:
     - Documento en `data/clientes_elite.csv`.
     - `Saldo a pagar > 3,000,000`.
     - `Dias Cartera >= 45`.
   - Se añade la columna `Motivo_Alto_Impacto` para que el equipo conozca la causal exacta.
6. **Paso 5: Pólizas a Revisar por Posible Cancelación**:
   - Pólizas presentes en la *Cartera Semana Anterior* que **NO** están en la *Cartera Actual* y **TAMPOCO** en el *Informe de Recaudos*.
7. **Paso 6: Cruce con Lista No Contactar**:
   - Consulta el Google Sheets institucional.
   - Aplica columna `Estado_Contacto = "No contactar"` a las 5 hojas.
8. **Exportación y Formato Condicional en Excel (XlsxWriter)**:
   - 5 hojas exactas:
     1. `Base depurada`
     2. `Devoluciones`
     3. `Débitos fallidos`
     4. `Alto impacto`
     5. `Revisar cancelación`
   - Cualquier fila marcada como `"No contactar"` se resalta por completo con relleno rojo suave (`#FFC7CE`) y texto rojo oscuro (`#9C0006`).

---

## 🔍 Observación Crítica y Recomendación Técnica (Senior Review)

> ⚠️ **Google Sheets 401 Unauthorized**:
> El enlace suministrado de Google Sheets (`gid=195218118`) requiere credenciales de acceso o está configurado como privado, devolviendo un error `HTTP 401/302 Redirect`. 
> 
> **Cómo lo resolvió esta arquitectura:**
> 1. El motor realiza una consulta tolerante a fallos. Si detecta el bloqueo 401 o falta de internet, **no quiebra el flujo**, sino que recurre automáticamente a la caché local `data/no_contactar_cache.csv`.
> 2. En la barra lateral se añadió un cargador de contingencia opcional que permite al usuario subir manualmente la lista si el enlace sigue restringido.
> 3. **Para activar la sincronización automática 100% en vivo:** Solicita al administrador del Google Sheet cambiar el acceso general a *"Cualquier persona con el enlace puede ver"*.

---

## 🧪 Pruebas Automatizadas
Para ejecutar la batería de pruebas:
```bash
python tests/test_cartera.py
```
Resultado: **100% PASS** en todas las suites.
