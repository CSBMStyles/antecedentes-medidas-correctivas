# Medidas Correctivas - Policía Nacional (RNMC)

Automatiza la consulta de medidas correctivas en el **Registro Nacional de
Medidas Correctivas** de la Policía Nacional de Colombia:

`https://srvcnpc.policia.gov.co/PSC/frm_cnp_consulta.aspx`

El sitio **no tiene CAPTCHA**: la aplicación consulta y guarda el certificado en
PDF de forma completamente automática.

---

## Qué hace

- Lee los empleados desde `CARACTERIZACION.xlsx` (hoja `CONVENIO EJECUCION`).
- Respeta **exactamente el orden de las filas del Excel** (ordenado por apellido).
- Permite elegir hoja, fila de encabezados y qué columna tiene cada dato:
  documento, fecha de expedición, nombres, apellidos, perfil y tipo de documento.
- Convierte la fecha de expedición al formato `DD/MM/AAAA` que pide el sitio.
- Consulta cada cédula y guarda el certificado en PDF:
  `MEDIDAS CORRECTIVAS/{PERFIL}/APELLIDO NOMBRE - CEDULA.pdf`
- Genera `REPORTE_FALLOS_<fecha>.xlsx`, `.csv` y `.txt` con un resumen claro.
- Guarda el avance en `estado_medidas_correctivas.json` para reanudar.
- Permite trabajar por tandas y por perfil.

---

## Fecha de expedición

El sitio pide **tipo de documento + cédula + fecha de expedición**. La fecha se
toma de la columna `Expedicion` del Excel y se envía como `DD/MM/AAAA`.

Si la fecha no coincide con el documento, el sitio rechaza la consulta y esa
persona queda en el reporte con el motivo `datos_no_coinciden`.

> En la base actual hay al menos una fecha futura (`23/05/2027`). El programa la
> detecta en la validación y la marca como probable error del Excel; el sitio
> probablemente la rechace. Corrige ese dato si es necesario.

---

## Instalación

### macOS

```bash
cd Codigo/ANTECEDENTES/obtener-medidas-correctivas
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Abre la interfaz con doble clic en **`iniciar.command`** o:

```bash
python app.py
```

### Windows

1. Instala **Python 3.11 o superior** y marca **“Add python.exe to PATH”**.
2. Instala **Google Chrome**.
3. En PowerShell:

```powershell
cd Codigo\ANTECEDENTES\obtener-medidas-correctivas
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Si PowerShell bloquea la activación:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

4. Abre la interfaz con doble clic en **`iniciar.bat`** o:

```powershell
python app.py
```

---

## La interfaz

### 1 · Datos

- Excel de caracterización y hoja.
- Fila de encabezados.
- Mapeo de columnas: documento, fecha de expedición, nombres, apellidos, perfil
  y tipo de documento (opcional).
- Vista previa con las columnas mapeadas primero y los valores alineados.
- **Validar datos**: cuenta válidos, incompletos y duplicados, y avisa de fechas
  futuras.

### 2 · Salida y navegador

- Carpeta base (`MEDIDAS CORRECTIVAS` por defecto).
- Nombre del PDF:
  - `APELLIDO NOMBRE - CEDULA` (recomendado).
  - `APELLIDO NOMBRE - CEDULA`.
  - `CEDULA NOMBRE APELLIDO`.
- Chrome en modo **headless** por defecto (no hay captcha). Si algo falla,
  desmarca la casilla para ver el navegador.
- Driver: `auto`, `uc` o `selenium`.

### 3 · Automático

- Explica que el trámite no tiene CAPTCHA.
- **Probar formulario**: abre el sitio y verifica que estén el tipo de documento,
  la cédula, la fecha y el botón Consultar. No consulta ninguna cédula.

### 4 · Ejecución

- Perfil (`TODOS`, `AUXILIAR`, `ENFERMERO`, `MEDICO`).
- Cantidad por tanda (`0` = todos).
- Reanudar o reiniciar el estado.
- Intentos por persona si la consulta falla.
- Barra de progreso y botón **Detener**.

---

## Salida

```
MEDIDAS CORRECTIVAS/
├── AUXILIAR/
│   ├── PEREZ GOMEZ ANA MARIA - 1000000001.pdf
│   └── ...
├── ENFERMERO/
└── MEDICO/
```

El nombre empieza por el apellido para conservar el orden de la base.

---

## Uso por línea de comandos

```bash
python main.py
python main.py --perfil AUXILIAR --tanda 10 --si
python main.py --perfil TODOS --si
python main.py --reiniciar-estado --si
python main.py --excel "/ruta/CARACTERIZACION.xlsx" --salida "/ruta/salida"
python main.py --hoja "CONVENIO EJECUCION" --fila-encabezado 1 --si
python main.py --dry-run --si
python main.py --visible --si
```

`--dry-run` valida el Excel sin abrir Chrome. `--visible` abre Chrome para ver el
proceso.

---

## Problemas frecuentes

### “This version of ChromeDriver only supports Chrome version …”

`undetected-chromedriver` tiene un driver viejo. El programa cae automáticamente a
**Selenium estándar**, que descarga el driver correcto. Si quieres forzarlo:

```bash
DRIVER_BACKEND=selenium
```

o elige `selenium` en la pestaña 2.

### La consulta falla con “la fecha de expedición no es correcta”

La fecha de expedición del Excel no coincide con la del documento o está en el
futuro. Corrige la columna `Expedicion` y vuelve a intentar. Esa persona queda en
el reporte con `fecha_invalida` y la explicación de qué corregir.

### “La fecha o el documento no coinciden”

El sitio no encontró coincidencia entre la cédula y la fecha. Revisa ambos datos
en el Excel. El reporte lo marca como `datos_no_coinciden`.

### El sitio no muestra el certificado

- Aumenta los intentos en la pestaña 4.
- Prueba con Chrome visible (desmarca headless).
- Revisa la captura en `errors/` y el reporte.

### El PDF sale con dos páginas

El certificado puede ocupar más de una página si la persona tiene medidas. A
diferencia de los antecedentes judiciales, aquí **no se recorta**: se guarda todo
el contenido para no perder información.

### Windows

- Si el antivirus bloquea `undetected_chromedriver.exe`, usa `selenium`.
- Si el proyecto está en OneDrive, mueve la carpeta de salida a una ruta local.
- Los nombres se recortan a 120 caracteres para no superar el límite de Windows.

---

## Archivos

- `app.py`: lanzador de la interfaz.
- `gui.py`: ventana.
- `main.py`: núcleo de consulta y CLI.
- `config.py`: `config.json` y `.env`.
- `sources.py`: lectura del Excel, hojas, mapeo y orden.
- `utils.py`: fechas `DD/MM/AAAA`, tipos de documento y nombres seguros.
- `perfiles.py`: campos del proyecto (base para los otros antecedentes).
- `estado_medidas_correctivas.json`: control de avance.
- `errors/`: capturas cuando falla una consulta.
- `REPORTE_FALLOS_*.xlsx`: reporte en tabla, con una fila por persona.
- `REPORTE_FALLOS_*.csv` / `.txt`: los mismos datos en texto.

### Reporte de fallos

El archivo `.xlsx` tiene dos hojas:

- **Fallos**: una fila por persona con estas columnas:
  `Fila Excel`, `Documento`, `Nombre`, `Perfil`, `Fecha expedición`, `Motivo`,
  `Qué pasó`, `Qué hacer`, `Mensaje del sitio`, `Intentos`, `Captura`.
- **Resumen**: cuántos fallos hay por motivo y qué significa cada uno.

El `.txt` es la versión corta para leer rápido; el `.csv` sirve para filtrar en
Excel. La captura de cada fallo queda enlazada desde la columna `Captura`.

---

## Reutilizar en los otros antecedentes

`perfiles.py` declara los campos y el tipo de captcha. Para procuraduría, rethus,
inhabilidades, etc., se copia la carpeta y se cambia:

- `PROYECTO` (nombre y entidad).
- `CAMPOS` (qué columnas necesita ese trámite).
- El flujo del sitio en `main.py`.

La ventana, `config.py`, `sources.py` y el manejo de estado/reportes se reutilizan.

---

## Uso autorizado

Este proyecto está pensado para que una entidad autorizada consulte los
antecedentes de sus colaboradores. Usa la información solo para los fines
institucionales previstos.

---

## Datos y privacidad

Este repositorio contiene únicamente código y documentación. **No incluye**:

- API keys ni archivos `.env`.
- Bases de datos de empleados, cédulas, nombres, correos o teléfonos.
- PDFs, Excel, PPTX, CSV, capturas o reportes generados por las consultas.
- Configuraciones locales (`config.json`) ni estados de proceso.

Para usarlo, proporciona tu propia base de datos autorizada y configura tus
propias credenciales. No subas datos personales al repositorio.

El archivo `config.example.json` muestra una configuración mínima; cópialo a
`config.json` o deja que la interfaz lo cree.

## Licencia

MIT. Consulta el archivo `LICENSE`.

