import React, { useState, useRef, useEffect } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import { createWorker } from "tesseract.js";
import { Capacitor } from "@capacitor/core";
import { Filesystem, Directory } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";
import { Preferences } from "@capacitor/preferences";
import { BrowserMultiFormatReader } from "@zxing/library";
import JSZip from "jszip";

const emptyForm = { marca: "", modelo: "", serie: "", precio: "", ubicacion: "", notas: "", fotos: [] };

// Tipos de foto que se pueden asignar a cada placa
const TIPOS_FOTO = [
  { id: "PlacaBaseDelantera", label: "Base · Delantera", color: "#1f6d3f" },
  { id: "PlacaBaseTrasera", label: "Base · Trasera", color: "#4a9a68" },
  { id: "FuenteDelantera", label: "Fuente · Delantera", color: "#c07a12" },
  { id: "FuenteTrasera", label: "Fuente · Trasera", color: "#e0a23e" },
  { id: "TConDelantera", label: "T-Con · Delantera", color: "#1e5aa8" },
  { id: "TConTrasera", label: "T-Con · Trasera", color: "#4d8fd6" },
  { id: "Variadas", label: "Variadas / Otro", color: "#6b7280" },
];

const STORAGE_KEY = "catalogo_placas_items_v2";

function resizeImage(file, maxWidth = 900) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("read failed"));
    reader.onload = (e) => {
      const img = new Image();
      img.onerror = () => reject(new Error("image failed"));
      img.onload = () => {
        const scale = Math.min(1, maxWidth / img.width);
        const canvas = document.createElement("canvas");
        canvas.width = img.width * scale || img.width;
        canvas.height = img.height * scale || img.height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", 0.78));
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

function Icon({ children, size = 16, color = "#1c1f26" }) {
  return <span style={{ fontSize: size, color, lineHeight: 1, display: "inline-flex" }}>{children}</span>;
}

const KNOWN_BRANDS = [
  "Balay", "Bosch", "Siemens", "Samsung", "LG", "Philips", "Sony", "Panasonic",
  "Whirlpool", "Electrolux", "AEG", "Candy", "Beko", "Haier", "Sharp", "Hisense",
  "TCL", "Fagor", "Teka", "Zanussi", "Indesit", "Miele", "Grundig", "Toshiba",
];

function parseLabelText(rawText) {
  const lines = rawText.split("\n").map((l) => l.trim()).filter(Boolean);
  const result = { modelo: "", serie: "", marca: "" };
  const modeloRegex = /(MOD(ELO)?|MODEL)\s*[:.\-]?\s*([A-Z0-9\-\/]{3,})/i;
  const serieRegex = /(S\s*\/?\s*N|SERIE|SERIAL(\s*NO)?|NO\.?\s*SERIE)\s*[:.\-]?\s*([A-Z0-9\-]{4,})/i;

  const fullTextUpper = rawText.toUpperCase();
  for (const brand of KNOWN_BRANDS) {
    if (fullTextUpper.includes(brand.toUpperCase())) {
      result.marca = brand;
      break;
    }
  }
  for (const line of lines) {
    const mMatch = line.match(modeloRegex);
    if (mMatch && !result.modelo) result.modelo = mMatch[3];
    const sMatch = line.match(serieRegex);
    if (sMatch && !result.serie) result.serie = sMatch[3];
  }
  if (!result.modelo) {
    const codeLine = lines.find((l) => /[A-Z]/.test(l) && /\d/.test(l) && l.length >= 5 && l.length <= 20);
    if (codeLine) result.modelo = codeLine.replace(/\s+/g, "");
  }
  return result;
}

async function detectBarcode(dataUrl) {
  try {
    const reader = new BrowserMultiFormatReader();
    const img = new Image();
    await new Promise((resolve, reject) => {
      img.onload = resolve;
      img.onerror = reject;
      img.src = dataUrl;
    });
    const result = await reader.decodeFromImageElement(img);
    return result?.getText ? result.getText() : null;
  } catch (err) {
    return null;
  }
}

let ocrWorkerPromise = null;
function getOcrWorker() {
  if (!ocrWorkerPromise) {
    ocrWorkerPromise = createWorker("spa", 1, {
      workerPath: "vendor/worker.min.js",
      corePath: "vendor/tesseract-core-simd.wasm.js",
      langPath: "tessdata",
      gzip: true,
    });
  }
  return ocrWorkerPromise;
}

const TEXTO_INSTRUCCIONES = `INSTRUCCIONES PARA EL ORDENADOR
================================

1. Descomprime este ZIP en el ordenador, en cualquier carpeta.

2. Dentro veras estos archivos:
     - carpeta "fotos"
     - "catalogo.xlsx"
     - "generar_catalogo_final.py"
     - "subir_placas_final.py"
     - "EJECUTAR.bat"   <-- este es el importante

3. Haz DOBLE CLIC en "EJECUTAR.bat".
   Se abrira una ventana negra que hace todo solo:
     - instala lo necesario (solo tarda la primera vez)
     - genera el catalogo final
     - abre el navegador y empieza a subir los anuncios,
       preguntando lo que haga falta por el camino

4. Sigue lo que te vaya preguntando esa ventana.

Creado por MCG.
`;


// --- Contenido de los scripts Python, incrustado para meterlos en el ZIP ---
const SUBIR_PLACAS_PY = "\"\"\"\nSCRIPT FINAL - Sube las placas al panel WebActiva (electronicaaranjuez.com)\n============================================================================\n\nQU\u00c9 HACE, PASO A PASO:\n  1. Lee el Excel procesado (catalogo_placas_procesado.xlsx), que ya tiene\n     el t\u00edtulo de cada anuncio calculado (columna \"Titulo Anuncio\").\n  2. Por cada placa:\n       a) Genera un COLLAGE con todas sus fotos (las junta en una sola imagen,\n          en cuadr\u00edcula), porque la web solo admite 1 foto por producto.\n       b) Te ense\u00f1a un resumen de la placa (n\u00famero, t\u00edtulo, modelo...) y\n          te pregunta \"\u00bfSubir esta placa? (s/n)\" -- SOLO CONTIN\u00daA SI DICES \"s\".\n       c) Si dices que s\u00ed: abre \"Productos > Nuevo\" en el panel, rellena\n          los campos, sube el collage, y guarda.\n       d) Si dices que no: la salta y sigue con la siguiente.\n  3. Al final te da un resumen de cu\u00e1ntas subiste, cu\u00e1ntas saltaste y si\n     hubo errores.\n\nANTES DE EJECUTAR:\n  pip install playwright openpyxl pillow\n  playwright install chromium\n\nC\u00d3MO EJECUTAR:\n  python subir_placas_final.py\n\nLA PRIMERA VEZ: pru\u00e9balo con 2-3 placas (contestando \"n\" al resto) para\ncomprobar que todo se rellena bien antes de lanzarlo con las 74.\n\nCAMPOS QUE FALTAN POR CONFIRMAR (b\u00fascalos con \"# >>> AJUSTAR\" si hace falta):\n  - Selectores exactos de login (usuario/contrase\u00f1a/bot\u00f3n) -- puestos seg\u00fan\n    lo visto en tus capturas, pero revisa si falla.\n  - Palabra clave (tagsAC): de momento uso la MARCA. Si prefieres otra cosa,\n    cambia la funci\u00f3n `obtener_palabra_clave()` m\u00e1s abajo.\n\"\"\"\n\nimport re\nimport time\nimport getpass\nfrom pathlib import Path\nfrom collections import defaultdict\n\nfrom openpyxl import load_workbook\nfrom PIL import Image\nfrom playwright.sync_api import sync_playwright\n\n# =========================================================\n# CONFIGURACI\u00d3N\n# =========================================================\n\nURL_LOGIN = \"https://www.electronicaaranjuez.com/admin/login\"\n\nRUTA_EXCEL = \"catalogo_final_anuncios.xlsx\"\nCARPETA_FOTOS = \"fotos\"\nCARPETA_COLLAGES = \"collages\"  # se crea sola, aqu\u00ed se guardan las im\u00e1genes combinadas\n\nESPERA = 1.0  # segundos entre acciones; sube esto si la web va lenta\n\nPROGRESO_ARCHIVO = \"subidos_progreso.txt\"\n\n\ndef cargar_progreso():\n    p = Path(PROGRESO_ARCHIVO)\n    if not p.exists():\n        return set()\n    return set(p.read_text(encoding=\"utf-8\").splitlines())\n\n\ndef marcar_subido(identificador):\n    with open(PROGRESO_ARCHIVO, \"a\", encoding=\"utf-8\") as f:\n        f.write(identificador + \"\\n\")\n\n# =========================================================\n# 1) GENERAR EL COLLAGE DE CADA PLACA\n# =========================================================\n\ndef fotos_de_la_fila(fila, carpeta):\n    \"\"\"Lee las fotos exactas asignadas a este anuncio (columna 'Fotos Asignadas').\"\"\"\n    carpeta = Path(carpeta)\n    nombres = str(fila.get(\"Fotos Asignadas\", \"\")).split(\";\")\n    return [carpeta / n.strip() for n in nombres if n.strip()]\n\n\ndef crear_collage(fotos, ruta_salida, max_fotos=6):\n    \"\"\"Junta las fotos de una placa en una sola imagen en cuadr\u00edcula.\"\"\"\n    fotos = fotos[:max_fotos]  # por si acaso hay m\u00e1s de la cuenta\n    imgs = [Image.open(f).convert(\"RGB\") for f in fotos]\n\n    # tama\u00f1o de cada miniatura dentro del collage\n    miniatura = (500, 500)\n    imgs = [im.copy() for im in imgs]\n    for im in imgs:\n        im.thumbnail(miniatura)\n\n    n = len(imgs)\n    columnas = 2 if n > 1 else 1\n    filas = (n + columnas - 1) // columnas\n\n    ancho = columnas * miniatura[0]\n    alto = filas * miniatura[1]\n    collage = Image.new(\"RGB\", (ancho, alto), color=\"white\")\n\n    for i, im in enumerate(imgs):\n        x = (i % columnas) * miniatura[0]\n        y = (i // columnas) * miniatura[1]\n        collage.paste(im, (x, y))\n\n    collage.save(ruta_salida, quality=85)\n    return ruta_salida\n\n\n# =========================================================\n# 2) LEER EL EXCEL\n# =========================================================\n\ndef leer_excel(ruta):\n    wb = load_workbook(ruta)\n    ws = wb.active\n    cabeceras = [c.value for c in ws[1]]\n    filas = []\n    for row in ws.iter_rows(min_row=2, values_only=True):\n        d = dict(zip(cabeceras, row))\n        if d.get(\"N\u00ba Original\") is None:\n            continue\n        filas.append(d)\n    return filas\n\n\ndef obtener_palabra_clave(fila):\n    # >>> AJUSTAR aqu\u00ed si la palabra clave no debe ser la marca\n    return str(fila.get(\"Marca\", \"\")).strip()\n\n\nPRECIOS_POR_TIPO = {\n    \"T-Con\": \"20\",\n    \"Fuente\": \"26\",\n    \"Placa Base\": \"35\",\n    \"PACK / Variadas\": \"45\",\n}\n\n\ndef obtener_precio(fila):\n    tipo = str(fila.get(\"Tipo Placa\", \"\")).strip()\n    return PRECIOS_POR_TIPO.get(tipo, \"25\")  # 25 por defecto si no coincide ningun tipo\n\n\n# IDs reales de categor\u00eda sacados del desplegable de la web (ver HTML del formulario)\nCATEGORIAS_POR_TIPO = {\n    \"Fuente\": \"2361425\",       # Fuentes alimentaci\u00f3n TV\n    \"Placa Base\": \"2361424\",   # Placas base para televisi\u00f3n\n    \"T-Con\": \"2361424\",        # No hay categoria especifica de T-Con -> se manda a Placas base\n    \"PACK / Variadas\": \"2361424\",\n}\n\n\ndef obtener_categoria(fila):\n    tipo = str(fila.get(\"Tipo Placa\", \"\")).strip()\n    return CATEGORIAS_POR_TIPO.get(tipo, \"2361424\")\n\n\nTIPO_TEXTO_TITULO = {\n    \"Placa Base\": \"PLACA MAIN\",\n    \"Fuente\": \"PLACA FUENTE\",\n    \"T-Con\": \"PLACA T-CON\",\n    \"PACK / Variadas\": \"PLACAS VARIAS\",\n}\n\n\ndef construir_titulo(fila):\n    marca = str(fila.get(\"Marca\", \"\") or \"\").strip().upper()\n    modelo = str(fila.get(\"Modelo\", \"\") or \"\").strip()\n    tipo = str(fila.get(\"Tipo Placa\", \"\")).strip()\n    tipo_texto = TIPO_TEXTO_TITULO.get(tipo, tipo.upper())\n    partes = [p for p in [marca, modelo, tipo_texto] if p]\n    return \" \".join(partes)\n\n\n# =========================================================\n# 3) AUTOMATIZACI\u00d3N DEL NAVEGADOR\n# =========================================================\n\ndef login(page, usuario, contrasena):\n    page.goto(URL_LOGIN)\n    page.fill('input[name=\"_username\"]', usuario)\n    page.fill('input[name=\"_password\"]', contrasena)\n    page.click('form.form-login button[type=\"submit\"]')\n    page.wait_for_load_state(\"networkidle\")\n\n\nURL_NUEVO_PRODUCTO = \"https://www.electronicaaranjuez.com/admin/backendbundle/producto/producto/create\"\n\n\ndef ir_a_nuevo_producto(page):\n    page.goto(URL_NUEVO_PRODUCTO)\n    page.wait_for_load_state(\"networkidle\")\n\n\ndef rellenar_formulario(page, fila, ruta_collage):\n    titulo = construir_titulo(fila)\n    modelo = str(fila.get(\"Modelo\", \"\") or \"\")\n    n_serie = str(fila.get(\"N\u00ba Serie\", \"\") or \"\")\n\n    # Los campos llevan un prefijo \u00fanico que cambia cada vez (uniqid de Sonata),\n    # por eso buscamos por el FINAL del name en vez del nombre completo.\n    page.fill('input[name$=\"[nombre]\"]', titulo)\n    page.fill('input[name$=\"[mpn]\"]', modelo)\n    if n_serie:\n        page.fill('input[name$=\"[referencia]\"]', n_serie)\n    page.fill('input[name$=\"[stock]\"]', \"1\")\n    page.fill('input[name$=\"[precio]\"]', obtener_precio(fila))\n\n    # Marcar \"Destacado\" (feature) y \"Portada\" para que salga en la tienda.\n    # Estas casillas est\u00e1n ocultas visualmente (las pinta un plugin encima),\n    # por eso usamos force=True para marcarlas igualmente.\n    try:\n        page.check('input[name$=\"[feature]\"]', force=True)\n    except Exception:\n        pass\n    try:\n        page.check('input[name$=\"[portada]\"]', force=True)\n    except Exception:\n        pass\n\n    # Categor\u00eda (obligatoria) -- sin esto el producto no sale en la tienda publica\n    categoria_id = obtener_categoria(fila)\n    page.select_option('select[name$=\"[categPorDefecto_enCreate]\"]', value=categoria_id)\n    time.sleep(0.3)\n\n    # Palabra clave / tagsAC (autocompletar)\n    palabra = obtener_palabra_clave(fila)\n    if palabra:\n        campo_tag = page.locator('input[type=\"text\"][name$=\"[tagsAC]\"]')\n        campo_tag.fill(palabra)\n        time.sleep(ESPERA)\n        # Si aparece una sugerencia ya existente en la lista, la seleccionamos\n        # (para no crear una etiqueta duplicada); si no aparece, se deja el texto tal cual.\n        sugerencia = page.locator(f'ul.ui-autocomplete li:has-text(\"{palabra}\")').first\n        if sugerencia.count() > 0:\n            sugerencia.click()\n        else:\n            page.keyboard.press(\"Escape\")  # cierra el desplegable sin seleccionar nada raro\n\n    # Foto (collage)\n    page.set_input_files('input[type=\"file\"][name$=\"[imagenPrevia]\"]', str(ruta_collage))\n    time.sleep(ESPERA)\n\n\ndef guardar_y_nuevo(page):\n    # Este bot\u00f3n guarda la placa Y abre autom\u00e1ticamente un formulario \"Nuevo\" en blanco,\n    # as\u00ed que no hace falta volver a navegar a Productos > Nuevo en la siguiente vuelta.\n    page.click('button[name=\"btn_create_and_create\"]')\n    page.wait_for_load_state(\"networkidle\")\n\n\n# =========================================================\n# 4) PROGRAMA PRINCIPAL\n# =========================================================\n\ndef main():\n    print(\"=== Subida autom\u00e1tica de placas a WebActiva ===\\n\")\n    usuario = input(\"Usuario: \").strip()\n    contrasena = getpass.getpass(\"Contrase\u00f1a: \").strip()\n\n    filas = leer_excel(RUTA_EXCEL)\n    ya_subidos = cargar_progreso()\n    pendientes = [f for f in filas if f\"{f['N\u00ba Original']}_{f.get('Sub','')}\" not in ya_subidos]\n\n    print(f\"\\nTotal de anuncios en el Excel: {len(filas)}\")\n    print(f\"Ya subidos en sesiones anteriores: {len(ya_subidos)}\")\n    print(f\"Pendientes: {len(pendientes)}\\n\")\n\n    if not pendientes:\n        print(\"No queda nada pendiente por subir. (Si quieres volver a subir todo,\")\n        print(f\"borra el archivo '{PROGRESO_ARCHIVO}' y vuelve a ejecutar.)\")\n        return\n\n    modo = input(\"\u00bfConfirmar cada anuncio uno a uno (m) o subir autom\u00e1tico sin preguntar (a)?: \").strip().lower()\n    automatico = modo == \"a\"\n\n    limite_txt = input(f\"\u00bfCu\u00e1ntos quieres subir en esta sesi\u00f3n? (Enter = los {len(pendientes)} pendientes): \").strip()\n    limite = int(limite_txt) if limite_txt.isdigit() else len(pendientes)\n    tanda = pendientes[:limite]\n\n    Path(CARPETA_COLLAGES).mkdir(exist_ok=True)\n\n    with sync_playwright() as p:\n        browser = p.chromium.launch(headless=False, slow_mo=150)\n        page = browser.new_page()\n\n        login(page, usuario, contrasena)\n        print(\"Login correcto.\\n\")\n\n        subidas = 0\n        saltadas = 0\n        errores = []\n        primera_subida = True\n\n        for i, fila in enumerate(tanda, start=1):\n            numero = fila[\"N\u00ba Original\"]\n            sub = fila.get(\"Sub\", \"\")\n            identificador = f\"{numero}_{sub}\"\n            titulo = construir_titulo(fila)\n\n            print(f\"\\n[{i}/{len(tanda)}] Placa n\u00ba {numero} (parte {sub})\")\n            print(f\"   T\u00edtulo: {titulo}\")\n\n            if not automatico:\n                respuesta = input(\"   \u00bfSubir este anuncio? (s/n): \").strip().lower()\n                if respuesta != \"s\":\n                    print(\"   -> Saltada.\")\n                    saltadas += 1\n                    continue\n\n            fotos = fotos_de_la_fila(fila, CARPETA_FOTOS)\n\n            try:\n                ruta_collage = Path(CARPETA_COLLAGES) / f\"{identificador}_collage.jpg\"\n                crear_collage(fotos, ruta_collage)\n\n                if primera_subida:\n                    ir_a_nuevo_producto(page)\n                    primera_subida = False\n\n                rellenar_formulario(page, fila, ruta_collage)\n                guardar_y_nuevo(page)\n\n                print(\"   -> Subida correctamente.\")\n                marcar_subido(identificador)\n                subidas += 1\n                time.sleep(ESPERA)\n\n            except Exception as e:\n                print(f\"   -> ERROR: {e}\")\n                errores.append(numero)\n                primera_subida = True\n                continue\n\n        browser.close()\n\n        print(\"\\n--- RESUMEN DE ESTA SESI\u00d3N ---\")\n        print(f\"Subidas: {subidas}\")\n        print(f\"Saltadas: {saltadas}\")\n        print(f\"Con error: {len(errores)}\")\n        if errores:\n            print(\"Placas con error (revisar a mano):\", errores)\n        print(f\"\\nProgreso guardado en '{PROGRESO_ARCHIVO}'.\")\n        print(\"Si vuelves a ejecutar el script, seguir\u00e1 donde lo dejaste.\")\n\n\nif __name__ == \"__main__\":\n    main()\n";
const GENERAR_CATALOGO_PY = "\"\"\"\nGENERAR CATALOGO FINAL\n=======================\nEste script une lo que exporta la app del m\u00f3vil (catalogo.xlsx + carpeta\nfotos/ con el tipo ya incluido en el nombre de cada foto) y genera el\nExcel final con los anuncios ya divididos (Placa Base / Fuente / T-Con),\nlisto para que lo use \"subir_placas_final.py\".\n\nC\u00d3MO USARLO:\n  1. Descomprime el ZIP que manda la app del m\u00f3vil en esta misma carpeta.\n     Deber\u00edas tener aqu\u00ed: \"catalogo.xlsx\" y la carpeta \"fotos\".\n  2. Ejecuta:\n        py generar_catalogo_final.py\n  3. Se genera \"catalogo_final_anuncios.xlsx\" en esta misma carpeta.\n  4. Despu\u00e9s ejecuta \"subir_placas_final.py\" como siempre.\n\nCreado por MCG.\n\"\"\"\n\nimport re\nfrom pathlib import Path\nfrom collections import defaultdict\nfrom openpyxl import load_workbook, Workbook\n\nRUTA_CATALOGO = \"catalogo.xlsx\"\nCARPETA_FOTOS = \"fotos\"\nRUTA_SALIDA = \"catalogo_final_anuncios.xlsx\"\n\n# Debe coincidir exactamente con los ids usados en la app (TIPOS_FOTO)\nPARES_TIPO = [\n    (\"PlacaBaseDelantera\", \"PlacaBaseTrasera\", \"Placa Base\"),\n    (\"FuenteDelantera\", \"FuenteTrasera\", \"Fuente\"),\n    (\"TConDelantera\", \"TConTrasera\", \"T-Con\"),\n]\nTODOS_LOS_TIPOS = [t for par in PARES_TIPO for t in par[:2]] + [\"Variadas\", \"SinClasificar\"]\n\n\ndef detectar_tipo(nombre_archivo):\n    for tipo in TODOS_LOS_TIPOS:\n        if tipo in nombre_archivo:\n            return tipo\n    return \"SinClasificar\"\n\n\nTIPO_TEXTO_TITULO = {\n    \"Placa Base\": \"PLACA MAIN\",\n    \"Fuente\": \"PLACA FUENTE\",\n    \"T-Con\": \"PLACA T-CON\",\n    \"PACK / Variadas\": \"PLACAS VARIAS\",\n}\n\n\ndef main():\n    if not Path(RUTA_CATALOGO).exists():\n        print(f\"ERROR: no encuentro '{RUTA_CATALOGO}' en esta carpeta.\")\n        input(\"Pulsa Enter para salir...\")\n        return\n    if not Path(CARPETA_FOTOS).exists():\n        print(f\"ERROR: no encuentro la carpeta '{CARPETA_FOTOS}' en esta carpeta.\")\n        input(\"Pulsa Enter para salir...\")\n        return\n\n    # 1) Agrupar fotos por numero de placa y tipo\n    por_placa = defaultdict(lambda: defaultdict(list))\n    for f in sorted(Path(CARPETA_FOTOS).glob(\"*\")):\n        m = re.match(r\"(\\d+)_\", f.name)\n        if not m:\n            continue\n        num = int(m.group(1))\n        tipo = detectar_tipo(f.name)\n        por_placa[num][tipo].append(f.name)\n\n    # 2) Leer datos del catalogo (marca, modelo, etc.)\n    wb_cat = load_workbook(RUTA_CATALOGO)\n    ws_cat = wb_cat.active\n    datos = {}\n    for i, row in enumerate(ws_cat.iter_rows(min_row=2, values_only=True), start=1):\n        if row[0] is None:\n            continue\n        num = int(row[0])\n        datos[num] = dict(marca=row[1], modelo=row[2], n_serie=row[3], precio=row[4], ubic=row[5], notas=row[6])\n\n    # 3) Construir el Excel final\n    wb_new = Workbook()\n    ws_new = wb_new.active\n    ws_new.title = \"Anuncios Finales\"\n    ws_new.append([\"N\u00ba Original\", \"Sub\", \"Marca\", \"Modelo\", \"N\u00ba Serie\", \"Precio\",\n                    \"Ubicacion\", \"Notas\", \"Tipo Placa\", \"Titulo Anuncio\", \"Fotos Asignadas\", \"Aviso\"])\n\n    avisos = []\n\n    for num in sorted(por_placa.keys()):\n        d = datos.get(num, {})\n        marca = d.get(\"marca\")\n        modelo = d.get(\"modelo\")\n        n_serie = d.get(\"n_serie\")\n        precio = d.get(\"precio\")\n        ubic = d.get(\"ubic\")\n        notas = d.get(\"notas\")\n        tipos = por_placa[num]\n        sub = 1\n\n        for delantera, trasera, nombre_tipo in PARES_TIPO:\n            fotos_d = tipos.get(delantera, [])\n            fotos_t = tipos.get(trasera, [])\n            if not fotos_d and not fotos_t:\n                continue\n            aviso = \"\"\n            if len(fotos_d) != len(fotos_t):\n                aviso = f\"AVISO: {len(fotos_d)} delanteras vs {len(fotos_t)} traseras\"\n                avisos.append((num, nombre_tipo, aviso))\n            n_unidades = max(len(fotos_d), len(fotos_t))\n            for i in range(n_unidades):\n                fd = fotos_d[i] if i < len(fotos_d) else None\n                ft = fotos_t[i] if i < len(fotos_t) else None\n                fotos_asignadas = [f for f in [fd, ft] if f]\n                titulo = f\"{str(marca or '').upper()} {modelo} {TIPO_TEXTO_TITULO.get(nombre_tipo, nombre_tipo.upper())}\"\n                ws_new.append([num, sub, marca, modelo, n_serie, precio, ubic, notas,\n                                nombre_tipo, titulo, \";\".join(fotos_asignadas), aviso])\n                sub += 1\n\n        variadas = tipos.get(\"Variadas\", []) + tipos.get(\"SinClasificar\", [])\n        if variadas:\n            titulo = f\"{str(marca or '').upper()} {modelo} {TIPO_TEXTO_TITULO['PACK / Variadas']}\"\n            ws_new.append([num, sub, marca, modelo, n_serie, precio, ubic, notas,\n                            \"PACK / Variadas\", titulo, \";\".join(variadas), \"Revisar - fotos variadas o sin clasificar\"])\n            avisos.append((num, \"Variadas/SinClasificar\", f\"{len(variadas)} fotos\"))\n            sub += 1\n\n    wb_new.save(RUTA_SALIDA)\n\n    print(f\"Listo. Generado '{RUTA_SALIDA}' con {ws_new.max_row - 1} anuncios.\")\n    if avisos:\n        print(\"\\nAvisos para revisar a mano:\")\n        for a in avisos:\n            print(\" \", a)\n    print(\"\\nAhora puedes ejecutar: py subir_placas_final.py\")\n    input(\"\\nPulsa Enter para salir...\")\n\n\nif __name__ == \"__main__\":\n    main()\n";

const EJECUTAR_BAT = `@echo off
chcp 65001 >nul
title Catalogo de Placas - Proceso automatico (Creado por MCG)
echo ==========================================
echo   CATALOGO DE PLACAS - Proceso automatico
echo   Creado por MCG
echo ==========================================
echo.
echo Instalando dependencias necesarias (solo hace falta la primera vez)...
py -m pip install --quiet --upgrade pip
py -m pip install --quiet playwright openpyxl pillow
py -m playwright install chromium
echo.
echo ------------------------------------------
echo Paso 1 de 2: generando el catalogo final...
echo ------------------------------------------
py generar_catalogo_final.py
echo.
echo ------------------------------------------
echo Paso 2 de 2: subiendo los anuncios a la web
echo ------------------------------------------
py subir_placas_final.py
echo.
echo Proceso terminado.
pause
`;

function CatalogoPlacas() {
  const [items, setItems] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [form, setForm] = useState(emptyForm);
  const [editingId, setEditingId] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [photoError, setPhotoError] = useState("");
  const [ocrRunning, setOcrRunning] = useState(false);
  const [showSplash, setShowSplash] = useState(true);
  const [showInstrucciones, setShowInstrucciones] = useState(false);
  const [procesando, setProcesando] = useState(false);
  const [resultadoPaquete, setResultadoPaquete] = useState(null);
  const cameraInputRef = useRef(null);
  const galleryInputRef = useRef(null);

  useEffect(() => {
    const t = setTimeout(() => setShowSplash(false), 1700);
    return () => clearTimeout(t);
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const { value } = await Preferences.get({ key: STORAGE_KEY });
        if (value) setItems(JSON.parse(value));
      } catch (err) {}
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    Preferences.set({ key: STORAGE_KEY, value: JSON.stringify(items) }).catch(() => {});
  }, [items, loaded]);

  useEffect(() => {
    getOcrWorker().catch(() => {});
  }, []);

  const runSmartDetection = async (dataUrl) => {
    setOcrRunning(true);
    setPhotoError("");
    try {
      const barcode = await detectBarcode(dataUrl);
      const worker = await getOcrWorker();
      const { data } = await worker.recognize(dataUrl);
      const text = (data && data.text) ? data.text.trim() : "";
      const guess = parseLabelText(text);

      setForm((f) => ({
        ...f,
        marca: f.marca || guess.marca || f.marca,
        modelo: f.modelo || guess.modelo || f.modelo,
        serie: f.serie || barcode || guess.serie || f.serie,
        notas: f.notas || (text ? `Texto detectado (revisa):\n${text.slice(0, 200)}` : f.notas),
      }));

      if (!text && !barcode) {
        setPhotoError("No se detecto texto ni codigo legible en esta foto. Prueba con mas luz o mas de cerca.");
      }
    } catch (err) {
      setPhotoError("Error del lector: " + (err && err.message ? err.message : String(err)));
    }
    setOcrRunning(false);
  };

  const handlePhoto = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      const dataUrl = await resizeImage(file);
      setForm((f) => ({ ...f, fotos: [...f.fotos, { url: dataUrl, tipo: null }] }));
      setPhotoError("");
      runSmartDetection(dataUrl);
    } catch (err) {
      setPhotoError("No se pudo cargar la foto, prueba de nuevo.");
    }
    e.target.value = "";
  };

  const etiquetarFoto = (index, tipo) => {
    setForm((f) => ({
      ...f,
      fotos: f.fotos.map((foto, i) => (i === index ? { ...foto, tipo } : foto)),
    }));
  };

  const removePhoto = (index) => {
    setForm((f) => ({ ...f, fotos: f.fotos.filter((_, i) => i !== index) }));
  };

  const resetForm = () => {
    setForm(emptyForm);
    setEditingId(null);
    setShowForm(false);
    setPhotoError("");
  };

  const saveItem = () => {
    if (!form.marca && !form.modelo) return;
    if (editingId) {
      setItems((prev) => prev.map((it) => (it.id === editingId ? { ...it, ...form } : it)));
    } else {
      setItems((prev) => [...prev, { id: Date.now().toString(), ...form }]);
    }
    resetForm();
  };

  const editItem = (item) => {
    setForm({
      marca: item.marca, modelo: item.modelo, serie: item.serie,
      precio: item.precio, ubicacion: item.ubicacion, notas: item.notas,
      fotos: item.fotos || [],
    });
    setEditingId(item.id);
    setShowForm(true);
  };

  const deleteItem = (id) => setItems((prev) => prev.filter((it) => it.id !== id));

  const totalFotos = items.reduce((acc, it) => acc + (it.fotos ? it.fotos.length : 0), 0);
  const sinClasificar = items.reduce(
    (acc, it) => acc + (it.fotos ? it.fotos.filter((f) => !f.tipo).length : 0),
    0
  );

  const construirFilasExcel = () => items.map((it, i) => ({
    "Nº": i + 1, Marca: it.marca, Modelo: it.modelo,
    "Nº Serie": it.serie, Precio: it.precio, "Ubicación almacén": it.ubicacion, Notas: it.notas,
  }));

  const generarExcelBase64 = () => {
    const ws = XLSX.utils.json_to_sheet(construirFilasExcel());
    ws["!cols"] = [{ wch: 5 }, { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 10 }, { wch: 16 }, { wch: 30 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Placas");
    return XLSX.write(wb, { type: "base64", bookType: "xlsx" });
  };

  const compartirOBajar = async (base64, fileName, mime, titulo, texto) => {
    if (Capacitor.isNativePlatform()) {
      const result = await Filesystem.writeFile({ path: fileName, data: base64, directory: Directory.Cache });
      await Share.share({ title: titulo, text: texto, url: result.uri, dialogTitle: titulo });
    } else {
      const link = document.createElement("a");
      link.href = `data:${mime};base64,${base64}`;
      link.download = fileName;
      link.click();
    }
  };

  const exportExcel = async () => {
    if (!items.length) return;
    setProcesando(true);
    try {
      const base64 = generarExcelBase64();
      await compartirOBajar(base64, `catalogo_${Date.now()}.xlsx`,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Excel del catálogo", "Excel con marca, modelo y demás datos");
    } catch (err) {
      setPhotoError("No se pudo generar el Excel.");
    }
    setProcesando(false);
  };

  const exportFotos = async () => {
    const withPhotos = items.filter((it) => it.fotos && it.fotos.length);
    if (!withPhotos.length) {
      setPhotoError("No hay fotos guardadas todavía para exportar.");
      return;
    }
    setProcesando(true);
    try {
      const zip = new JSZip();
      items.forEach((it, i) => {
        if (!it.fotos || !it.fotos.length) return;
        const num = String(i + 1).padStart(3, "0");
        const safeModelo = (it.modelo || "sin_modelo").replace(/[^a-zA-Z0-9\-]/g, "_");
        it.fotos.forEach((foto, j) => {
          const tipo = foto.tipo || "SinClasificar";
          const base64Data = foto.url.split(",")[1];
          zip.file(`${num}_${safeModelo}_${tipo}_${j + 1}.jpg`, base64Data, { base64: true });
        });
      });
      const zipBase64 = await zip.generateAsync({ type: "base64" });
      await compartirOBajar(zipBase64, `fotos_placas_${Date.now()}.zip`, "application/zip",
        "Fotos del catálogo", "Cada foto lleva el número, modelo y tipo en el nombre");
    } catch (err) {
      setPhotoError("No se pudo generar el ZIP de fotos.");
    }
    setProcesando(false);
  };

  const procesarYExportar = async () => {
    if (!items.length) return;
    setProcesando(true);
    setResultadoPaquete(null);
    try {
      const zip = new JSZip();
      const carpetaFotos = zip.folder("fotos");

      items.forEach((it, i) => {
        if (!it.fotos || !it.fotos.length) return;
        const num = String(i + 1).padStart(3, "0");
        const safeModelo = (it.modelo || "sin_modelo").replace(/[^a-zA-Z0-9\-]/g, "_");
        it.fotos.forEach((foto, j) => {
          const tipo = foto.tipo || "SinClasificar";
          const base64Data = foto.url.split(",")[1];
          carpetaFotos.file(`${num}_${safeModelo}_${tipo}_${j + 1}.jpg`, base64Data, { base64: true });
        });
      });

      zip.file("catalogo.xlsx", generarExcelBase64(), { base64: true });
      zip.file("generar_catalogo_final.py", GENERAR_CATALOGO_PY);
      zip.file("subir_placas_final.py", SUBIR_PLACAS_PY);
      zip.file("EJECUTAR.bat", EJECUTAR_BAT);
      zip.file("INSTRUCCIONES.txt", TEXTO_INSTRUCCIONES);

      const zipBase64 = await zip.generateAsync({ type: "base64" });
      const fileName = `paquete_placas_${Date.now()}.zip`;

      if (Capacitor.isNativePlatform()) {
        const result = await Filesystem.writeFile({
          path: fileName,
          data: zipBase64,
          directory: Directory.Cache,
        });
        setResultadoPaquete({ uri: result.uri, fileName });
      } else {
        const link = document.createElement("a");
        link.href = `data:application/zip;base64,${zipBase64}`;
        link.download = fileName;
        link.click();
        setResultadoPaquete({ uri: null, fileName });
      }
    } catch (err) {
      setPhotoError("No se pudo generar el paquete. Vuelve a intentarlo.");
    }
    setProcesando(false);
  };

  const compartirPaquete = async () => {
    if (!resultadoPaquete || !resultadoPaquete.uri) return;
    await Share.share({
      title: "Paquete de placas",
      text: "Fotos + Excel listos para el ordenador. Lee INSTRUCCIONES.txt dentro del ZIP.",
      url: resultadoPaquete.uri,
      dialogTitle: "Enviar paquete al ordenador",
    });
  };

  return (
    <div style={styles.page}>
      {showSplash && (
        <div style={styles.splash}>
          <img src="icon.png" alt="Electrónica Imperial" style={styles.splashLogo} />
          <div style={styles.splashTitle}>Catálogo de Placas</div>
          <div style={styles.splashSub}>Electrónica Imperial</div>
          <div style={styles.splashWatermark}>Creado por MCG</div>
        </div>
      )}

      {showInstrucciones && (
        <div style={styles.overlay} onClick={() => setShowInstrucciones(false)}>
          <div style={styles.sheet} onClick={(e) => e.stopPropagation()}>
            <div style={styles.sheetHeader}>
              <span style={styles.sheetTitle}>Cómo usar la app</span>
              <button style={styles.closeBtn} onClick={() => setShowInstrucciones(false)}><Icon size={20} color="#5b6270">✕</Icon></button>
            </div>
            <div style={styles.instruccionesTexto}>
              <p><b>1.</b> Toca "Añadir placa" y luego "Hacer foto".</p>
              <p><b>2.</b> Después de cada foto, toca el botón que diga qué es (Base Delantera, Fuente Trasera, T-Con, etc.). Así no hay que clasificar nada después.</p>
              <p><b>3.</b> Repite hasta tener todas las fotos de esa placa (normalmente 2 por cada pieza: delante y detrás).</p>
              <p><b>4.</b> Rellena marca, modelo y nº de serie (o deja que la app lo intente leer sola de la foto).</p>
              <p><b>5.</b> Toca "Añadir a la lista" y repite con la siguiente placa.</p>
              <p><b>6.</b> Cuando termines todas: pulsa el botón grande dorado <b>"Paquete completo (.bat)"</b>. Esto genera un único ZIP con las fotos, el Excel, y un archivo "EJECUTAR.bat" que hace TODO solo en el ordenador con un doble clic (no hace falta escribir ningún comando).</p>
              <p><b>7.</b> Envía ese ZIP a tu ordenador (Drive, USB, WhatsApp a ti mismo), descomprímelo, y haz doble clic en "EJECUTAR.bat".</p>
              <p style={{ marginTop: 10, color: "#8a8f98" }}>Los botones pequeños "Excel" y "Fotos" sirven solo si quieres esos archivos por separado, para otros usos.</p>
            </div>
          </div>
        </div>
      )}

      {procesando && (
        <div style={styles.overlay}>
          <div style={styles.procesandoBox}>
            <div style={styles.spinner} />
            <div style={styles.procesandoTexto}>Generando el paquete…</div>
            <div style={styles.procesandoSub}>Fotos, Excel e instrucciones</div>
          </div>
        </div>
      )}

      {resultadoPaquete && (
        <div style={styles.overlay} onClick={() => setResultadoPaquete(null)}>
          <div style={styles.sheet} onClick={(e) => e.stopPropagation()}>
            <div style={styles.resultadoCheck}>✅</div>
            <div style={styles.resultadoTitulo}>Paquete listo</div>
            <div style={styles.resultadoTexto}>
              {items.length} placas · {totalFotos} fotos, junto con el Excel, los
              2 programas y el archivo <b>EJECUTAR.bat</b>. En el ordenador solo
              hay que descomprimir y hacer doble clic en EJECUTAR.bat.
            </div>
            {resultadoPaquete.uri ? (
              <button style={styles.saveBtn} onClick={compartirPaquete}>
                <Icon color="#fff">📤</Icon><span>Enviar ZIP al ordenador</span>
              </button>
            ) : (
              <div style={styles.resultadoTexto}>Descarga iniciada: {resultadoPaquete.fileName}</div>
            )}
            <button style={styles.linkBtn} onClick={() => setResultadoPaquete(null)}>Cerrar</button>
          </div>
        </div>
      )}

      <header style={styles.header}>
        <div style={styles.headerTop}>
          <div style={styles.headerBrand}>
            <img src="icon.png" alt="" style={styles.headerLogo} />
            <div>
              <div style={styles.plateTitle}>Catálogo de Placas</div>
              <div style={styles.plateSerial}>{String(items.length).padStart(3, "0")} placas · {totalFotos} fotos</div>
            </div>
          </div>
          <button style={styles.infoBtn} onClick={() => setShowInstrucciones(true)}>
            <Icon size={17} color="#fff">ℹ</Icon>
          </button>
        </div>
        <div style={styles.watermark}>Electrónica Imperial · Creado por MCG</div>
      </header>

      <main style={styles.main}>
        {items.length === 0 && !showForm && (
          <div style={styles.empty}>
            <div style={styles.emptyCircle}><Icon size={30} color="#c7a44a">📦</Icon></div>
            <p style={styles.emptyText}>Aún no has añadido ninguna placa</p>
            <p style={styles.emptySub}>Toca "Añadir placa" abajo y empieza a fotografiar</p>
          </div>
        )}

        <div style={styles.list}>
          {items.map((it, idx) => {
            const fotos = it.fotos || [];
            const sinTipo = fotos.filter((f) => !f.tipo).length;
            return (
              <div key={it.id} style={styles.card}>
                <div style={styles.cardPhoto}>
                  {fotos[0] ? (
                    <>
                      <img src={fotos[0].url} alt={it.modelo} style={styles.thumb} />
                      {fotos.length > 1 && <div style={styles.photoCountBadge}>+{fotos.length - 1}</div>}
                    </>
                  ) : (
                    <div style={styles.thumbPlaceholder}><Icon size={20} color="#c9cdd4">📷</Icon></div>
                  )}
                </div>
                <div style={styles.cardBody}>
                  <div style={styles.cardIndex}>#{String(idx + 1).padStart(3, "0")}</div>
                  <div style={styles.cardMarca}>{it.marca || "Sin marca"}</div>
                  <div style={styles.cardModelo}>{it.modelo || "Sin modelo"}</div>
                  {it.serie && <div style={styles.cardSerie}>Serie: {it.serie}</div>}
                  {sinTipo > 0 && <div style={styles.avisoSinTipo}>⚠ {sinTipo} foto(s) sin clasificar</div>}
                  <div style={styles.cardBottomRow}>
                    <span style={styles.cardPrecio}>{it.precio ? `${it.precio} €` : "—"}</span>
                    <div style={styles.cardActions}>
                      <button style={styles.iconBtn} onClick={() => editItem(it)}><Icon color="#5b6270">✎</Icon></button>
                      <button style={styles.iconBtn} onClick={() => deleteItem(it.id)}><Icon color="#c0483c">🗑</Icon></button>
                    </div>
                  </div>
                  {it.ubicacion && <div style={styles.cardUbicacion}>📍 {it.ubicacion}</div>}
                </div>
              </div>
            );
          })}
        </div>
      </main>

      {showForm && (
        <div style={styles.overlay}>
          <div style={styles.sheet}>
            <div style={styles.sheetHeader}>
              <span style={styles.sheetTitle}>{editingId ? "Editar placa" : "Nueva placa"}</span>
              <button style={styles.closeBtn} onClick={resetForm}><Icon size={20} color="#5b6270">✕</Icon></button>
            </div>

            <div style={styles.photoGallery}>
              {form.fotos.map((foto, idx) => (
                <div key={idx} style={styles.photoThumbWrap}>
                  <img src={foto.url} alt={`foto ${idx + 1}`} style={styles.photoThumbImg} />
                  {foto.tipo && (
                    <div style={styles.photoTipoTag}>
                      {(TIPOS_FOTO.find((t) => t.id === foto.tipo) || {}).label}
                    </div>
                  )}
                  <button type="button" style={styles.photoRemoveBtn} onClick={() => removePhoto(idx)}>
                    <Icon size={12} color="#fff">✕</Icon>
                  </button>
                </div>
              ))}
              <button type="button" style={styles.photoAddTile} onClick={() => cameraInputRef.current && cameraInputRef.current.click()}>
                <Icon size={22} color="#c07a12">📷</Icon>
                <span style={styles.photoAddTileText}>{form.fotos.length ? "Añadir otra" : "Hacer foto"}</span>
              </button>
            </div>

            {form.fotos.some((f) => !f.tipo) && (
              <div style={styles.tagBox}>
                <div style={styles.tagBoxTitle}>¿Qué es la última foto? Toca una opción:</div>
                <div style={styles.tagGrid}>
                  {TIPOS_FOTO.map((t) => (
                    <button
                      key={t.id}
                      type="button"
                      style={{ ...styles.tagBtn, background: t.color }}
                      onClick={() => {
                        const idx = form.fotos.map((f) => !f.tipo).lastIndexOf(true);
                        if (idx !== -1) etiquetarFoto(idx, t.id);
                      }}
                    >
                      {t.label}
                    </button>
                  ))}
                </div>
              </div>
            )}

            <div style={styles.photoRow}>
              <button type="button" style={styles.photoAltBtn} onClick={() => cameraInputRef.current && cameraInputRef.current.click()}>
                <Icon size={14} color="#5b6270">📷</Icon><span>Cámara</span>
              </button>
              <button type="button" style={styles.photoAltBtn} onClick={() => galleryInputRef.current && galleryInputRef.current.click()}>
                <Icon size={14} color="#5b6270">🖼</Icon><span>Elegir de galería</span>
              </button>
            </div>
            {photoError && <div style={styles.photoErrorText}>{photoError}</div>}
            {ocrRunning && <div style={styles.ocrStatus}>🔍 Leyendo la etiqueta...</div>}
            <input ref={cameraInputRef} type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={handlePhoto} />
            <input ref={galleryInputRef} type="file" accept="image/*" style={{ display: "none" }} onChange={handlePhoto} />

            <div style={styles.fieldGroup}>
              <label style={styles.label}>Marca</label>
              <input style={styles.input} value={form.marca} onChange={(e) => setForm((f) => ({ ...f, marca: e.target.value }))} placeholder="Ej. Philips" />
            </div>
            <div style={styles.fieldGroup}>
              <label style={styles.label}>Modelo</label>
              <input style={styles.input} value={form.modelo} onChange={(e) => setForm((f) => ({ ...f, modelo: e.target.value }))} placeholder="Ej. 50PUS7406/12" />
            </div>
            <div style={styles.fieldGroup}>
              <label style={styles.label}>Nº de serie / etiqueta</label>
              <input style={styles.input} value={form.serie} onChange={(e) => setForm((f) => ({ ...f, serie: e.target.value }))} placeholder="Mira la foto y escríbelo aquí" />
            </div>
            <div style={styles.fieldGroup}>
              <label style={styles.label}>Precio (€) — opcional</label>
              <div style={styles.priceRow}>
                <input style={{ ...styles.input, flex: 1 }} value={form.precio} onChange={(e) => setForm((f) => ({ ...f, precio: e.target.value }))} placeholder="Se puede dejar en blanco" inputMode="decimal" />
                <button
                  type="button"
                  style={styles.priceCheckBtn}
                  disabled={!form.marca && !form.modelo}
                  onClick={() => {
                    const q = encodeURIComponent(`${form.marca} ${form.modelo}`.trim());
                    window.open(`https://www.google.com/search?tbm=shop&q=${q}`, "_blank");
                  }}
                >
                  <Icon size={13} color="#8a6d1c">🔎</Icon><span>Ver precios</span>
                </button>
              </div>
            </div>
            <div style={styles.fieldGroup}>
              <label style={styles.label}>Localización en almacén</label>
              <input style={styles.input} value={form.ubicacion} onChange={(e) => setForm((f) => ({ ...f, ubicacion: e.target.value }))} placeholder="Ej. Estantería 3, Caja B" />
            </div>
            <div style={styles.fieldGroup}>
              <label style={styles.label}>Notas</label>
              <textarea style={{ ...styles.input, minHeight: 60, resize: "none" }} value={form.notas} onChange={(e) => setForm((f) => ({ ...f, notas: e.target.value }))} placeholder="Opcional" />
            </div>

            <button style={styles.saveBtn} onClick={saveItem}>
              <Icon color="#fff">✓</Icon><span>{editingId ? "Guardar cambios" : "Añadir a la lista"}</span>
            </button>
          </div>
        </div>
      )}

      {!showForm && (
        <div style={styles.footerWrap}>
          {sinClasificar > 0 && (
            <div style={styles.footerAviso}>⚠ {sinClasificar} foto(s) sin clasificar todavía</div>
          )}
          <button style={styles.addBtn} onClick={() => setShowForm(true)}>
            <Icon color="#fff">➕</Icon><span>Añadir placa</span>
          </button>
          <div style={styles.exportRow}>
            <button style={{ ...styles.exportSmallBtn, opacity: items.length ? 1 : 0.45 }} onClick={exportExcel} disabled={!items.length || procesando}>
              <Icon size={14} color="#3a3f4a">📊</Icon><span>Excel</span>
            </button>
            <button style={{ ...styles.exportSmallBtn, opacity: items.length ? 1 : 0.45 }} onClick={exportFotos} disabled={!items.length || procesando}>
              <Icon size={14} color="#3a3f4a">🖼</Icon><span>Fotos</span>
            </button>
          </div>
          <button
            style={{ ...styles.procesarBtn, opacity: items.length ? 1 : 0.45 }}
            onClick={procesarYExportar}
            disabled={!items.length || procesando}
          >
            <Icon color="#1c1f26" size={18}>⚙</Icon>
            <span>{procesando ? "Procesando..." : "Paquete completo (.bat)"}</span>
          </button>
        </div>
      )}
    </div>
  );
}

const NAVY = "#0d1b34";
const GOLD = "#d4a94c";
const GOLD_SOFT = "#f2dfa8";
const CREAM = "#f6f3ec";

const styles = {
  splash: { position: "fixed", inset: 0, background: NAVY, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", zIndex: 100, gap: 8 },
  splashLogo: { width: 116, height: 116, borderRadius: 26, boxShadow: "0 8px 30px rgba(0,0,0,0.45)" },
  splashTitle: { color: "#fff", fontSize: 19, fontWeight: 800, marginTop: 14, letterSpacing: 0.3 },
  splashSub: { color: GOLD, fontSize: 12, fontWeight: 600, letterSpacing: 1.5, textTransform: "uppercase" },
  splashWatermark: { position: "absolute", bottom: 26, color: "#5b6b8f", fontSize: 11, letterSpacing: 1.2 },

  page: { minHeight: "100vh", background: CREAM, color: "#1c1f26", display: "flex", flexDirection: "column" },

  header: { background: NAVY, padding: "18px 16px 14px", borderBottom: `3px solid ${GOLD}` },
  headerTop: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  headerBrand: { display: "flex", alignItems: "center", gap: 10 },
  headerLogo: { width: 38, height: 38, borderRadius: 10 },
  plateTitle: { fontSize: 16, fontWeight: 800, color: "#fff" },
  plateSerial: { fontSize: 11, color: GOLD_SOFT, marginTop: 1, fontWeight: 600 },
  infoBtn: { border: `1px solid ${GOLD}`, background: "rgba(212,169,76,0.12)", borderRadius: 18, width: 32, height: 32, display: "flex", alignItems: "center", justifyContent: "center" },
  watermark: { fontSize: 10, color: "#5b6b8f", marginTop: 10, letterSpacing: 0.6 },

  main: { flex: 1, padding: "14px 12px 168px" },
  empty: { marginTop: 70, textAlign: "center", display: "flex", flexDirection: "column", alignItems: "center", gap: 6 },
  emptyCircle: { width: 64, height: 64, borderRadius: 32, background: "#fff", border: `2px solid ${GOLD_SOFT}`, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 6 },
  emptyText: { color: "#3a3f4a", fontSize: 15, fontWeight: 700, margin: 0 },
  emptySub: { color: "#8a8f98", fontSize: 13, margin: 0 },

  list: { display: "flex", flexDirection: "column", gap: 10 },
  card: { display: "flex", background: "#fff", borderRadius: 12, overflow: "hidden", boxShadow: "0 2px 8px rgba(13,27,52,0.08)", border: "1px solid #eee6d3" },
  cardPhoto: { width: 82, flexShrink: 0, background: "#f3efe4", position: "relative" },
  thumb: { width: "100%", height: "100%", objectFit: "cover" },
  thumbPlaceholder: { width: "100%", height: "100%", display: "flex", alignItems: "center", justifyContent: "center" },
  cardBody: { flex: 1, padding: "9px 11px" },
  cardIndex: { fontFamily: "'Courier New', monospace", fontSize: 10, color: "#b8bdc6" },
  cardMarca: { fontSize: 13, fontWeight: 800, textTransform: "uppercase", letterSpacing: 0.3, color: NAVY },
  cardModelo: { fontSize: 13, color: "#5b6270" },
  cardSerie: { fontSize: 11, color: "#8a8f98", fontFamily: "'Courier New', monospace" },
  avisoSinTipo: { fontSize: 10, color: "#a15c00", fontWeight: 700, marginTop: 3 },
  cardBottomRow: { display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 5 },
  cardPrecio: { fontSize: 14, fontWeight: 800, color: "#1f6d3f" },
  cardActions: { display: "flex", gap: 4 },
  iconBtn: { border: "none", background: "transparent", padding: 4 },
  cardUbicacion: { fontSize: 11, color: "#8a6d1c", marginTop: 4, fontWeight: 700 },
  photoCountBadge: { position: "absolute", bottom: 3, right: 3, background: "rgba(13,27,52,0.85)", color: "#fff", fontSize: 10, fontWeight: 700, borderRadius: 6, padding: "1px 5px" },

  footerWrap: { position: "fixed", bottom: 0, left: 0, right: 0, display: "flex", flexDirection: "column", gap: 8, padding: 12, background: `linear-gradient(to top, ${CREAM} 75%, transparent)` },
  footerAviso: { textAlign: "center", fontSize: 11, color: "#a15c00", fontWeight: 700, background: "#fff3de", border: "1px solid #f0d99a", borderRadius: 8, padding: "5px 0" },
  addBtn: { width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: NAVY, color: "#fff", border: "none", borderRadius: 12, padding: "13px 0", fontSize: 14, fontWeight: 700 },
  exportRow: { display: "flex", gap: 8 },
  exportSmallBtn: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 5, background: "#fff", border: "1px solid #e2dcc8", borderRadius: 10, padding: "9px 0", fontSize: 12, color: "#3a3f4a", fontWeight: 700 },
  procesarBtn: { width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: GOLD, color: "#1c1f26", border: "none", borderRadius: 12, padding: "14px 0", fontSize: 14, fontWeight: 800, boxShadow: "0 4px 14px rgba(212,169,76,0.4)" },

  overlay: { position: "fixed", inset: 0, background: "rgba(13,27,52,0.55)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 20 },
  sheet: { background: "#fff", width: "100%", borderRadius: "18px 18px 0 0", padding: "18px 18px 26px", maxHeight: "88vh", overflowY: "auto" },
  sheetHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 },
  sheetTitle: { fontSize: 16, fontWeight: 800, color: NAVY },
  closeBtn: { border: "none", background: "transparent" },
  instruccionesTexto: { fontSize: 13, lineHeight: 1.6, color: "#3a3f4a" },

  photoGallery: { display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 10 },
  photoThumbWrap: { position: "relative", width: 84, height: 84, borderRadius: 10, overflow: "hidden" },
  photoThumbImg: { width: "100%", height: "100%", objectFit: "cover" },
  photoTipoTag: { position: "absolute", bottom: 0, left: 0, right: 0, background: "rgba(13,27,52,0.8)", color: "#fff", fontSize: 8, fontWeight: 700, textAlign: "center", padding: "2px 1px", lineHeight: 1.1 },
  photoRemoveBtn: { position: "absolute", top: 3, right: 3, width: 20, height: 20, borderRadius: "50%", background: "rgba(0,0,0,0.55)", border: "none", display: "flex", alignItems: "center", justifyContent: "center", padding: 0 },
  photoAddTile: { width: 84, height: 84, background: "#f6f3ec", border: `2px dashed ${GOLD_SOFT}`, borderRadius: 10, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 3, padding: 4 },
  photoAddTileText: { fontSize: 10, color: "#5b6270", fontWeight: 700, textAlign: "center", lineHeight: 1.2 },

  tagBox: { background: "#fff7e6", border: "1px solid #f0d99a", borderRadius: 10, padding: 10, marginBottom: 12 },
  tagBoxTitle: { fontSize: 12, fontWeight: 800, color: "#8a6d1c", marginBottom: 8 },
  tagGrid: { display: "flex", flexWrap: "wrap", gap: 6 },
  tagBtn: { flex: "1 1 46%", border: "none", borderRadius: 8, padding: "10px 4px", fontSize: 11, fontWeight: 800, color: "#fff", textAlign: "center" },

  photoRow: { display: "flex", gap: 8, marginBottom: 10 },
  photoAltBtn: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 5, background: "#f3efe4", border: "1px solid #e2dcc8", borderRadius: 8, padding: "8px 0", fontSize: 12, color: "#5b6270", fontWeight: 700 },
  photoErrorText: { color: "#c0483c", fontSize: 12, marginBottom: 8 },
  ocrStatus: { color: "#1f6d3f", fontSize: 12, marginBottom: 8, fontWeight: 700 },

  priceRow: { display: "flex", gap: 6, alignItems: "center" },
  priceCheckBtn: { display: "flex", alignItems: "center", gap: 4, background: "#fdf4e0", border: "1px solid #ecd28f", borderRadius: 8, padding: "9px 10px", fontSize: 12, color: "#8a6d1c", fontWeight: 800, whiteSpace: "nowrap" },

  fieldGroup: { marginBottom: 10 },
  label: { fontSize: 12, color: "#5b6270", fontWeight: 700, display: "block", marginBottom: 4 },
  input: { width: "100%", padding: "10px 12px", border: "1px solid #e2dcc8", borderRadius: 8, fontSize: 14, boxSizing: "border-box", fontFamily: "system-ui, sans-serif" },
  saveBtn: { width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#1f6d3f", color: "#fff", border: "none", borderRadius: 10, padding: "13px 0", fontSize: 14, fontWeight: 800, marginTop: 4 },
  linkBtn: { width: "100%", textAlign: "center", background: "transparent", border: "none", color: "#8a8f98", fontSize: 13, fontWeight: 700, padding: "12px 0 0" },

  procesandoBox: { background: "#fff", borderRadius: 16, padding: "34px 30px", display: "flex", flexDirection: "column", alignItems: "center", gap: 10, margin: "auto" },
  spinner: { width: 34, height: 34, borderRadius: "50%", border: `4px solid ${GOLD_SOFT}`, borderTopColor: GOLD, animation: "girar 0.8s linear infinite" },
  procesandoTexto: { fontSize: 14, fontWeight: 800, color: NAVY },
  procesandoSub: { fontSize: 12, color: "#8a8f98" },

  resultadoCheck: { fontSize: 40, textAlign: "center", marginBottom: 6 },
  resultadoTitulo: { fontSize: 17, fontWeight: 800, textAlign: "center", color: NAVY, marginBottom: 6 },
  resultadoTexto: { fontSize: 13, color: "#5b6270", textAlign: "center", lineHeight: 1.5, marginBottom: 14 },
};

const styleTag = document.createElement("style");
styleTag.innerHTML = "@keyframes girar { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }";
document.head.appendChild(styleTag);

const root = createRoot(document.getElementById("root"));
root.render(<CatalogoPlacas />);
