# NEXOVA · Organizador Fotográfico de Inventarios (versión web)

Organiza automáticamente las fotos de un inventario de activos fijos en
carpetas **Familia / Subfamilia / Código**, y genera el Excel de conciliación
y el Informe PDF de cobertura fotográfica.

**Todo corre en la computadora del consultor. Cero backend, cero servidor,
ninguna foto sale a internet.**

---

## ANTES DE TODO: el test de 30 segundos

En **la laptop de un consultor** (no en la tuya), abre Chrome →
`F12` → pestaña **Console** → pega esto → Enter:

```js
'showDirectoryPicker' in window
```

- `true`  → adelante.
- `false` → esta herramienta no va a funcionar ahí. No publiques nada todavía.

---

## Publicar (no requiere instalar nada)

1. Entra a **vercel.com** y crea una cuenta (gratis).
2. `Add New…` → `Project` → **Deploy without Git** / arrastra esta carpeta completa.
3. Vercel te devuelve una URL tipo `https://organizador-nexova.vercel.app`.
4. Esa URL es la herramienta. Se la pasas a los consultores. No instalan nada.

**Alternativas equivalentes:** Netlify Drop (`app.netlify.com/drop`), Cloudflare Pages.
Cualquier hosting estático con HTTPS sirve. Costo: **S/ 0**.

### Por qué no funciona con doble clic en `index.html`
El acceso a carpetas del disco exige un **origen seguro (HTTPS)**. Con `file://`
el navegador bloquea la API. Publicarla no es un capricho: es un requisito técnico.

---

## Requisitos del consultor

| | |
|---|---|
| Navegador | **Chrome o Edge de escritorio** |
| No funciona en | Firefox, Safari, celular, tablet |
| Instalación | **Ninguna** |
| Permisos de admin | **No** |

La primera vez, Chrome pedirá permiso para leer/escribir en la carpeta.
Es una ventana del sistema; el consultor solo acepta.

---

## Cómo se usa

1. **Inventario** — el Excel del cliente. La fila de encabezado se detecta sola.
2. **Catálogo** — el Excel de familias/subfamilias. Si el inventario ya trae la
   familia, marca la casilla y te saltas este paso.
3. **Carpetas** — origen (las fotos) y destino (vacía).
4. **Cómo organizar** — COPIAR o MOVER, niveles de carpeta, nombre del cliente.
5. **SIMULAR** → revisa el resultado → **EJECUTAR**.

### Nombre de las fotos
El archivo debe llamarse como el código del activo: `EC-001.jpg`.
Fotos múltiples: `EC-001-1.jpg`, `EC-001_2.jpg`, `EC-001 (3).jpg`.

---

## Entregables que produce

| Archivo | Para quién |
|---|---|
| Carpetas organizadas | Uso interno / entrega al cliente |
| `Conciliacion.xlsx` | 6 hojas: resumen, detalle, fotos sin código, activos sin foto, por familia, alertas |
| `Informe_Cobertura_Fotografica.pdf` | **El cliente.** Portada NEXOVA, cobertura, gráficos, listado de activos pendientes de fotografiar |

Al ejecutar, los dos reportes quedan guardados dentro de la carpeta destino.

---

## Antes de soltar 10,000 fotos

Copia **30 fotos** a una carpeta aparte, con un Excel recortado, y corre el
proceso completo en modo **COPIAR**. Recién cuando eso salga limpio, ve por
el inventario entero.

**El error más probable no es del programa: es un mapeo de columna equivocado.**
Por eso existe el botón SIMULAR y por eso la pantalla te avisa en rojo si
muchos activos quedaron *sin familia* — casi siempre significa que la columna
de ENLACE está mal elegida.

---

## Estructura

```
index.html          Interfaz
js/motor.js         Emparejamiento y cruce de datos (sin DOM: testeable en Node)
js/archivos.js      Acceso al disco (File System Access API)
js/reportes.js      Gráficos, Excel y PDF
js/app.js           Orquestación de la pantalla
vendor/             Librerías locales (no dependen de ningún CDN)
test-motor.mjs      14 casos de aceptación → `node test-motor.mjs`
```

El motor está separado de la interfaz a propósito: se puede probar sin
navegador y se puede reutilizar el día que migres a una versión con build.

---

NEXOVA Software Empresarial · nexova.pe · +51 949 287 897
