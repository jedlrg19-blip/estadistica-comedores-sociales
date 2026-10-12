# Guía de administración del tablero

## Qué archivo alimenta cada número
Pestaña **Fuentes y trazabilidad** del tablero: archivo, hoja, fecha, huella SHA-256, responsable, conciliación entre archivos y catálogo número → archivo.
Cada gráfica tiene un chip `ⓘ archivo.xlsx` (lleva a esa ficha) y un botón `⤓ CSV` con sus datos.

| Archivo en `data/` | Alimenta |
|---|---|
| `cuota.xlsx` (Cuota por Semana) | Velocímetros del mes, gráficas semanales, top comedores, modalidad, contratación, desayuno/almuerzo |
| `dps.xlsx` (Ejecución DPS 2026) | Mapa, ranking por departamento, tendencia mensual, acumulado anual, hoja INTEGRADO (meta oficial) |
| `clasificador.xlsx` | Sexo / edad / pueblo |
| `semestral.xlsx` | Pagos a proveedores |

## Qué datos manda (precedencia)
1. **Archivos cargados por el administrador en su navegador** (borradores): solo los ve el admin hasta publicar.
2. **Lo publicado en `data/*.xlsx`** (con `data/publicado.json`, que guarda versión y huellas).
3. **`datos.json`**: respaldo que se usa **solo si no existe ningún `data/*.xlsx`**. Nunca tapa un archivo de `data/`.
Si un `data/*.xlsx` no se reconoce, el sitio **no** usa cifras viejas de `datos.json`: muestra «se está actualizando» al público y una advertencia en *Fuentes y trazabilidad* al administrador.

## Publicar datos desde el sitio (botón «Publicar datos»)
1. Entre por `https://<sitio>/admin/` (también `?editor=1`, que redirige a `/admin/`).
2. Cargue los Excel actualizados en «Fuentes de datos» y revise Resumen y Fuentes.
3. Presione **Publicar datos**: se hace **un commit** en GitHub con `data/*.xlsx` (los archivos cargados), `datos.json` y `data/publicado.json`. El diálogo espera a Cloudflare y avisa **«Ya está en vivo»** con un botón para recargar.
4. Token: GitHub → Settings → Developer settings → Fine-grained tokens. Solo este repositorio, permiso **Contents: Read and write**, caducidad corta. Se escribe en el diálogo; no queda en el repositorio.

## Actualizar datos (administrador, por Git)
1. Reemplace el Excel en `data/` **con el mismo nombre** (commit en GitHub → Cloudflare Pages publica solo).
2. `data/` tiene prioridad sobre `datos.json`; este solo se usa para lo que falte en `data/`.
3. Revise la pestaña *Fuentes* → *Conciliación*: ahí aparecen diferencias entre archivos.
4. Responsables/correos: edite `data/fuentes.json` (el botón «Solicitar actualización» abre un correo prellenado).
5. Meta anual oficial: `data/config.json` con `{"projRac": 9219150, "projFin": 0}` (ver advertencia en Conciliación).

## Modo administrador
Entre por `https://<sitio>/admin/`. Esa ruta activa el modo administrador (insignia amarilla «Administrador», clic para salir) y habilita: pestaña **Fuentes y trazabilidad**, descarga del Excel original, edición de proyección y publicar datos. Para el público la pestaña de fuentes no existe. `?editor=1` ya no funciona.

## Resumen ejecutivo
Primera pestaña: hallazgos generados con los datos cargados (cada uno cita archivo y hoja), avance contra el ritmo esperado y tabla por departamento. La meta anual usa, en este orden: `data/config.json` → meta vigente de la hoja INTEGRADO → total del mes × 12 (el origen se muestra en pantalla).

## Seguridad real (Cloudflare Zero Trust → Access)
Automatizado: `scripts/cloudflare-access.sh` (requiere `CLOUDFLARE_API_TOKEN` y `ADMIN_EMAILS` en el entorno; no guarda secretos).

El modo administrador solo oculta/mostrar controles en el navegador; **la seguridad real la da Cloudflare Access sobre `/admin*`**. Además, todo lo que está en `data/` es público por URL.
Para que solo el admin vea/descargue los originales:
1. Cloudflare Zero Trust → Access → Applications → *Self-hosted*.
2. Dominio del sitio, rutas `/admin*` y `/data/*`, política *Allow* solo para correos del admin.
3. Si `/data/*` se protege, el sitio público dejará de leer los Excel: publique entonces un paquete `datos.json` (botón «Publicar datos») para el público y deje los originales solo bajo Access.

## Consolidar el dps.xlsx desde el ZIP de ejecuciones (pestaña «Consolidar DPS»)

Los programas envían cada mes sus archivos crudos (Comedores, Beca Superior, Jóvenes Protagonistas, Bolsa Social/TMCA) y **no se parecen** al `dps.xlsx` que lee el tablero. La pestaña **Consolidar DPS** (solo visible en modo administrador) arma el dps correcto:

1. Entre por `/admin/` y abra la pestaña **Consolidar DPS**.
2. Arrastre el ZIP (o los Excel sueltos). Cada archivo se reconoce **por su contenido**, no por su nombre.
3. Revise el resultado por programa, los avisos (meses cerrados que cambian, comedores nuevos/fusionados, pagos sin raciones, etc.) y la tabla de cifras de INTEGRADO que cambian.
4. Escriba los datos que **no vienen en los archivos** (presupuestos y ejecución contable, columnas E, F, G, K, N) y presione **Recalcular**.
5. **Descargar dps.xlsx** (incluye la hoja «NOTAS DE CONSOLIDACIÓN») o **Usar en el tablero (borrador)** y luego **Publicar datos**.

Reglas: se parte del dps vigente; lo que no viene en el ZIP se conserva y se avisa (p. ej. Beca Media); si la lista de municipios de un programa cambia, ese programa se conserva y se avisa en lugar de insertar filas. Jóvenes Protagonistas solo trae el acumulado: el detalle enero–julio se conserva y el acumulado se agrega en las columnas AC:AF. El motor está en `consolidador.js`.
