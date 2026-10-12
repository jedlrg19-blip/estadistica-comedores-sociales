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

## Actualizar datos (administrador)
1. Reemplace el Excel en `data/` **con el mismo nombre** (commit en GitHub → Cloudflare Pages publica solo).
2. `data/` tiene prioridad sobre `datos.json`; este solo se usa para lo que falte en `data/`.
3. Revise la pestaña *Fuentes* → *Conciliación*: ahí aparecen diferencias entre archivos.
4. Responsables/correos: edite `data/fuentes.json` (el botón «Solicitar actualización» abre un correo prellenado).
5. Meta anual oficial: `data/config.json` con `{"projRac": 9219150, "projFin": 0}` (ver advertencia en Conciliación).

## Modo administrador
`https://<sitio>/admin/` (o `/?editor=1`) habilita: descarga del Excel original, edición de proyección, publicar datos.

## Seguridad real (Cloudflare Zero Trust → Access)
`?editor=1` **no es seguridad**, solo oculta controles. Además, todo lo que está en `data/` es público por URL.
Para que solo el admin vea/descargue los originales:
1. Cloudflare Zero Trust → Access → Applications → *Self-hosted*.
2. Dominio del sitio, rutas `/admin*` y `/data/*`, política *Allow* solo para correos del admin.
3. Si `/data/*` se protege, el sitio público dejará de leer los Excel: publique entonces un paquete `datos.json` (botón «Publicar datos») para el público y deje los originales solo bajo Access.
