# Vacaciones CX

**Página:** https://facundoferreyra-cell.github.io/vacaciones-cx/  
**Datos:** Google Sheet "Vacaciones CX" (Drive de facundo.ferreyra@coderhouse.com) → Extensiones → Apps Script.

Página para que el equipo de CX de Coderhouse postule sus fechas de vacaciones y coordinación las asigne por prioridad (antigüedad → métricas → orden de llegada), detectando solapamientos por turno.

- **Página:** GitHub Pages (`index.html`, sin dependencias).
- **Datos:** un Google Sheet, a través de un script de Google Apps Script (`Code.gs`) publicado como app web.
- **Acceso:** el equipo entra con una clave compartida; coordinación con otra clave que habilita las pestañas *Asignación* y *Equipo y criterios*.

## Puesta en marcha

### 1. Crear el Sheet y el script

1. Creá un Google Sheet nuevo (por ejemplo, "Vacaciones CX").
2. Menú **Extensiones → Apps Script**.
3. Borrá el contenido de `Código.gs` y pegá el de `Code.gs` de este repo.
4. Arriba del archivo cambiá las dos claves:
   ```js
   const TEAM_CODE  = 'cx2027';       // la que compartís con el equipo
   const ADMIN_CODE = 'coordinacion'; // la tuya
   ```
5. **Implementar → Nueva implementación → ⚙️ Tipo: App web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
   - Implementar → autorizar permisos → copiar la **URL de la app web** (termina en `/exec`).

### 2. Conectar la página

En `index.html`, pegá esa URL en la constante `API_URL` (está al principio del `<script>`):

```js
const API_URL = 'https://script.google.com/macros/s/XXXX/exec';
```

Subí el cambio al repo. GitHub Pages se actualiza solo en un minuto.

### 3. Cargar el equipo

Entrá a la página con la clave de coordinación → pestaña **Equipo y criterios**: período, regla, criterios y la tabla del equipo (nombre, turno, área, fecha de ingreso, métrica). Guardar.

Después compartí el link + la clave del equipo.

## Cómo funciona la asignación

1. Cada persona carga uno o más rangos (desde/hasta) o se marca como **"no tengo fechas definidas"**.
2. Los pedidos se ordenan por los criterios de desempate configurados.
3. Se aprueban de a uno: un rango se aprueba si no se pisa con otro ya aprobado dentro de su grupo (turno, turno + área o todo el equipo, según la regla).
4. Coordinación puede forzar cualquier fila (aprobar/rechazar), cargar pedidos a nombre de alguien o importar las respuestas de un Google Form pegándolas.
5. Nada llega al equipo hasta que se aprieta **Publicar asignación**.

## Actualizar el script

Cada vez que cambies `Code.gs` en Apps Script: **Implementar → Administrar implementaciones → ✏️ → Versión: Nueva → Implementar**. La URL no cambia.

## Datos en el Sheet

| Pestaña | Contenido |
|---|---|
| `Config` | título, período, regla, criterios, abierto/cerrado, fecha de última publicación |
| `Equipo` | id, nombre, turno (`am`/`pm`), área, activo, fecha de ingreso, métrica |
| `Pedidos` | un rango por fila (o una fila "sin fechas" con `sin_fechas = TRUE`) |
| `Forzados` | aprobaciones/rechazos manuales |
| `Resultados` | la última asignación publicada |

Podés corregir cualquier cosa directamente en el Sheet; la página lo lee al recargar.
