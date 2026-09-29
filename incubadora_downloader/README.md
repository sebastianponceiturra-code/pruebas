# Descarga diaria de resultados de nacimiento (incubación)

Automatiza la consulta manual de `http://192.168.9.218:50000/NewIncubadora` y
deja todo consolidado en un único Excel acumulado (`nacimientos_acumulado.xlsx`).

**Importante:** este script debe correr en una máquina con acceso a la red
`192.168.9.x` (tu PC, o un servidor de esa red). No puede correr desde una
máquina en la nube sin VPN a esa subred.

## Instalación (una sola vez, no requiere permisos de administrador)

Ya validamos que tu PC tiene Python 3.13 (instalado desde Microsoft Store) y
Playwright con el navegador Chromium instalados. Falta solo `openpyxl`:

```powershell
python -m pip install --user openpyxl
```

(Si en algún momento partes de cero en otra máquina: instala Python desde la
Microsoft Store escribiendo `python` en PowerShell, luego
`python -m pip install --user -r requirements.txt` y
`python -m playwright install chromium`.)

## Uso manual

Desde la carpeta `incubadora_downloader`:

```powershell
# Descarga el día de ayer (uso normal para la tarea diaria)
python descargar_resultados.py

# Descarga un rango de fechas puntual (para completar historial atrasado)
python descargar_resultados.py --desde 2026-09-01 --hasta 2026-09-28

# Para ver el navegador mientras corre (debug)
python descargar_resultados.py --visible
```

Los resultados se van agregando a `nacimientos_acumulado.xlsx` en esta misma
carpeta. Si un día ya está en ese archivo, se omite automáticamente al
volver a correr el script (no genera filas duplicadas).

## Programarlo para que corra solo, todos los días

Usa el Programador de Tareas de Windows (Task Scheduler) — no requiere
permisos de administrador para crear una tarea que corre bajo tu propio
usuario:

1. Abre "Programador de tareas" (Task Scheduler) desde el menú inicio.
2. "Crear tarea básica..." → nombre: `Descarga resultados incubación`.
3. Desencadenador: Diariamente, a la hora que prefieras (ej. 07:00, después
   de que el sistema termine de cargar los nacimientos del día anterior).
4. Acción: "Iniciar un programa".
   - Programa/script: la ruta completa a tu `python.exe`, por ejemplo:
     `C:\Users\sponce\AppData\Local\Microsoft\WindowsApps\python.exe`
   - Argumentos: `descargar_resultados.py`
   - Iniciar en: la ruta completa a esta carpeta `incubadora_downloader`.
5. Finalizar.

## Si el sistema cambia y el script deja de funcionar

Esta página está hecha en GeneXus, así que los parámetros de cada consulta
se generan por JavaScript en el momento y no se pueden armar a mano. Si el
sistema se actualiza y el script empieza a fallar (por ejemplo, ya no
encuentra un botón o un campo), la forma más rápida de volver a captar los
selectores correctos es grabar el flujo de nuevo con Playwright:

```powershell
python -m playwright codegen http://192.168.9.218:50000/NewIncubadora/servlet/com.incubadora.resultincub
```

Repite la consulta manual completa en la ventana que se abre, copia el
código generado y compártelo para actualizar el script.

## Qué hace exactamente el script

Por cada día del rango pedido:

1. Abre la página y configura: Año, Mes Desde/Hasta, Día Desde/Hasta (mismo
   día), Origen = Reproductoras, Sexo = Todos, Rango de Fechas, y confirma.
2. Por cada Sala que aparece con datos ese día, entra a "Sectores".
3. Agrupa automáticamente los pabellones por "Planta". Si una sala tiene más
   de una Planta (como pasa con Incubación Lo Miranda), genera un Excel por
   cada Planta por separado, porque mezclarlas da error en el sistema; si
   solo tiene una Planta, baja todo en un solo Excel.
4. Cada Excel descargado ya trae identificada la ubicación (Planta,
   SubPlanta, Sector, Desc.Sector); el script solo le agrega la columna
   "Fecha" y lo agrega al acumulado.

Solo se usa la zona general "Incubación" (no cada zona individual), porque ya
trae automáticamente todas las salas que tuvieron nacimientos ese día.
