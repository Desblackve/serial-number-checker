Creador: Desblack "Dess"
Contacto: desblackve@gmail.com


SERIAL NUMBER CHECKER
=====================

Versión 2.0


1. QUÉ HACE LA APLICACIÓN
-------------------------

Serial Number Checker tiene dos modos (botones arriba a la izquierda):

- CHECKER: carga una lista de números de serie (S/N) desde un archivo CSV
  o Excel y comprueba, escaneando o escribiendo, cuáles están, cuáles
  faltan y cuáles sobran. Al terminar exporta un reporte Excel coloreado.
- CREAR LISTA S/N: arma una lista nueva con lo que escaneas, organizada
  por pallets (en columnas o en hojas) o, con el Modo Rack, por
  ubicaciones de un rack. La exporta a Excel o CSV.

La interfaz está en español, inglés y portugués. El botón README (arriba
a la derecha) abre esta guía dentro de la aplicación.


2. EJEMPLOS RÁPIDOS
-------------------

EJEMPLO 1 - Revisar un inventario contra un archivo

  1. Pulsa "Cargar CSV / Excel" y elige el archivo (por ejemplo,
     inventario_octubre.xlsx).
  2. Escanea los equipos uno a uno. Cada lectura muestra:
       Verde      ENCONTRADO        (Hoja Lote1 · fila 7)
       Amarillo   YA ESCANEADO      (el contador no sube)
       Rojo       NO ENCONTRADO     (se guarda en Sobrantes)
  3. Mira el contador: "128 / 340" y "38% completado · 212 pendientes".
  4. Pulsa "Exportar reporte": se crea inventario_octubre_resultado.xlsx
     con los S/N en verde (escaneado), rojo (pendiente) y azul
     (duplicado), más las hojas Resumen, Escaneos y Sobrantes.

EJEMPLO 2 - Me equivoqué de código

  Si escaneas el código equivocado, pulsa Ctrl+Z (o el botón "Deshacer"
  de Últimos escaneos): se anula el último escaneo y el contador vuelve
  a como estaba.

EJEMPLO 3 - Tengo los S/N en un correo o en otra hoja de Excel

  Copia la lista (uno por línea) y pégala con Ctrl+V en el campo de
  escaneo. La aplicación procesa todos de una vez y muestra un resumen:
  "Pegados 25: 22 encontrados · 2 no encontrados · 1 repetidos".

EJEMPLO 4 - Armar una lista de pallets

  1. Pasa al modo "Crear lista S/N".
  2. Escanea los S/N del primer pallet: se guardan en la columna A
     (Pallet 1).
  3. Pulsa "Siguiente pallet" (o F2): los siguientes van a la columna B
     (Pallet 2), y así sucesivamente.
  4. Si quieres un pallet en una hoja aparte, pulsa "Nueva hoja" junto a
     las pestañas de hojas.
  5. Pulsa "Exportar lista" (o Ctrl+E) para guardarla en Excel o CSV.

EJEMPLO 5 - Llenar un rack ubicación por ubicación (Modo Rack)

  1. En "Crear lista S/N", marca la casilla "Modo Rack".
  2. Escribe la ubicación (por ejemplo, DC05), las filas y columnas del
     rack (por ejemplo, 6 y 8) y el orden: Horizontal o Vertical.
  3. Pulsa "Agregar nuevo rack": se crea DC05-1.
  4. Escanea: el contador muestra la siguiente ubicación, por ejemplo
     (0,0), y cada S/N se guarda ahí. Luego (1,0), (2,0)...
  5. Cuando se llena el rack, la aplicación avisa "Se completó el rack
     DC05-1". Pulsa "Agregar nuevo rack" para seguir con DC05-2.

EJEMPLO 6 - Ver el avance de lejos

  Pulsa F11 (o "Vista grande"): la pantalla muestra solo el contador y el
  resultado en tamaño enorme. F11 o Esc para volver.


3. NOVEDADES DE LA VERSIÓN 2.0
------------------------------

- Diseño nuevo: tarjetas oscuras sobre fondo de líneas, botones con
  iconos y barra de título oscura en Windows.
- Ejecutable para Windows (SerialNumberChecker.exe): ya no hace falta
  instalar Python.
- Botón README: esta guía se puede leer dentro de la aplicación.
- Modo Rack: lista guiada por ubicaciones (x,y) de cada rack.
- Sobrantes: los S/N que no están en el archivo se guardan en su propia
  pestaña y se exportan en una hoja aparte.
- Deshacer el último escaneo (Ctrl+Z o botón "Deshacer").
- Hoja y fila: al encontrar un S/N se indica dónde está en el archivo.
- Filtro de formato: si se escanea otro código de la caja (por ejemplo,
  uno de 13 caracteres), avisa "NO PARECE UN S/N" y no lo cuenta.
- Pegar varios S/N a la vez con Ctrl+V.
- Contador con barra de progreso, ritmo (S/N por minuto) y tiempo
  estimado para terminar.
- Pestañas Pendientes / Sobrantes con buscador.
- Vista grande (F11).
- El reporte incluye las hojas Resumen, Escaneos (hora de cada lectura)
  y Sobrantes.
- Imprimir (Ctrl+P): resumen, pendientes y sobrantes (Checker) o la hoja
  actual / todas las hojas (Crear lista).
- Crear lista: botón "Nueva hoja", columnas que se pueden estirar y
  tecla Retroceso para eliminar el registro seleccionado.
- Se reconocen más encabezados de S/N (ver sección 14).
- La sesión se guarda sola: al volver a abrir la aplicación sigues donde
  lo dejaste.


4. MODO CHECKER, PASO A PASO
----------------------------

  1. Abre Serial Number Checker.
  2. Si lo necesitas, cambia el idioma en el selector de arriba a la
     derecha (Español, English o Português). El cambio es inmediato.
  3. Pulsa "Cargar CSV / Excel" y selecciona el archivo con los S/N.
  4. Se muestra el nombre del archivo, los números leídos y los S/N
     únicos.
  5. Si el archivo es un reporte exportado antes, la aplicación pregunta
     si quieres continuar donde te quedaste (sección 9).
  6. Escanea o escribe un S/N en el campo de escaneo.
  7. Al llegar a 17 caracteres la comprobación es automática. También
     puedes pulsar ENTER.
  8. Repite con todos los S/N que quieras comprobar.
  9. Al terminar, pulsa "Exportar reporte".
  10. Para empezar de cero, usa el botón "Reiniciar" (sección 15).

La casilla "Alternativas A/B/U" de la barra superior activa o desactiva
la búsqueda alternativa de la sección 5.


5. REGLAS DE COINCIDENCIA
-------------------------

- Primero se busca una coincidencia exacta con un S/N cargado.
- Si no la hay y "Alternativas A/B/U" está marcada, se prueba sustituyendo
  el séptimo carácter por A, B y U. Si una alternativa coincide, se marca
  como encontrado el S/N que realmente aparece en la lista (el mensaje
  dice "variante" y qué se leyó).
- Un S/N ya encontrado no vuelve a sumar en el contador.
- Si el código leído no tiene el largo de ningún S/N del archivo, se
  avisa "NO PARECE UN S/N" (probablemente se escaneó otro código de la
  caja) y no se guarda como sobrante.

IMPORTANTE: la sustitución del séptimo carácter es una búsqueda
alternativa. No garantiza que el número de serie sea auténtico ni es una
validación completa del formato.


6. COLORES, SONIDOS Y MENSAJES
------------------------------

En la aplicación:

- VERDE: S/N encontrado. Sonido de confirmación. El mensaje indica si fue
  exacto o una variante, y en qué hoja y fila del archivo está.
- AMARILLO: S/N ya escaneado antes. Otro sonido. El contador no sube y la
  aplicación queda en amarillo hasta el siguiente escaneo o hasta pulsar
  Esc.
- ROJO: S/N no encontrado (se guarda en Sobrantes) o código que no
  parece un S/N. Sonido de error.
- El borde del campo de escaneo destella en el color del resultado.
- Bajo el campo de escaneo hay una leyenda con los colores del reporte.

En el reporte Excel:

- VERDE: S/N encontrado durante la revisión.
- ROJO: S/N que todavía no se ha encontrado.
- AZUL: S/N duplicado en el archivo. La primera aparición conserva su
  verde o rojo y las repeticiones (en cualquier hoja) salen en azul.

Si se encontró por una alternativa A/B/U, se colorea el S/N real de la
lista.


7. CONTADOR, PROGRESO Y RITMO
-----------------------------

El contador muestra los S/N encontrados respecto al total de S/N únicos:

    128 / 340
    38% completado · 212 pendientes
    12/min · faltan ~18 min

- Un S/N repetido en el archivo cuenta una sola vez.
- Al cargar un archivo se indica cuántos números se leyeron y cuántos son
  únicos. Ejemplo: "12 números leídos · 10 S/N únicos" significa que hay
  2 valores repetidos en el archivo.
- El ritmo y el tiempo estimado aparecen tras unos escaneos y se calculan
  con los últimos 10 minutos.


8. PENDIENTES, SOBRANTES E HISTORIAL
------------------------------------

- ÚLTIMOS ESCANEOS: los 5 últimos S/N leídos y su resultado. El botón
  "Deshacer" (o Ctrl+Z) anula el último escaneo; se puede repetir.
- PENDIENTES: los S/N del archivo que faltan por escanear.
- SOBRANTES: los S/N escaneados que no están en el archivo, con cuántas
  veces se leyeron y la hora de la primera lectura.
- BUSCAR: filtra la pestaña abierta mientras escribes. Enter o Esc
  devuelven el cursor al campo de escaneo.


9. CONTINUAR UN REPORTE EXPORTADO
---------------------------------

Si cargas un reporte XLSX exportado antes por la aplicación, puedes
seguir donde lo dejaste.

- La aplicación detecta los S/N en verde del reporte y pregunta en el
  aviso amarillo si quieres continuar.
- "Continuar": esos S/N se marcan como encontrados. "Empezar de cero": el
  archivo se carga sin progreso.
- Las hojas Resumen, Escaneos y Sobrantes del reporte se ignoran al leer
  y se reemplazan al volver a exportar.
- Los CSV no se pueden continuar porque no guardan colores.

NOTA: se reconoce el verde exacto del reporte. Si un archivo tiene celdas
con ese mismo verde por otra razón (por ejemplo, el estilo "Bueno" de
Excel), también preguntará; en ese caso elige "Empezar de cero".


10. EXPORTAR REPORTE
--------------------

"Exportar reporte" (o Ctrl+E) guarda un Excel (.xlsx) con el resultado.

- Si el archivo cargado es XLSX, el reporte conserva sus hojas.
- Si es CSV, se crea un XLSX para poder colorear las celdas.
- Las columnas de S/N se colorean en verde, rojo o azul (sección 6).
- Se añaden tres hojas:
    Resumen      archivo, totales, encontrados, pendientes, sobrantes y
                 horas de la primera y última lectura.
    Escaneos     cada lectura con su hora, lo leído y el resultado.
    Sobrantes    los S/N que no están en el archivo (si hay).
- El archivo original no se modifica: el reporte se guarda aparte y no
  se permite guardarlo encima del original.
- Excepción: si estás continuando un reporte (sección 9), puedes guardar
  el avance sobre ese mismo reporte. Windows pedirá confirmar que deseas
  reemplazarlo.
- El nombre sugerido es nombre_del_archivo_resultado.xlsx.


11. MODO CREAR LISTA S/N
------------------------

Sirve para armar una lista nueva con lo que escaneas. Acepta cualquier
texto (sin cambiar mayúsculas ni limitar el largo).

- DESTINO del siguiente pallet:
    Columna nueva      cada pallet en la columna siguiente (por defecto).
    Hoja nueva         cada pallet en su propia hoja.
    Desde la celda     empieza en la celda seleccionada del visor.
- NOMBRE (opcional): título del pallet o de la hoja. Si se deja vacío se
  usa "Pallet 1", "Pallet 2"...
- "Siguiente pallet" (F2): empieza el pallet siguiente.
- "Nueva hoja" (junto a las pestañas de hojas): crea una hoja nueva y
  salta a ella.
- VISOR: muestra la hoja tal como se exportará. Clic para seleccionar una
  celda, doble clic para editarla. Las columnas se estiran arrastrando su
  borde en la fila de letras; doble clic en el borde las ajusta al texto.
- DUPLICADOS: entran en azul y la aplicación queda en amarillo con un
  aviso: "Eliminar duplicados" u "Ocultar aviso".
- ELIMINAR: Retroceso o Supr con el campo vacío, o el botón "Eliminar",
  borra el registro seleccionado. "Deshacer" (Ctrl+Z) lo recupera.
- EXPORTAR LISTA (Ctrl+E): a Excel (con las mismas hojas y colores del
  visor) o a CSV (S/N, pallet, hoja, celda y si es duplicado).
- Un valor que empiece con "=" se guarda como texto, no como fórmula.


12. MODO RACK
-------------

Para guardar equipos en racks, ubicación por ubicación. Se activa con la
casilla "Modo Rack" en la barra superior del modo Crear lista.

- UBICACIÓN: el lugar físico, por ejemplo DC05. Los racks se llaman
  DC05-1, DC05-2, DC05-3...
- FILAS y COLUMNAS: tamaño del rack (por defecto 6 filas y 8 columnas).
- COORDENADAS (x,y): x es la columna e y la fila, empezando en (0,0)
  arriba a la izquierda. El visor numera desde 0.
- ORDEN: Horizontal llena (0,0), (1,0), (2,0)... fila por fila. Vertical
  llena (0,0), (0,1), (0,2)... columna por columna.
- "AGREGAR NUEVO RACK" (o F2, o junto a las pestañas): crea el rack
  siguiente de la ubicación y empieza en (0,0).
- GUÍA: el contador muestra en grande la siguiente ubicación y el visor
  la marca con un recuadro verde. Cada escaneo se guarda ahí y avanza.
- SALTAR: haz clic en una ubicación vacía del visor y el siguiente
  escaneo irá ahí.
- RACK COMPLETO: al llenar la última ubicación la aplicación avisa "Se
  completó el rack DC05-1" y no registra más en ese rack hasta que
  agregues uno nuevo.
- EXPORTAR: una hoja por rack dibujada como la estantería (coordenadas x
  arriba e y a la izquierda) y una hoja "Lista" con Ubicación, Rack,
  Posición (x,y), X, Y, S/N y si es duplicado.


13. IMPRIMIR Y VISTA GRANDE
---------------------------

IMPRIMIR (botón "Imprimir" o Ctrl+P): la aplicación pregunta en el aviso
amarillo qué imprimir.

- Checker: marca Resumen, Pendientes y/o Sobrantes y pulsa "Imprimir".
- Crear lista: "Hoja actual" o "Todas las hojas".
- Se imprime en la impresora predeterminada de Windows. Si el equipo
  tiene Excel se imprime con Excel; si no, con el Bloc de notas.

VISTA GRANDE (F11): solo el contador y el último resultado en tamaño
grande, para verlos de lejos. Se puede seguir escaneando. F11 o Esc para
volver.


14. ARCHIVOS COMPATIBLES Y ENCABEZADOS
--------------------------------------

Formatos:

- XLSX: archivos de Excel. Se leen todas las hojas del libro.
- CSV: una sola tabla. El separador (coma, punto y coma, tabulación o |)
  se detecta solo.

Encabezados reconocidos como columna de S/N (sin importar mayúsculas,
acentos, espacios ni signos):

    SN, S/N, NS, N/S
    SERIAL, SERIAL NO, SERIAL NR, SERIAL NBR, SERIAL NUM, SERIAL NUMBER
    SERIE, N SERIE, Nº DE SERIE, NO SERIE, NO. DE SERIE
    NUM SERIE, NUM. DE SERIE, NÚMERO SERIE, NÚMERO DE SERIE
    N SERIAL, NÚMERO SERIAL

Reglas de lectura:

- Si hay varios encabezados de S/N (en columnas u hojas distintas), se
  leen todos.
- Las columnas que no son de S/N no se usan.
- Las celdas vacías y los valores NONE o NAN se ignoran.
- Si una celda tiene una fórmula, se usa el valor calculado que guardó
  Excel.
- Los números se leen sin decimales (123456.0 se lee como 123456).


15. ATAJOS, REINICIO Y SESIÓN
-----------------------------

Atajos de teclado:

    Enter        comprobar / registrar
    Ctrl+Z       deshacer
    Ctrl+E       exportar
    Ctrl+P       imprimir
    Ctrl+V       pegar varios S/N
    F11          vista grande
    Esc          quitar el aviso / salir de la vista grande
    F2           siguiente pallet o nuevo rack (Crear lista)
    Retroceso    eliminar el registro seleccionado (Crear lista, con el
                 campo vacío). Supr hace lo mismo.

REINICIAR: el botón "Reiniciar" (abajo a la derecha) borra la sesión del
modo activo. Si hay datos, el primer clic muestra "¿Seguro?" y hay que
pulsar otra vez en 3 segundos. El archivo original no se borra de tu
equipo.

SESIÓN: el avance se guarda solo. Si cierras la aplicación, al abrirla
recuperas el archivo cargado, lo escaneado, los sobrantes y la lista o
los racks en creación. Aun así, exporta el reporte antes de reiniciar.


16. REQUISITOS E INSTALACIÓN
----------------------------

EJECUTABLE (recomendado):

- Descomprime la carpeta completa y abre SerialNumberChecker.exe. No lo
  separes de la carpeta _internal.
- No necesita Python ni otras instalaciones.
- Si Windows muestra "Windows protegió su PC", pulsa "Más información" y
  "Ejecutar de todas formas". Ocurre porque el programa no está firmado
  digitalmente.

DESDE EL CÓDIGO (opcional):

- Python 3 para Windows (incluye Tkinter) y la biblioteca openpyxl:

      pip install openpyxl

- Después ejecuta serial_number_checker.py.
- Para crear el ejecutable, usa build_exe.bat.

Los sonidos funcionan en Windows. El icono de la aplicación es
arthur.ico.


17. RECOMENDACIONES
-------------------

- Verifica que el archivo cargado sea el correcto.
- Comprueba que las columnas de S/N tengan un encabezado reconocido
  (sección 14).
- Recuerda que se leen todas las hojas de los archivos XLSX.
- Si un S/N no aparece como encontrado, revisa que esté bien escrito y
  que exista en una columna de S/N.
- Revisa la pestaña Sobrantes antes de exportar.
- Exporta el reporte antes de reiniciar.
- Conserva una copia del archivo original sin modificar.


18. CONTACTO
------------

Proyecto, novedades y descargas en GitHub:

    https://github.com/Desblackve/serial-number-checker

------------------------------------------------------------
