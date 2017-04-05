# MachiMatch
Software to rename file

Author:Hechelion (hechelion@gmail.com)www.deitec.cl 
programing language: Visual Basic 6 - SP6


Para la comunidad de http://arcadespain.info/
Agradecimientos:-Machiminax-Gucaza-Getterrobot-Pevalle-Empardopo-OnoferPor probar el programa pese a todos los errores. Y por sus sugerencias sobre como mejorarlo.

Cambios:1.3.0 (05-04-17)
-- Agregado el cambio de "&apos;" por "'" cuando se leen nombres desde un XML
-- Agregado el cambio de "&amp;" por "&" cuando se leen nombres desde un XML
-- Agreado sistema para generar lista de reemplazo personalidazas cuando se leen nombres desde un XML
----- Agregar la etiqueta "[XML]" al final del archivo "config.ini"
----- Bajo el tag [XML] agregar "serach_N=<palabra a ser reemplazada>"
----- agregar "replace_N=<palabra nueva>"
----- solo en replace pueden usar la palabra reservada "ascii_M" donde M es el valor DECIMAL del caracter ascii
----- N es un indice que debe comenzar por 0 e ir aumentado de forma secuencial por cada palabra que deseen reemplazar 

1.2.0 (02-04-17)
-- Correjido un error en el algoritmo que llena los resultados cuando se usan filtros y se cambiaba el valor del scroll vertical
-- Cambiado el HMI para el sistema de filtros
-- Agregado un nuevo filtro. "Marcados para usar" que muestra solo los valores que están marcados para ser usados
-- Implementado sistema para ordenar los datos
-- Implementado sistema para buscar "nombres" o "ROM"1.1.4 (25-02-17)
-- Agregado el check "Los nombres tienen extensión" que fija manualmente si los nombres tienen o no tiene extensión cuando se obtiene desde un archivo de texto o desde un archivo XML.

1.1.3 (21-02-17)
-- Selección automática de idioma por defecto según sea el idioma del OS.
-- Al cambiar manualmente el idioma del programa, este pasa a ser el idioma por defecto para cada vez que se vuelva a lanzar el programa.

1.1.2 (21-02-17)
-- Corregido un error al compilar la versión 1.1.1 que daba problemas al mover los archivos.1.1.1 (17-01-17)
-- Agregada la opción de listar tdos los archivos de un directorio mediante el filtro ".*"
-- Cambiado el código que mueve las snap, ahora el programa determina si debe copiar o mover los archivos, optimizando el tiempo que tomaba mover archivos grandes.

1.1.0 (27-12-16)
-- Modificadas todas las variables que manejaban la memoria de listas de rom y snap para superar el limite de 32.000 archivos. Actualmente el programa puede manejar 2.000.000.000 de Snap y 640.000 rom
-- El programa ahora recuerda la última configuración completa usada.
-- Cambiados los nombres de la tabla resultados: Nombre -> Nombre original, Archivo -> Archivo a renombrar.
-- Adaptado el formulario para resoluciones de 800*600.
-- Centrado automático del programa cada vez que se inicia.
-- Agregada una opción de aceleración de búsqueda para "comprobación exacta"

1.0.0 (25-11-16)
-- Modificado sistema que guarda la última busqueda realizada, ahora se guarda los campos tipo, nodo y propiedad
-- Modificada completamente la tabla de resultados, ahora se adapta de forma dinámica al tamaño del formulario
-- Simplificado el sistema de errores y uso de SNAP repetidas          
----Rojo intenso: ROM sin SNAP.		  
----Rojo débil: ROM con SNAP repetida que no será usada		  
----Amarillo débil: SNAP repetida pero marcada para ser procesada al renombrar		  
----Naranaja: Sin error. Pero NO marcado para ser procesado al renombrar
-- Modificado el Check "Permitir SNAP repetidas al renombrar", ahora, si el check no está marcado, el programa no permite marcar la ROM como "usar" si tiene un archivo o SNAP repetido
-- Agregado un nuevo filtro "Sin usar", que solo muestra los nombres que están desmarcados para ser procesados al renombrar
-- Agregada una ventana desplegable al hacer doble clic sobre la columna "Nombres" que permite asociar cualquier archivo a cualquier nombre independiente del porcentaje de coincidencia.
-- Agregado un selector de algoritmo de comparación.
-- Implementado el idioma inglés dentro del programa.
-- Implementado el algoritmo de comparación "Distancia de Levenshtein"
-- Elimniado el check "match exacto". 
