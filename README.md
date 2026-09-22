# TFA-OEFA Analizador v4

Herramienta desarrollada en Python para automatizar la búsqueda,
extracción y análisis documental de resoluciones del Tribunal de
Fiscalización Ambiental (TFA) del OEFA.

## Funcionalidades

La versión 4 permite:

-   Buscar un tema, varios temas, todos los temas configurados o una
    palabra/frase directa.
-   Obtener las resoluciones y sus PDF desde la colección pública del
    TFA-OEFA.
-   Procesar los PDF directamente en memoria mediante `BytesIO`, sin
    guardarlos localmente.
-   Identificar páginas con coincidencias y extraer fragmentos de
    contexto.
-   Identificar automáticamente el administrado de la resolución cuando
    el patrón documental lo permite.
-   Generar estadísticas por tema, resolución y administrado.
-   Analizar la coocurrencia de temas.
-   Generar gráficos.
-   Exportar los resultados a Excel.

## Fuente

Colección pública del Tribunal de Fiscalización Ambiental de OEFA:

https://www.gob.pe/institucion/oefa/colecciones/1716-sala-especializada-en-mineria-energia-actividades-productivas-e-infraestructura-y-servicios-tfa-se

## Procesamiento de documentos

El flujo general es:

1.  Obtener la resolución desde la colección del TFA-OEFA.
2.  Identificar la URL de su PDF.
3.  Cargar temporalmente el PDF en memoria.
4.  Extraer el texto de sus páginas.
5.  Buscar los términos configurados.
6.  Registrar coincidencias, páginas y fragmentos.
7.  Identificar el administrado cuando es posible.
8.  Generar los resultados estructurados.

Los PDF no necesitan almacenarse en el disco local para realizar el
análisis.

## Temas

La versión 4 contempla categorías relacionadas con componentes
ambientales, aspectos ambientales, gestión ambiental y fiscalización,
entre ellas:

-   Ruido
-   Aire
-   Agua
-   Suelo
-   Sedimentos
-   Vibraciones
-   Olores
-   Flora
-   Fauna
-   Biodiversidad
-   Emisiones
-   Efluentes
-   Relaves
-   Residuos
-   Hidrocarburos
-   Monitoreo ambiental
-   Instrumento de gestión ambiental
-   ECA
-   LMP
-   Obligaciones ambientales
-   Fiscalización
-   Incumplimiento
-   Infracción
-   Sanción
-   Medidas administrativas
-   Beneficio ilícito
-   Daño ambiental

Los términos asociados se encuentran definidos en el código y pueden
ampliarse o modificarse.

## Administrado

La V4 incorpora la identificación automática del administrado al que
corresponde cada resolución.

La salida consolida:

-   administrado;
-   número de resoluciones;
-   número de coincidencias;
-   años asociados.

Los casos que no pueden identificarse automáticamente quedan registrados
como **No identificado** para revisión.

Esta identificación es una extracción documental automática y no
constituye una validación jurídica independiente.

## Resultados Excel

El archivo generado contiene hojas orientadas a diferentes niveles de
análisis:

### Resumen

Indicadores generales del procesamiento y estadísticas por tema.

### Resoluciones

Una fila por resolución, con URL, PDF, administrado, páginas,
coincidencias y resultados por tema.

### Detalle

Registro de cada coincidencia con resolución, administrado, tema,
término, página, sección aproximada y fragmento.

### Administrados

Consolidación por administrado, con resoluciones, coincidencias y años.

### Coocurrencia

Pares de temas que aparecen conjuntamente en las mismas resoluciones.

### Gráficos

Visualizaciones automáticas de los resultados por tema.

## Interpretación

Las coincidencias son ocurrencias textuales de los términos
configurados. Una mayor cantidad de coincidencias no implica
necesariamente mayor importancia jurídica, ambiental o técnica del tema.

Un término puede aparecer en antecedentes, fundamentos, citas
normativas, referencias u otras secciones. Por ello, cuando se requiera
una interpretación sustantiva, debe revisarse el contenido de la
resolución.

## Requisitos

-   Python 3.10 o superior.
-   Conexión a Internet.
-   Dependencias indicadas en `requirements.txt`.

## Instalación

``` bash
pip install -r requirements.txt
```

## Ejecución

``` bash
python buscar_resoluciones_v4_final.py
```

Ejemplos de búsqueda:

``` text
ruido
```

``` text
ruido, agua, emisiones
```

``` text
todos
```

## Consideraciones

La herramienta depende de la disponibilidad y estructura de las
publicaciones del sitio web de OEFA y de la capacidad de extracción de
texto de los PDF.

Cambios en la estructura web, enlaces o formato de los documentos pueden
requerir ajustes posteriores.

## Alcance

TFA-OEFA Analizador v4 es una herramienta de automatización y análisis
documental. Su objetivo es localizar, estructurar y cuantificar
información textual.

No realiza por sí sola una calificación jurídica, determinación de
responsabilidad ni interpretación definitiva de la importancia de una
materia.

## Tecnologías

-   Python
-   Requests
-   BeautifulSoup
-   PyPDF
-   OpenPyXL
-   BytesIO
-   Expresiones regulares (`re`)

## Autor

Proyecto desarrollado como iniciativa de automatización y análisis
documental aplicada a información ambiental y resoluciones del Tribunal
de Fiscalización Ambiental del OEFA.

**Versión: TFA-OEFA Analizador v4 --- versión base para publicación**
