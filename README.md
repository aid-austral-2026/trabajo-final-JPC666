# Control de Gestión: Rentabilidad por Unidades de Negocio

Trabajo final de **Análisis Inteligente de Datos** (AID Austral, 2026).

**Autor:** Calvino Joaquin Pablo

**Producto:** aplicación web Shiny (`app.R`).

## Descripción

La aplicación evalúa la rentabilidad comercial, la estructura de costos y el punto de equilibrio de un comercio con tres unidades de negocio (Rosario, Funes y Roldán), a partir de sus libros de IVA y sus planillas de gastos mensuales. Incluye además una comparación interanual (2025 vs. 2026) ajustada por inflación.

Pestañas de la aplicación:

1. **Tablero de Rentabilidad:** ventas, compras de mercadería y gastos operativos; evolución mensual de la rentabilidad por unidad de negocio; ventas por unidad; y top 5 de proveedores de mercadería.
2. **Punto de Equilibrio:** ventas vs. punto de equilibrio por unidad, y tabla de márgenes de contribución.
3. **Comparación Interanual (IPC):** ventas, compras y gastos de 2025 ajustados por IPC contra 2026 a valores históricos, con la variación real mes a mes.
4. **Estadísticas Descriptivas:** medidas descriptivas de las ventas, resumen de compras de mercadería y resumen de gastos agrupados.
5. **Metodología (Selección y Limpieza):** problema en estudio, fuentes y pasos de limpieza.

Los filtros del panel lateral (período mensual y unidad de negocio) permiten explorar los datos de forma interactiva.

## Etapas del análisis

| Etapa | Dónde se refleja |
|---|---|
| Selección | Pestaña "Metodología" (tarjeta a) y sección *Fuente de los datos* de este README |
| Limpieza | Pestaña "Metodología" (tarjeta b) y sección 2 del código (`app.R`) |
| Exploración | Tablero, punto de equilibrio, comparación interanual y estadísticas descriptivas, con filtros dinámicos |
| Comunicación | Aplicación Shiny con tema `bslib` |

## Fuente de los datos

- **IVA Ventas e IVA Compras:** libros contables mensuales de IVA (AFIP), exportados a Excel desde el ERP del comercio.
- **Planillas de gastos operativos:** planillas mensuales de gastos asignados por punto de venta, generados con carga manual por el comercio.
- **IPC:** índice de precios al consumidor, nivel general nacional, publicado por el INDEC.
- **Período cubierto:** 2025 y 2026 [desde 01/2025 hasta 10/2026]. Los datos de 2026 alimentan el tablero, el punto de equilibrio y las estadísticas descriptivas; los de 2025 se usan solo en la comparación interanual.

### Nota sobre los datos

[completar según la decisión que tomes: por ejemplo, "los datos fueron anonimizados (clientes, proveedores y CUIT)" o "el repositorio es privado a pedido de los docentes porque los datos son reales".]

## Estructura del repositorio

```
.
├── app.R              # Aplicación Shiny (código completo)
├── README.md
└── datos/
    ├── ventas/        # Libros de IVA Ventas mensuales (.xls / .xlsx)
    ├── compras/       # Libros de IVA Compras mensuales (.xls / .xlsx)
    ├── gastos/        # Planillas mensuales de gastos (.xls / .xlsx)
    └── ipc/           # IPC nivel general nacional, INDEC (.xlsx)
```

La aplicación lee automáticamente todos los archivos de cada carpeta, de modo que para sumar un mes nuevo alcanza con copiar el archivo en la carpeta que corresponda.

## Cómo ejecutarla

Los paquetes necesarios son `shiny`, `bslib`, `tidyverse` (incluye `lubridate` y `purrr`), `scales` y `readxl`. Se recomienda `tidyverse` 2.0 o superior y `purrr` 1.0 o superior.

```r
install.packages(c("shiny", "bslib", "tidyverse", "scales", "readxl"))
```

Desde RStudio, con la raíz del repositorio como *working directory*:

```r
shiny::runApp("app.R")
```

Directamente desde GitHub:

```r
shiny::runGitHub("trabajo-final-JPC666", "aid-austral-2026")
```

## Criterios y supuestos

- **Unidades de negocio:** se asignan según el punto de venta (emisor): 3 y 6 corresponden a Rosario, 
    5 y 7 a Funes y 8 a Roldán.
- **Rentabilidad:** ventas netas gravadas menos compras de mercadería menos gastos operativos. 
    Solo se consideran las compras del rubro *Mercaderias*; el rubro *Varios* no se resta porque ya está 
    cubierto por la planilla de gastos.
- **Compras por unidad (estimación):** el libro de IVA Compras no informa la unidad de negocio, por lo que 
    las compras de cada mes se reparten según la participación de cada unidad en las ventas de ese mes. 
    Como consecuencia, todas las unidades tienen el mismo margen bruto, y las diferencias de rentabilidad 
    entre ellas provienen del volumen de ventas y de los gastos directos.
- **Punto de equilibrio:** gastos fijos de la unidad divididos por el margen de contribución global 
    (porcentaje de las ventas que no se destina a mercadería).
- **Comparación interanual:** cada mes de 2025 se multiplica por el coeficiente IPC 
    (IPC del mes en 2026 / IPC del mismo mes en 2025) y se compara contra el valor histórico de 2026. 
    Solo se comparan los meses con datos en ambos años.
- **Estadísticas descriptivas:** se excluyen las notas de crédito y las facturas asociadas 
    (emparejadas por cliente o proveedor e importe), para que no distorsionen los valores mínimo y máximo.
- **Agrupación de gastos:** en las estadísticas descriptivas, los conceptos de la planilla de gastos se 
    agrupan en *Servicios* (EPE, telefonía, internet, alarmas, limpieza, Sinergia, refrigerios y varios), 
    *SUSS* (F931, AEC, FAECYS y Estrella) e *Impuestos* (DREI, IIBB, 
    Impuesto al Valor Agregado, IIGG y autónomos). 
    Los conceptos que no pertenecen a ningún grupo se muestran con su nombre original.



