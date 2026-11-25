# ==============================================================================
# TRABAJO FINAL: ANÁLISIS INTELIGENTE DE DATOS (2026)
# PRODUCTO: Aplicación Web Shiny Multi-Período para Rentabilidad Comercial
# DATASET: IVA Compras, IVA Ventas y Gastos Operativos (.xlsx)
# INTEGRANTES / AUTOR: CALVINO JOAQUIN PABLO
# ==============================================================================

# 1. CARGA DE PAQUETES Y CONFIGURACIÓN -----------------------------------------
library(shiny)
library(bslib)
library(tidyverse)
library(lubridate)
library(scales)
library(readxl)

mi_tema <- bs_theme(
  version = 5,
  bootswatch = "flatly",
  primary = "#2c3e50",
  secondary = "#18bc9c"
)

# Un color fijo por unidad de negocio (se mantiene al filtrar)
colores_unidad <- c(
  "Rosario" = "#619CFF",
  "Funes"   = "#F8766D",
  "Roldán"  = "#00BA38",
  "Otros"   = "#95a5a6"
)

# 2. ETAPA: SELECCIÓN Y LIMPIEZA (PROCESAMIENTO EN LOTE) ------------------------

# Función auxiliar segura para convertir fechas mixtas de Excel
parsear_fecha_excel <- function(col_fecha) {
  col_texto <- as.character(col_fecha)
  num_fecha <- suppressWarnings(as.numeric(col_texto))
  
  fecha_texto <- coalesce(
    suppressWarnings(as.Date(ymd_hms(col_texto))),
    suppressWarnings(as.Date(ymd(col_texto))),
    suppressWarnings(as.Date(dmy_hms(col_texto))),
    suppressWarnings(as.Date(dmy(col_texto)))
  )
  
  if_else(
    !is.na(num_fecha),
    as.Date(num_fecha, origin = "1899-12-30"),
    fecha_texto
  )
}

# Formato de pesos argentinos; muestra "No calculable" cuando el valor es NA
fmt_pesos <- function(x) {
  if_else(
    is.na(x),
    "No calculable",
    dollar(x, accuracy = 1, prefix = "$", big.mark = ".", decimal.mark = ",")
  )
}

# Etiquetas de ejes en miles de pesos (por ejemplo, $12.500 equivale a $12.500.000)
etiqueta_miles <- label_number(
  scale = 1e-3, prefix = "$",
  accuracy = 1, big.mark = ".", decimal.mark = ","
)

# Detecta notas de crédito por el texto del tipo de comprobante
es_nota_credito <- function(tipo) {
  str_detect(tipo, regex("cr[eé]dito|^NC\\b|\\bNC\\b", ignore_case = TRUE))
}

# Excluye las NC y, por cada una, una factura del mismo cliente/proveedor
# con el mismo monto (emparejamiento 1 a 1, se descarta la más antigua)
excluir_nc_y_facturas <- function(df, col_id) {
  if (nrow(df) == 0) return(df)
  
  df <- df %>%
    mutate(
      .es_nc = es_nota_credito(tipo) %in% TRUE,
      .key = .data[[col_id]],
      .monto_abs = round(abs(neto_gravado), 2)
    )
  
  nc_por_grupo <- df %>%
    filter(.es_nc) %>%
    count(.key, .monto_abs, name = ".n_nc")
  
  df %>%
    filter(!.es_nc) %>%
    arrange(fecha) %>%
    group_by(.key, .monto_abs) %>%
    mutate(.orden = row_number()) %>%
    ungroup() %>%
    left_join(nc_por_grupo, by = c(".key", ".monto_abs")) %>%
    mutate(.n_nc = replace_na(.n_nc, 0L)) %>%
    filter(.orden > .n_nc) %>%
    select(-starts_with("."))
}

# Rubros de compras que se consideran costo de mercadería vendida (CMV).
# En los libros de compras el rubro viene sin tilde ("Mercaderias").
# El rubro "Varios" no se suma como gasto: ya está cubierto por la planilla de gastos.
rubros_mercaderia <- c("Mercaderias")

# Tablas vacías con la estructura esperada. Se usan cuando no se detecta ningún
# archivo válido, para que el resto del código no falle por columnas inexistentes
# y la aplicación muestre el aviso de carga en lugar de cortarse.
vacio_ventas <- tibble(
  archivo = character(), fecha = as.Date(character()), periodo = character(),
  tipo = character(), punto_venta = numeric(), comprobante = character(),
  cliente = character(), categoria = character(), cuit = character(),
  neto_no_gravado = numeric(), neto_gravado = numeric(), iva_21 = numeric(),
  total = numeric(), unidad_negocio = character()
)

vacio_compras <- tibble(
  archivo = character(), fecha = as.Date(character()), periodo = character(),
  tipo = character(), comprobante = character(), provincia = character(),
  proveedor = character(), rubro = character(), neto_gravado = numeric(),
  iva_21 = numeric(), perc_iibb = numeric(), perc_iva = numeric(),
  total = numeric()
)

# Función para Limpiar Ventas
procesar_archivo_ventas <- function(ruta_archivo) {
  tryCatch({
    read_excel(ruta_archivo, skip = 3, col_types = "text", .name_repair = "unique_quiet") %>%
      filter(!is.na(Fecha), !is.na(Numero), Fecha != "Totales:") %>%
      transmute(
        archivo = basename(ruta_archivo),
        fecha = parsear_fecha_excel(Fecha),
        periodo = format(fecha, "%Y-%m"),
        tipo = Tipo,
        punto_venta = suppressWarnings(as.numeric(Emisor)),
        comprobante = Numero,
        cliente = Cliente,
        categoria = Categoria,
        cuit = Cuit,
        neto_no_gravado = suppressWarnings(as.numeric(`Neto No Gravado`)),
        neto_gravado = suppressWarnings(as.numeric(`Neto Gravado`)),
        iva_21 = suppressWarnings(as.numeric(`IVA 21.00%`)),
        total = suppressWarnings(as.numeric(Total)),
        unidad_negocio = case_when(
          punto_venta %in% c(3, 6) ~ "Rosario",
          punto_venta %in% c(5, 7) ~ "Funes",
          punto_venta == 8 ~ "Roldán",
          TRUE ~ "Otros"
        )
      ) %>%
      filter(!is.na(fecha))
  }, error = function(e) {
    message(paste("ERROR EN VENTAS ->", ruta_archivo, ":", conditionMessage(e)))
    vacio_ventas
  })
}

# Función para Limpiar Compras
procesar_archivo_compras <- function(ruta_archivo) {
  tryCatch({
    df_raw <- read_excel(ruta_archivo, skip = 3, col_types = "text", .name_repair = "unique_quiet") %>%
      filter(!is.na(Fecha), !is.na(Comprobante), Fecha != "Totales:")
    
    if (!"Perc. de IIBB" %in% names(df_raw)) df_raw$`Perc. de IIBB` <- "0"
    if (!"Perc. de IVA" %in% names(df_raw)) df_raw$`Perc. de IVA` <- "0"
    
    df_raw %>%
      transmute(
        archivo = basename(ruta_archivo),
        fecha = parsear_fecha_excel(Fecha),
        periodo = format(fecha, "%Y-%m"),
        tipo = Tipo,
        comprobante = Comprobante,
        provincia = NombreProvincia,
        proveedor = Proveedor,
        rubro = Rubro,
        neto_gravado = suppressWarnings(as.numeric(`Neto Gravado`)),
        iva_21 = suppressWarnings(as.numeric(`IVA 21.00%`)),
        perc_iibb = suppressWarnings(as.numeric(`Perc. de IIBB`)),
        perc_iva = suppressWarnings(as.numeric(`Perc. de IVA`)),
        total = suppressWarnings(as.numeric(Total))
      ) %>%
      filter(!is.na(fecha))
  }, error = function(e) {
    message(paste("ERROR EN COMPRAS ->", ruta_archivo, ":", conditionMessage(e)))
    vacio_compras
  })
}

# --- Función para Limpiar Planilla de Gastos (.xls / .xlsx) ---
procesar_archivo_gastos <- function(ruta_archivo) {
  tryCatch({
    df_raw <- read_excel(ruta_archivo, col_names = FALSE, col_types = "text", .name_repair = "unique_quiet") %>%
      select(1:5) %>%
      set_names(c("col_periodo", "concepto", "3", "5", "8"))
    
    # Detectar el período desde la celda 46023 o desde el nombre del archivo (_012026)
    fechas_detectadas <- parsear_fecha_excel(df_raw$col_periodo)
    fecha_valida <- fechas_detectadas[!is.na(fechas_detectadas)]
    
    if (length(fecha_valida) > 0) {
      periodo_archivo <- format(fecha_valida[1], "%Y-%m")
    } else {
      mmyyyy <- str_extract(basename(ruta_archivo), "\\d{6}")
      periodo_archivo <- paste0(substr(mmyyyy, 3, 6), "-", substr(mmyyyy, 1, 2))
    }
    
    df_raw %>%
      mutate(periodo = periodo_archivo, archivo = basename(ruta_archivo)) %>%
      filter(
        !is.na(concepto),
        !toupper(trimws(concepto)) %in% c("CONCEPTOS", "CONCEPTO", "TOTALES", "TOTAL", "IVA")
      ) %>%
      select(archivo, periodo, concepto, `3`, `5`, `8`) %>%
      pivot_longer(
        cols = c(`3`, `5`, `8`),
        names_to = "punto_venta",
        values_to = "monto"
      ) %>%
      mutate(
        concepto = trimws(concepto),
        punto_venta = suppressWarnings(as.numeric(punto_venta)),
        monto = suppressWarnings(as.numeric(monto)),
        unidad_negocio = case_when(
          punto_venta %in% c(3, 6) ~ "Rosario",
          punto_venta %in% c(5, 7) ~ "Funes",
          punto_venta == 8 ~ "Roldán",
          TRUE ~ "Otros"
        )
      ) %>%
      filter(!is.na(monto), monto != 0)
  }, error = function(e) {
    message(paste("ERROR EN GASTOS ->", ruta_archivo, ":", conditionMessage(e)))
    tibble(
      periodo = character(),
      concepto = character(),
      punto_venta = numeric(),
      monto = numeric(),
      unidad_negocio = character()
    )
  })
}

# Ingesta desde carpetas relativas (importante: \\.xlsx?$ acepta .xls y .xlsx)
archivos_ventas <- list.files(path = "datos/ventas", pattern = "^[^~].*\\.xlsx?$", full.names = TRUE, ignore.case = TRUE)
archivos_compras <- list.files(path = "datos/compras", pattern = "^[^~].*\\.xlsx?$", full.names = TRUE, ignore.case = TRUE)
archivos_gastos <- list.files(path = "datos/gastos", pattern = "^[^~].*\\.xlsx?$", full.names = TRUE, ignore.case = TRUE)

# --- Control de archivos duplicados ----------------------------------------------
# Un archivo es duplicado si tiene exactamente el mismo contenido (misma huella
# MD5) que otro de la misma carpeta. Se conserva el primero (orden alfabético) y
# los demás se excluyen para no contar dos veces los mismos datos.
separar_duplicados <- function(archivos, carpeta) {
  vacio <- tibble(carpeta = character(), archivo = character(), igual_a = character())
  if (length(archivos) == 0) return(list(unicos = archivos, excluidos = vacio))
  
  huella <- unname(tools::md5sum(archivos))
  es_duplicado <- duplicated(huella)
  
  excluidos <- tibble(
    carpeta = carpeta,
    archivo = basename(archivos[es_duplicado]),
    igual_a = basename(archivos[match(huella[es_duplicado], huella)])
  )
  list(unicos = archivos[!es_duplicado], excluidos = excluidos)
}

control_ventas <- separar_duplicados(archivos_ventas, "ventas")
control_compras <- separar_duplicados(archivos_compras, "compras")
control_gastos <- separar_duplicados(archivos_gastos, "gastos")

archivos_ventas <- control_ventas$unicos
archivos_compras <- control_compras$unicos
archivos_gastos <- control_gastos$unicos

archivos_excluidos <- bind_rows(
  control_ventas$excluidos,
  control_compras$excluidos,
  control_gastos$excluidos
)
if (nrow(archivos_excluidos) > 0) {
  message(
    "ARCHIVOS DUPLICADOS EXCLUIDOS -> ",
    paste0(archivos_excluidos$carpeta, "/", archivos_excluidos$archivo, collapse = ", ")
  )
}

ventas_todas <- bind_rows(vacio_ventas, map(archivos_ventas, procesar_archivo_ventas) %>% list_rbind())
compras_todas <- bind_rows(vacio_compras, map(archivos_compras, procesar_archivo_compras) %>% list_rbind())
gastos_todos <- map(archivos_gastos, procesar_archivo_gastos) %>% list_rbind()

# Dos archivos con contenido distinto pueden corresponder al mismo mes (por ejemplo,
# una versión corregida). No se excluyen, pero se avisa para que se revise.
# Se toma el período con más registros de cada archivo como su período principal.
periodos_repetidos <- function(df, carpeta) {
  vacio <- tibble(carpeta = character(), periodo = character(), archivos = character())
  if (nrow(df) == 0 || !"archivo" %in% names(df)) return(vacio)
  
  df %>%
    count(archivo, periodo) %>%
    group_by(archivo) %>%
    slice_max(n, n = 1, with_ties = FALSE) %>%
    ungroup() %>%
    group_by(periodo) %>%
    filter(n() > 1) %>%
    summarise(archivos = paste(archivo, collapse = " / "), .groups = "drop") %>%
    mutate(carpeta = carpeta) %>%
    select(carpeta, periodo, archivos)
}

periodos_duplicados <- bind_rows(
  periodos_repetidos(ventas_todas, "ventas"),
  periodos_repetidos(compras_todas, "compras"),
  periodos_repetidos(gastos_todos, "gastos")
)

# Años de la comparación interanual:
# anio_base se ajusta por IPC; anio_actual se muestra a valores históricos
anio_base <- 2025
anio_actual <- 2026

solo_anio <- function(df, anio) {
  if (nrow(df) == 0) return(df)
  df %>% filter(substr(periodo, 1, 4) == as.character(anio))
}

# El tablero, el punto de equilibrio y las estadísticas descriptivas
# trabajan solo con el año actual; los archivos de 2025 van a la comparación
ventas_limpias <- solo_anio(ventas_todas, anio_actual)
compras_limpias <- solo_anio(compras_todas, anio_actual)
gastos_limpios <- solo_anio(gastos_todos, anio_actual)

# Versiones sin notas de crédito ni facturas asociadas (solo para estadísticas
# descriptivas). Se calculan una sola vez sobre todo el período, porque la
# factura y su NC pueden estar en meses distintos.
ventas_sin_nc <- excluir_nc_y_facturas(ventas_limpias, "cuit")
compras_sin_nc <- excluir_nc_y_facturas(compras_limpias, "proveedor")

# --- Comparación interanual ajustada por IPC -----------------------------------
# IPC nivel general nacional (INDEC, base dic-2016 = 100).
# Archivo: datos/ipc/ipc_2016-2026.xlsx (hoja con las fechas en la fila 1, desde la
# columna B, y la fila "Nivel general" con los índices). Se usa el primer .xlsx de datos/ipc.
leer_ipc <- function(ruta) {
  tryCatch({
    crudo <- read_excel(ruta, col_names = FALSE, col_types = "text", .name_repair = "unique_quiet")
    fila_nivel <- which(str_to_lower(trimws(crudo[[1]])) == "nivel general")[1]
    
    tibble(
      fecha = parsear_fecha_excel(unlist(crudo[1, -1], use.names = FALSE)),
      indice = suppressWarnings(as.numeric(unlist(crudo[fila_nivel, -1], use.names = FALSE)))
    ) %>%
      filter(!is.na(fecha), !is.na(indice)) %>%
      transmute(periodo = format(fecha, "%Y-%m"), indice) %>%
      distinct(periodo, .keep_all = TRUE)
  }, error = function(e) {
    message(paste("ERROR EN IPC ->", ruta, ":", conditionMessage(e)))
    tibble(periodo = character(), indice = numeric())
  })
}

archivo_ipc <- list.files(path = "datos/ipc", pattern = "^[^~].*\\.xlsx?$", full.names = TRUE, ignore.case = TRUE)[1]
ipc_tabla <- if (is.na(archivo_ipc)) {
  tibble(periodo = character(), indice = numeric())
} else {
  leer_ipc(archivo_ipc)
}

# Coeficiente por mes = IPC del mes en el año actual / IPC del mismo mes en el año base
coef_ipc <- tibble(mes = sprintf("%02d", 1:12)) %>%
  mutate(
    indice_base = ipc_tabla$indice[match(paste0(anio_base, "-", mes), ipc_tabla$periodo)],
    indice_actual = ipc_tabla$indice[match(paste0(anio_actual, "-", mes), ipc_tabla$periodo)],
    coef = indice_actual / indice_base
  )

sumar_mensual <- function(df, columna, por_unidad = TRUE) {
  if (nrow(df) == 0) {
    return(tibble(periodo = character(), unidad_negocio = character(), monto = numeric()))
  }
  if (por_unidad) {
    df %>%
      group_by(periodo, unidad_negocio) %>%
      summarise(monto = sum(.data[[columna]], na.rm = TRUE), .groups = "drop")
  } else {
    # Las compras no se registran por unidad de negocio
    df %>%
      group_by(periodo) %>%
      summarise(monto = sum(.data[[columna]], na.rm = TRUE), .groups = "drop") %>%
      mutate(unidad_negocio = NA_character_)
  }
}

# Compras de mercadería por unidad: el libro de IVA Compras no trae la unidad de negocio,
# así que cada mes se reparte según la participación de cada unidad en las ventas de ese
# mes (mismo criterio que el CMV del Punto de Equilibrio). Es una estimación.
# Solo se consideran las compras del rubro mercadería (rubros_mercaderia).

# Participación de cada unidad en las ventas de cada mes (criterio de reparto)
participacion_mensual <- sumar_mensual(ventas_todas, "neto_gravado") %>%
  group_by(periodo) %>%
  mutate(participacion = if (sum(monto) > 0) monto / sum(monto) else NA_real_) %>%
  ungroup() %>%
  select(periodo, unidad_negocio, participacion)

compras_por_unidad <- inner_join(
  sumar_mensual(
    compras_todas %>% filter(rubro %in% rubros_mercaderia),
    "neto_gravado",
    por_unidad = FALSE
  ) %>%
    select(periodo, monto_compras = monto),
  participacion_mensual,
  by = "periodo"
) %>%
  transmute(periodo, unidad_negocio, monto = monto_compras * participacion)

# Totales mensuales de todos los años cargados (mismos criterios que los KPI del tablero)
resumen_mensual <- bind_rows(
  sumar_mensual(ventas_todas, "neto_gravado") %>% mutate(concepto = "Ventas"),
  compras_por_unidad %>% mutate(concepto = "Compras"),
  sumar_mensual(gastos_todos, "monto") %>% mutate(concepto = "Gastos")
) %>%
  mutate(anio = substr(periodo, 1, 4), mes = substr(periodo, 6, 7))

etiqueta_base <- paste0(anio_base, " (ajustado por IPC)")
etiqueta_actual <- paste0(anio_actual, " (histórico)")
meses_es <- c("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")

# Solo se comparan meses que tienen datos en ambos años
armar_comparacion <- function(resumen) {
  resumen <- resumen %>%
    group_by(anio, mes, concepto) %>%
    summarise(monto = sum(monto, na.rm = TRUE), .groups = "drop")
  
  inner_join(
    resumen %>% filter(anio == as.character(anio_base)) %>% select(mes, concepto, monto_base = monto),
    resumen %>% filter(anio == as.character(anio_actual)) %>% select(mes, concepto, monto_actual = monto),
    by = c("mes", "concepto")
  ) %>%
    left_join(coef_ipc %>% select(mes, coef), by = "mes") %>%
    mutate(
      monto_base_aj = monto_base * coef,
      var_real = if_else(monto_base_aj > 0, monto_actual / monto_base_aj - 1, NA_real_),
      concepto = factor(concepto, levels = c("Ventas", "Compras", "Gastos"))
    )
}

# 3. INTERFAZ DE USUARIO (UI) --------------------------------------------------
ui <- page_sidebar(
  theme = mi_tema,
  title = "Control de Gestión: Rentabilidad por Unidades de Negocio",
  
  sidebar = sidebar(
    title = "Filtros de Exploración",
    uiOutput("selector_periodo"),
    uiOutput("selector_unidad")
  ),
  
  uiOutput("alerta_carga"),
  uiOutput("alerta_duplicados"),
  
  navset_card_underline(
    nav_panel(
      title = "Tablero de Rentabilidad",
      layout_columns(
        value_box(
          title = "Ventas",
          value = textOutput("txt_ventas"),
          theme = "primary"
        ),
        value_box(
          title = "Compras (Mercadería)",
          value = textOutput("txt_compras_netas"),
          theme = "secondary"
        ),
        value_box(
          title = "Gastos Operativos",
          value = textOutput("txt_gastos"),
          theme = "info"
        )
      ),
      card(
        card_header("Evolución Mensual de Rentabilidad por Unidad de Negocio (Ventas (−) Compras (−) Gastos)"),
        plotOutput("grafico_rentabilidad", height = "300px")
      ),
      layout_columns(
        card(
          card_header("Ventas por Unidad de Negocio"),
          plotOutput("grafico_barras", height = "280px")
        ),
        card(
          card_header("Top 5 Proveedores de Compras (Mercaderías)"),
          plotOutput("grafico_proveedores", height = "280px")
        )
      )
    ),
    nav_panel(
      title = "Punto de Equilibrio",
      card(
        card_header("Ventas vs. Punto de Equilibrio por Unidad de Negocio"),
        plotOutput("grafico_equilibrio", height = "320px")
      ),
      card(
        card_header("Márgenes de Contribución y Punto de Equilibrio"),
        tableOutput("tabla_equilibrio_ui")
      )
    ),
    nav_panel(
      title = "Comparación Interanual (IPC)",
      uiOutput("alerta_comparacion"),
      card(
        min_height = 480,
        card_header(paste0("Ventas, compras y gastos: ", anio_base, " (ajustado por IPC) vs. ", anio_actual, " (histórico)")),
        plotOutput("grafico_comparacion", height = "400px")
      ),
      card(
        min_height = 420,
        card_header("Detalle mensual y variación real"),
        tableOutput("tabla_comparacion"),
        p(
          class = "text-muted",
          paste0(
            "Coeficiente = IPC del mes en ", anio_actual, " / IPC del mismo mes en ", anio_base,
            ". Variación real = ", anio_actual, " (histórico) / ", anio_base,
            " (ajustado) - 1. Responde a los filtros del panel lateral (el período elige el mismo mes en ambos años). Las compras por unidad son una estimación: cada mes se reparte según la participación de la unidad en las ventas."
          )
        )
      )
    ),
    nav_panel(
      title = "Estadísticas Descriptivas",
      card(
        card_header("Resumen Ventas"),
        tableOutput("tabla_desc_ventas")
      ),
      layout_columns(
        card(
          card_header("Resumen de compras (mercaderías)"),
          tableOutput("tabla_desc_compras"),
          p(
            class = "text-muted",
            "Con una unidad seleccionada, los importes se asignan según la participación de la unidad en las ventas de cada mes (estimación). La cantidad de operaciones corresponde al total de compras del período."
          )
        ),
        card(
          card_header("Resumen de gastos (por rubro)"),
          tableOutput("tabla_desc_gastos")
        )
      )
    ),
    nav_panel(
      title = "Metodología (Selección y Limpieza)",
      card(
        card_header("a. Selección de Datos y Problema en Estudio"),
        p("El objetivo de esta aplicación es evaluar la rentabilidad comercial, la estructura de costos y el punto de equilibrio multi-período por unidad de negocio."),
        p("Fuentes de datos: Libros contables mensuales de IVA Ventas e IVA Compras (AFIP) y planillas mensuales de gastos operativos directos asignados por unidad de negocio (datos/ventas, datos/compras y datos/gastos)."),
        p("Para la comparación interanual se usa además el índice de precios al consumidor (IPC), nivel general nacional, publicado por el INDEC (datos/ipc). Se cargan los períodos de 2025 y 2026: los de 2026 alimentan el tablero, el punto de equilibrio y las estadísticas descriptivas, y los de 2025 se usan solo en la comparación interanual.")
      ),
      card(
        card_header("b. Limpieza y Transformación de Datos"),
        tags$ul(
          tags$li("Ingesta automatizada en lote desde rutas relativas (datos/ventas, datos/compras y datos/gastos), ignorando archivos temporales."),
          tags$li("Control de archivos duplicados: dentro de cada carpeta se comparan los archivos por su contenido (huella MD5) y se excluyen los idénticos, conservando el primero. Si dos archivos distintos corresponden al mismo mes, la aplicación muestra un aviso para su revisión."),
          tags$li("Omisión de encabezados institucionales superiores y filtrado de filas de totales y conceptos impositivos duplicados (IVA)."),
          tags$li("Normalización de fechas seriales de Excel y formatos mixtos hacia estándar ISO (AAAA-MM-DD) para extracción de períodos mensuales."),
          tags$li("Agrupación de puntos de venta: Emisores 3 y 6 como Rosario, Emisores 5 y 7 como Funes, y Emisor 8 como Roldán."),
          tags$li("Cálculo del Punto de Equilibrio combinando el margen de contribución sobre mercaderías, el CMV repartido por participación en ventas y los gastos fijos tomados de la planilla de gastos por unidad."),
          tags$li("En las estadísticas descriptivas se excluyen las notas de crédito y las facturas asociadas, emparejadas por cliente/proveedor e importe, para que no distorsionen los valores mínimo y máximo."),
          tags$li("Los archivos de 2025 se cargan junto con los de 2026 pero se usan solo en la comparación interanual: cada mes de 2025 se ajusta por el coeficiente IPC (nivel general nacional, INDEC) del mismo mes de 2026 y se compara contra el valor histórico de 2026."),
          tags$li("La rentabilidad mensual por unidad se calcula como ventas netas gravadas menos compras de mercadería (rubro Mercaderias) menos gastos operativos. Las compras se asignan a cada unidad según su participación en las ventas del mes (estimación); el rubro Varios no se incluye porque ya está cubierto por la planilla de gastos."),
          tags$li("En las estadísticas descriptivas, los conceptos de la planilla de gastos se agrupan en tres rubros: Servicios (EPE, telefonía, internet, alarmas, limpieza, Sinergia, refrigerios y varios), SUSS (F931, AEC, FAECYS y Estrella) e Impuestos (DREI, IIBB, Impuesto al Valor Agregado, IIGG y autónomos). Los conceptos que no pertenecen a ningún grupo, como los sueldos, se muestran con su nombre original.")
        )
      )
    )
  )
)


# 4. LÓGICA DEL SERVIDOR (SERVER) ----------------------------------------------
server <- function(input, output, session) {
  
  output$alerta_carga <- renderUI({
    if (nrow(ventas_limpias) == 0 || nrow(compras_limpias) == 0) {
      div(
        class = "alert alert-warning",
        "⚠️ No se detectaron archivos válidos o las carpetas de datos están vacías."
      )
    } else {
      NULL
    }
  })
  
  # Avisos del control de duplicados (se calculan una sola vez al cargar los datos)
  output$alerta_duplicados <- renderUI({
    avisos <- list()
    
    if (nrow(archivos_excluidos) > 0) {
      avisos <- c(avisos, list(div(
        class = "alert alert-info",
        tags$strong("Archivos duplicados excluidos: "),
        "se usó solo el primero de cada grupo de archivos idénticos.",
        tags$ul(lapply(seq_len(nrow(archivos_excluidos)), function(i) {
          tags$li(paste0(
            archivos_excluidos$carpeta[i], "/", archivos_excluidos$archivo[i],
            " (igual a ", archivos_excluidos$igual_a[i], ")"
          ))
        }))
      )))
    }
    
    if (nrow(periodos_duplicados) > 0) {
      avisos <- c(avisos, list(div(
        class = "alert alert-warning",
        tags$strong("Revisar: "),
        "hay archivos distintos que corresponden al mismo mes, por lo que los datos de ese mes podrían estar duplicados.",
        tags$ul(lapply(seq_len(nrow(periodos_duplicados)), function(i) {
          tags$li(paste0(
            periodos_duplicados$carpeta[i], " (", periodos_duplicados$periodo[i], "): ",
            periodos_duplicados$archivos[i]
          ))
        }))
      )))
    }
    
    if (length(avisos) == 0) NULL else tagList(avisos)
  })
  
  output$selector_periodo <- renderUI({
    if (nrow(ventas_limpias) == 0) {
      return(selectInput("filtro_periodo", "Seleccionar Período Mensual:", choices = "Sin Datos"))
    }
    periodos_disponibles <- unique(ventas_limpias$periodo) %>% sort(decreasing = TRUE)
    selectInput(
      "filtro_periodo",
      "Seleccionar Período Mensual:",
      choices = c("Todos los Meses", periodos_disponibles),
      selected = "Todos los Meses"
    )
  })
  
  output$selector_unidad <- renderUI({
    if (nrow(ventas_limpias) == 0) {
      return(selectInput("filtro_unidad", "Filtrar por Unidad de Negocio:", choices = "Sin Datos"))
    }
    selectInput(
      "filtro_unidad",
      "Filtrar por Unidad de Negocio:",
      choices = c("Todas las Unidades", unique(ventas_limpias$unidad_negocio)),
      selected = "Todas las Unidades"
    )
  })
  
  ventas_filtradas <- reactive({
    req(input$filtro_periodo, input$filtro_unidad)
    df <- ventas_limpias
    if (nrow(df) == 0 || input$filtro_periodo == "Sin Datos") return(df)
    
    if (input$filtro_periodo != "Todos los Meses") {
      df <- df %>% filter(periodo == input$filtro_periodo)
    }
    if (input$filtro_unidad != "Todas las Unidades") {
      df <- df %>% filter(unidad_negocio == input$filtro_unidad)
    }
    df
  })
  
  compras_filtradas <- reactive({
    req(input$filtro_periodo)
    df <- compras_limpias
    if (nrow(df) == 0 || input$filtro_periodo == "Sin Datos") return(df)
    
    if (input$filtro_periodo != "Todos los Meses") {
      df <- df %>% filter(periodo == input$filtro_periodo)
    }
    df
  })
  
  gastos_filtrados <- reactive({
    req(input$filtro_periodo, input$filtro_unidad)
    df <- gastos_limpios
    if (nrow(df) == 0 || input$filtro_periodo == "Sin Datos") return(df)
    
    if (input$filtro_periodo != "Todos los Meses") {
      df <- df %>% filter(periodo == input$filtro_periodo)
    }
    if (input$filtro_unidad != "Todas las Unidades") {
      df <- df %>% filter(unidad_negocio == input$filtro_unidad)
    }
    df
  })
  
  # Ventas y compras sin NC ni facturas asociadas (para estadísticas descriptivas)
  ventas_desc <- reactive({
    req(input$filtro_periodo, input$filtro_unidad)
    df <- ventas_sin_nc
    if (nrow(df) == 0 || input$filtro_periodo == "Sin Datos") return(df)
    
    if (input$filtro_periodo != "Todos los Meses") {
      df <- df %>% filter(periodo == input$filtro_periodo)
    }
    if (input$filtro_unidad != "Todas las Unidades") {
      df <- df %>% filter(unidad_negocio == input$filtro_unidad)
    }
    df
  })
  
  compras_desc <- reactive({
    req(input$filtro_periodo)
    df <- compras_sin_nc
    if (nrow(df) == 0 || input$filtro_periodo == "Sin Datos") return(df)
    
    if (input$filtro_periodo != "Todos los Meses") {
      df <- df %>% filter(periodo == input$filtro_periodo)
    }
    df
  })
  
  output$txt_ventas <- renderText({
    if (nrow(ventas_filtradas()) == 0) return("$0,00")
    dollar(sum(ventas_filtradas()$neto_gravado, na.rm = TRUE), prefix = "$", big.mark = ".", decimal.mark = ",")
  })
  
  # Compras de mercadería. Con una unidad seleccionada se usa la estimación por
  # participación en las ventas del mes (las compras no traen unidad de negocio).
  output$txt_compras_netas <- renderText({
    req(input$filtro_periodo, input$filtro_unidad)
    
    total <- if (input$filtro_unidad %in% c("Todas las Unidades", "Sin Datos")) {
      compras_filtradas() %>%
        filter(rubro %in% rubros_mercaderia) %>%
        pull(neto_gravado) %>%
        sum(na.rm = TRUE)
    } else {
      df <- compras_por_unidad %>%
        filter(
          substr(periodo, 1, 4) == as.character(anio_actual),
          unidad_negocio == input$filtro_unidad
        )
      if (input$filtro_periodo != "Todos los Meses") {
        df <- df %>% filter(periodo == input$filtro_periodo)
      }
      sum(df$monto, na.rm = TRUE)
    }
    
    dollar(total, prefix = "$", big.mark = ".", decimal.mark = ",")
  })
  
  output$txt_gastos <- renderText({
    if (nrow(gastos_filtrados()) == 0) return("$0,00")
    dollar(sum(gastos_filtrados()$monto, na.rm = TRUE), prefix = "$", big.mark = ".", decimal.mark = ",")
  })
  
  # Evolución mensual de rentabilidad = ventas - compras de mercadería (estimadas
  # por unidad) - gastos operativos
  output$grafico_rentabilidad <- renderPlot({
    req(input$filtro_periodo, input$filtro_unidad)
    
    datos <- resumen_mensual %>%
      filter(anio == as.character(anio_actual), !is.na(unidad_negocio)) %>%
      pivot_wider(
        id_cols = c(periodo, unidad_negocio),
        names_from = concepto,
        values_from = monto,
        values_fill = 0
      ) %>%
      mutate(
        Ventas = coalesce(Ventas, 0),
        Compras = coalesce(Compras, 0),
        Gastos = coalesce(Gastos, 0),
        rentabilidad = Ventas - Compras - Gastos
      )
    
    if (input$filtro_unidad != "Todas las Unidades") {
      datos <- datos %>% filter(unidad_negocio == input$filtro_unidad)
    }
    if (!input$filtro_periodo %in% c("Todos los Meses", "Sin Datos")) {
      datos <- datos %>% filter(periodo == input$filtro_periodo)
    }
    req(nrow(datos) > 0)
    
    datos <- datos %>%
      mutate(mes = factor(
        as.integer(substr(periodo, 6, 7)),
        levels = 1:12, labels = meses_es
      ))
    
    ggplot(datos, aes(x = mes, y = rentabilidad, fill = unidad_negocio)) +
      geom_hline(yintercept = 0, color = "grey40") +
      geom_col(position = position_dodge(width = 0.8), width = 0.7) +
      scale_fill_manual(values = colores_unidad) +
      scale_x_discrete(drop = TRUE) +
      scale_y_continuous(labels = etiqueta_miles) +
      theme_minimal() +
      theme(
        panel.grid.major.x = element_blank(),
        axis.title.y = element_text(size = 10)) +
      labs(x = NULL, y = "Rentabilidad (miles $)", fill = "Unidad de Negocio")
  })
  
  output$grafico_barras <- renderPlot({
    req(nrow(ventas_filtradas()) > 0)
    ventas_filtradas() %>%
      group_by(unidad_negocio) %>%
      summarise(total_neto = sum(neto_gravado, na.rm = TRUE), .groups = "drop") %>%
      ggplot(aes(x = reorder(unidad_negocio, total_neto), y = total_neto, fill = unidad_negocio)) +
      geom_col(show.legend = FALSE) +
      coord_flip() +
      scale_fill_manual(values = colores_unidad) +
      theme_minimal() +
      labs(x = NULL, y = "Ventas (miles de $)") +
      scale_y_continuous(labels = etiqueta_miles)
  })
  
  output$grafico_proveedores <- renderPlot({
    datos <- compras_filtradas() %>%
      filter(rubro %in% rubros_mercaderia)
    req(nrow(datos) > 0)
    
    datos %>%
      group_by(proveedor) %>%
      summarise(gasto = sum(neto_gravado, na.rm = TRUE), .groups = "drop") %>%
      slice_max(gasto, n = 5, with_ties = FALSE) %>%
      ggplot(aes(x = reorder(str_trunc(proveedor, 28), gasto), y = gasto)) +
      geom_col(fill = "#e74c3c") +
      coord_flip() +
      theme_minimal() +
      theme(plot.margin = margin(t = 5, r = 25, b = 5, l = 5)) +
      labs(x = NULL, y = "Monto (miles de $)") +
      scale_y_continuous(
        labels = etiqueta_miles,
        expand = expansion(mult = c(0, 0.08))
      )
  })
  
  datos_equilibrio <- reactive({
    req(input$filtro_periodo, nrow(ventas_limpias) > 0, nrow(compras_limpias) > 0)
    
    v_periodo <- ventas_limpias
    c_periodo <- compras_limpias
    g_periodo <- gastos_limpios
    
    if (input$filtro_periodo != "Todos los Meses") {
      v_periodo <- v_periodo %>% filter(periodo == input$filtro_periodo)
      c_periodo <- c_periodo %>% filter(periodo == input$filtro_periodo)
      if (nrow(g_periodo) > 0) {
        g_periodo <- g_periodo %>% filter(periodo == input$filtro_periodo)
      }
    }
    
    ventas_tot <- sum(v_periodo$neto_gravado, na.rm = TRUE)
    cmv_tot <- c_periodo %>%
      filter(rubro %in% rubros_mercaderia) %>%
      pull(neto_gravado) %>%
      sum(na.rm = TRUE)
    
    # Margen de contribución global (% sobre ventas)
    mc_pct <- if (ventas_tot > 0) (ventas_tot - cmv_tot) / ventas_tot else NA_real_
    
    gf_directos <- if (nrow(g_periodo) > 0) {
      g_periodo %>%
        group_by(unidad_negocio) %>%
        summarise(gastos_directos = sum(monto, na.rm = TRUE), .groups = "drop")
    } else {
      tibble(unidad_negocio = character(), gastos_directos = numeric())
    }
    
    res <- v_periodo %>%
      group_by(unidad_negocio) %>%
      summarise(ventas_reales = sum(neto_gravado, na.rm = TRUE), .groups = "drop") %>%
      left_join(gf_directos, by = "unidad_negocio") %>%
      mutate(
        gastos_directos = replace_na(gastos_directos, 0),
        participacion = ventas_reales / sum(ventas_reales),
        cmv_asignado = cmv_tot * participacion,
        margen_contribucion = ventas_reales - cmv_asignado,
        gastos_fijos = gastos_directos,
        resultado_neto = margen_contribucion - gastos_fijos,
        punto_equilibrio = if (!is.na(mc_pct) && mc_pct > 0) {
          gastos_fijos / mc_pct
        } else {
          NA_real_
        }
      )
    
    if (!is.null(input$filtro_unidad) && input$filtro_unidad != "Todas las Unidades") {
      res <- res %>% filter(unidad_negocio == input$filtro_unidad)
    }
    res
  })
  
  output$tabla_equilibrio_ui <- renderTable({
    req(nrow(datos_equilibrio()) > 0)
    datos <- datos_equilibrio()
    
    fila_total <- datos %>%
      summarise(
        unidad_negocio = "TOTAL",
        across(
          c(ventas_reales, cmv_asignado, margen_contribucion,
            gastos_fijos, resultado_neto),
          ~ sum(.x, na.rm = TRUE)
        ),
        punto_equilibrio = if (all(is.na(punto_equilibrio))) {
          NA_real_
        } else {
          sum(punto_equilibrio, na.rm = TRUE)
        }
      )
    
    # El total tiene sentido porque hay más de una unidad en la tabla
    tabla_base <- if (nrow(datos) > 1) bind_rows(datos, fila_total) else datos
    
    tabla_base %>%
      transmute(
        "Unidad de Negocio" = unidad_negocio,
        "Ventas" = fmt_pesos(ventas_reales),
        "(-) Costo Mercad. Vendidas" = fmt_pesos(cmv_asignado),
        "(=) Margen de Contribución" = fmt_pesos(margen_contribucion),
        "(-) Gastos Operativos" = fmt_pesos(gastos_fijos),
        "(=) Resultado Neto" = fmt_pesos(resultado_neto),
        "Punto de Equilibrio ($)" = fmt_pesos(punto_equilibrio)
      )
  })
  
  output$tabla_desc_ventas <- renderTable({
    req(nrow(ventas_desc()) > 0)
    ventas_desc() %>%
      group_by(unidad_negocio) %>%
      summarise(
        Operaciones = n(),
        Media = mean(neto_gravado, na.rm = TRUE),
        Mediana = median(neto_gravado, na.rm = TRUE),
        `Desvío estándar` = sd(neto_gravado, na.rm = TRUE),
        Mínimo = min(neto_gravado, na.rm = TRUE),
        Máximo = max(neto_gravado, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      mutate(
        Operaciones = format(Operaciones, big.mark = ".", decimal.mark = ","),
        across(c(Media, Mediana, `Desvío estándar`, Mínimo, Máximo), fmt_pesos)
      ) %>%
      rename(`Unidad de Negocio` = unidad_negocio)
  })
  
  output$tabla_desc_compras <- renderTable({
    req(input$filtro_unidad)
    datos <- compras_desc() %>% filter(rubro %in% rubros_mercaderia)
    req(nrow(datos) > 0)
    
    # Con una unidad seleccionada, cada compra se asigna según la participación
    # de esa unidad en las ventas del mes de la compra (estimación)
    if (!input$filtro_unidad %in% c("Todas las Unidades", "Sin Datos")) {
      datos <- datos %>%
        left_join(
          participacion_mensual %>%
            filter(unidad_negocio == input$filtro_unidad) %>%
            select(periodo, participacion),
          by = "periodo"
        ) %>%
        mutate(neto_gravado = neto_gravado * replace_na(participacion, 0))
    }
    
    datos %>%
      group_by(rubro) %>%
      summarise(
        Operaciones = n(),
        Total = sum(neto_gravado, na.rm = TRUE),
        Mediana = median(neto_gravado, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      arrange(desc(Total)) %>%
      slice_head(n = 8) %>%
      mutate(
        Operaciones = format(Operaciones, big.mark = ".", decimal.mark = ","),
        across(c(Total, Mediana), fmt_pesos)
      ) %>%
      rename(Rubro = rubro)
  })
  
  # Resumen de gastos agrupado en SERVICIOS, SUSS e IMPUESTOS.
  # Los conceptos que no entran en ningún grupo (por ejemplo SUELDOS) conservan su nombre.
  # Responde a los filtros de período y unidad de negocio.
  output$tabla_desc_gastos <- renderTable({
    req(nrow(gastos_filtrados()) > 0)
    
    gastos_filtrados() %>%
      mutate(
        rubro_gasto = case_when(
          str_detect(
            concepto,
            regex("\\bEPE\\b|telef|internet|alarma|limpie|sinergia|refriger|\\bVARIOS\\b", ignore_case = TRUE)
          ) ~ "SERVICIOS",
          str_detect(
            concepto,
            regex("F\\.?\\s?931|\\bAEC\\b|FAECYS|ESTRELLA", ignore_case = TRUE)
          ) ~ "SUSS",
          str_detect(
            concepto,
            regex("\\bDREI\\b|\\bIIBB\\b|impuesto\\s*valor\\s*agregado|\\bIVA\\b|\\bIIGG\\b|aut[oó]nomo", ignore_case = TRUE)
          ) ~ "IMPUESTOS",
          TRUE ~ concepto
        )
      ) %>%
      group_by(rubro_gasto) %>%
      summarise(Total = sum(monto, na.rm = TRUE), .groups = "drop") %>%
      arrange(desc(Total)) %>%
      mutate(
        `% del total` = percent(Total / sum(Total), accuracy = 0.1, decimal.mark = ",", big.mark = "."),
        Total = fmt_pesos(Total)
      ) %>%
      rename(Rubro = rubro_gasto)
  })
  
  output$grafico_equilibrio <- renderPlot({
    req(nrow(datos_equilibrio()) > 0)
    datos_equilibrio() %>%
      select(unidad_negocio, "Ventas" = ventas_reales, "Punto de Equilibrio" = punto_equilibrio) %>%
      pivot_longer(-unidad_negocio, names_to = "concepto", values_to = "monto") %>%
      ggplot(aes(x = unidad_negocio, y = monto, fill = concepto)) +
      geom_col(position = "dodge") + theme_minimal() +
      scale_fill_manual(values = c("Ventas" = "#18bc9c", "Punto de Equilibrio" = "#2c3e50")) +
      labs(x = NULL, y = "Monto (miles de $)", fill = NULL) +
      scale_y_continuous(labels = etiqueta_miles)
  })
  
  # --- Comparación interanual ---
  comparacion_filtrada <- reactive({
    req(input$filtro_periodo, input$filtro_unidad)
    df <- resumen_mensual
    if (input$filtro_unidad != "Todas las Unidades") {
      df <- df %>% filter(unidad_negocio == input$filtro_unidad)
    }
    if (!input$filtro_periodo %in% c("Todos los Meses", "Sin Datos")) {
      df <- df %>% filter(mes == substr(input$filtro_periodo, 6, 7))
    }
    armar_comparacion(df)
  })
  
  output$alerta_comparacion <- renderUI({
    if (nrow(ipc_tabla) == 0 || all(is.na(coef_ipc$coef))) {
      div(
        class = "alert alert-warning",
        "⚠️ Faltan los índices IPC (datos/ipc/ipc_2016-2026.xlsx) de ambos años: sin ellos no se puede calcular el ajuste."
      )
    } else if (nrow(comparacion_filtrada()) == 0) {
      div(
        class = "alert alert-warning",
        paste0("⚠️ No hay meses con datos cargados en ambos años (", anio_base, " y ", anio_actual, ").")
      )
    } else if (!is.null(input$filtro_unidad) && !input$filtro_unidad %in% c("Todas las Unidades", "Sin Datos")) {
      div(
        class = "alert alert-info",
        "Las compras no se registran por unidad: con una unidad seleccionada se estiman repartiendo las compras de cada mes según la participación de esa unidad en las ventas del mes."
      )
    } else {
      NULL
    }
  })
  
  output$grafico_comparacion <- renderPlot({
    datos <- comparacion_filtrada() %>% filter(!is.na(coef))
    req(nrow(datos) > 0)
    
    # Montos en miles de pesos para que los ejes y las etiquetas sean legibles
    etiqueta_m <- label_number(
      scale = 1e-3, prefix = "$",
      accuracy = 1, big.mark = ".", decimal.mark = ","
    )
    
    datos_graf <- datos %>%
      select(mes, concepto, monto_base_aj, monto_actual) %>%
      pivot_longer(c(monto_base_aj, monto_actual), names_to = "serie", values_to = "monto") %>%
      mutate(
        serie = if_else(serie == "monto_base_aj", etiqueta_base, etiqueta_actual),
        serie = factor(serie, levels = c(etiqueta_base, etiqueta_actual)),
        mes = factor(mes, levels = sprintf("%02d", 1:12), labels = meses_es)
      )
    
    # Etiquetas de valor solo cuando hay pocos meses (si no, se superponen)
    capa_etiquetas <- if (n_distinct(datos_graf$mes) <= 4) {
      geom_text(
        aes(label = etiqueta_m(monto)),
        position = position_dodge(width = 0.8), vjust = -0.5, size = 3.6
      )
    } else {
      NULL
    }
    
    ggplot(datos_graf, aes(x = mes, y = monto, fill = serie)) +
      geom_col(position = position_dodge(width = 0.8), width = 0.7) +
      capa_etiquetas +
      facet_wrap(~ concepto, nrow = 1, scales = "free_y") +
      scale_fill_manual(values = setNames(c("#95a5a6", "#18bc9c"), c(etiqueta_base, etiqueta_actual))) +
      scale_y_continuous(labels = etiqueta_m, expand = expansion(mult = c(0, 0.15))) +
      theme_minimal(base_size = 13) +
      theme(
        legend.position = "top",
        strip.text = element_text(face = "bold", size = 13),
        panel.grid.major.x = element_blank(),
        panel.grid.minor = element_blank()
      ) +
      labs(x = NULL, y = "Monto (miles de $)", fill = NULL)
  })
  
  output$tabla_comparacion <- renderTable({
    req(nrow(comparacion_filtrada()) > 0)
    comparacion_filtrada() %>%
      arrange(mes, concepto) %>%
      transmute(
        "Mes" = meses_es[as.integer(mes)],
        "Concepto" = as.character(concepto),
        !!paste0(anio_base, " histórico") := fmt_pesos(monto_base),
        "Coef. IPC" = if_else(is.na(coef), "Sin IPC", number(coef, accuracy = 0.0001, decimal.mark = ",")),
        !!paste0(anio_base, " ajustado") := fmt_pesos(monto_base_aj),
        !!paste0(anio_actual, " histórico") := fmt_pesos(monto_actual),
        "Variación real" = if_else(is.na(var_real), "No calculable", percent(var_real, accuracy = 0.1, decimal.mark = ",", big.mark = "."))
      )
  })
  
}  # fin de server

# 5. LANZAMIENTO DE LA APLICACIÓN ----------------------------------------------
shinyApp(ui = ui, server = server)
