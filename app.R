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
library(scales)
library(readxl)

mi_tema <- bs_theme(
  version = 5,
  bootswatch = "flatly",
  primary = "#2c3e50",
  secondary = "#18bc9c"
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

# Rubros de compras que se consideran costo de mercadería vendida (CMV).
# En los libros de compras el rubro viene sin tilde ("Mercaderias").
# El rubro "Varios" no se suma como gasto: ya está cubierto por la planilla de gastos.
rubros_mercaderia <- c("Mercaderias")

# Función para Limpiar Ventas
procesar_archivo_ventas <- function(ruta_archivo) {
  tryCatch({
    read_excel(ruta_archivo, skip = 3, col_types = "text", .name_repair = "unique_quiet") %>%
      filter(!is.na(Fecha), !is.na(Numero), Fecha != "Totales:") %>%
      transmute(
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
          punto_venta == 8 ~ "Roldan",
          TRUE ~ "Otros"
        )
      ) %>%
      filter(!is.na(fecha))
  }, error = function(e) {
    message(paste("ERROR EN VENTAS ->", ruta_archivo, ":", conditionMessage(e)))
    tibble()
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
    tibble()
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
      mutate(periodo = periodo_archivo) %>%
      filter(
        !is.na(concepto),
        !toupper(trimws(concepto)) %in% c("CONCEPTOS", "CONCEPTO", "TOTALES", "TOTAL", "IVA")
      ) %>%
      select(periodo, concepto, `3`, `5`, `8`) %>%
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
          punto_venta == 8 ~ "Roldan",
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

ventas_limpias <- map(archivos_ventas, procesar_archivo_ventas) %>% list_rbind()
compras_limpias <- map(archivos_compras, procesar_archivo_compras) %>% list_rbind()
gastos_limpios <- map(archivos_gastos, procesar_archivo_gastos) %>% list_rbind()

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
  
  navset_card_underline(
    nav_panel(
      title = "Tablero de Rentabilidad",
      layout_columns(
        value_box(
          title = "Ventas Netas Gravadas",
          value = textOutput("txt_ventas"),
          theme = "primary"
        ),
        value_box(
          title = "Compras Netas Gravadas",
          value = textOutput("txt_compras_netas"),
          theme = "secondary"
        ),
        value_box(
          title = "Cantidad de Operaciones (Ventas)",
          value = textOutput("txt_ops"),
          theme = "info"
        )
      ),
      card(
        card_header("Evolución Diaria de Ventas por Unidad de Negocio"),
        plotOutput("grafico_lineas", height = "300px")
      ),
      layout_columns(
        card(
          card_header("Ventas Netas por Unidad de Negocio"),
          plotOutput("grafico_barras", height = "280px")
        ),
        card(
          card_header("Top 5 Proveedores por Compras Netas"),
          plotOutput("grafico_proveedores", height = "280px")
        )
      )
    ),
    nav_panel(
      title = "Punto de Equilibrio",
      card(
        card_header("Ventas Reales vs. Punto de Equilibrio por Unidad de Negocio"),
        plotOutput("grafico_equilibrio", height = "320px")
      ),
      card(
        card_header("Márgenes de Contribución y Punto de Equilibrio"),
        tableOutput("tabla_equilibrio_ui")
      )
    ),
    nav_panel(
      title = "Estadísticas Descriptivas",
      card(
        card_header("Resumen Ventas (monto neto gravado por operación)"),
        tableOutput("tabla_desc_ventas")
      ),
      layout_columns(
        card(
          card_header("Distribución Ventas (monto neto por operación en escala logarítmica)"),
          plotOutput("grafico_boxplot", height = "300px")
        ),
        card(
          card_header("Resumen de compras por rubro"),
          tableOutput("tabla_desc_compras")
        )
      )
    ),
    nav_panel(
      title = "Metodología (Selección y Limpieza)",
      card(
        card_header("a. Selección de Datos y Problema en Estudio"),
        p("El objetivo de esta aplicación es evaluar la rentabilidad comercial, la estructura de costos y el punto de equilibrio multi-período por unidad de negocio."),
        p("Fuentes de datos: Libros contables mensuales de IVA Ventas e IVA Compras (AFIP) y planillas mensuales de gastos operativos directos asignados por unidad de negocio (datos/ventas, datos/compras y datos/gastos).")
      ),
      card(
        card_header("b. Limpieza y Transformación de Datos"),
        tags$ul(
          tags$li("Ingesta automatizada en lote desde rutas relativas (datos/ventas, datos/compras y datos/gastos), ignorando archivos temporales."),
          tags$li("Omisión de encabezados institucionales superiores y filtrado de filas de totales y conceptos impositivos duplicados (IVA)."),
          tags$li("Normalización de fechas seriales de Excel y formatos mixtos hacia estándar ISO (AAAA-MM-DD) para extracción de períodos mensuales."),
          tags$li("Agrupación de puntos de venta: Emisores 3 y 6 como Rosario, Emisores 5 y 7 como Funes, y Emisor 8 como Roldan."),
          tags$li("Cálculo del Punto de Equilibrio combinando el margen de contribución sobre mercaderías, el CMV repartido por participación en ventas y los gastos fijos tomados de la planilla de gastos por unidad.")
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
  
  output$txt_ventas <- renderText({
    if (nrow(ventas_filtradas()) == 0) return("$0,00")
    dollar(sum(ventas_filtradas()$neto_gravado, na.rm = TRUE), prefix = "$", big.mark = ".", decimal.mark = ",")
  })
  
  output$txt_compras_netas <- renderText({
    if (nrow(compras_filtradas()) == 0) return("$0,00")
    dollar(sum(compras_filtradas()$neto_gravado, na.rm = TRUE), prefix = "$", big.mark = ".", decimal.mark = ",")
  })
  
  output$txt_ops <- renderText({
    format(nrow(ventas_filtradas()), big.mark = ".", decimal.mark = ",")
  })
  
  output$grafico_lineas <- renderPlot({
    req(nrow(ventas_filtradas()) > 0)
    ventas_filtradas() %>%
      filter(!is.na(fecha)) %>%
      group_by(fecha, unidad_negocio) %>%
      summarise(v_diaria = sum(neto_gravado, na.rm = TRUE), .groups = "drop") %>%
      ggplot(aes(x = fecha, y = v_diaria, color = unidad_negocio, group = unidad_negocio)) +
      geom_line(linewidth = 1.1) +
      geom_point(size = 2) +
      theme_minimal() +
      labs(x = "Fecha", y = "Monto Neto ($)", color = "Unidad de Negocio") +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
  })
  
  output$grafico_barras <- renderPlot({
    req(nrow(ventas_filtradas()) > 0)
    ventas_filtradas() %>%
      group_by(unidad_negocio) %>%
      summarise(total_neto = sum(neto_gravado, na.rm = TRUE), .groups = "drop") %>%
      ggplot(aes(x = reorder(unidad_negocio, total_neto), y = total_neto, fill = unidad_negocio)) +
      geom_col(show.legend = FALSE) +
      coord_flip() +
      theme_minimal() +
      labs(x = NULL, y = "Ventas Netas ($)") +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
  })
  
  output$grafico_proveedores <- renderPlot({
    req(nrow(compras_filtradas()) > 0)
    compras_filtradas() %>%
      group_by(proveedor) %>%
      summarise(gasto = sum(neto_gravado, na.rm = TRUE), .groups = "drop") %>%
      slice_max(gasto, n = 5, with_ties = FALSE) %>%
      ggplot(aes(x = reorder(str_trunc(proveedor, 28), gasto), y = gasto)) +
      geom_col(fill = "#e74c3c") +
      coord_flip() +
      theme_minimal() +
      labs(x = NULL, y = "Monto Neto ($)") +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
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
    
    # El total solo tiene sentido si hay más de una unidad en la tabla
    tabla_base <- if (nrow(datos) > 1) bind_rows(datos, fila_total) else datos
    
    tabla_base %>%
      transmute(
        "Unidad de Negocio" = unidad_negocio,
        "Ventas Netas" = fmt_pesos(ventas_reales),
        "(-) CMV" = fmt_pesos(cmv_asignado),
        "(=) Margen de Contribución" = fmt_pesos(margen_contribucion),
        "(-) Gastos Fijos (planilla)" = fmt_pesos(gastos_fijos),
        "(=) Resultado Neto" = fmt_pesos(resultado_neto),
        "Punto de Equilibrio ($)" = fmt_pesos(punto_equilibrio)
      )
  })
  
  output$tabla_desc_ventas <- renderTable({
    req(nrow(ventas_filtradas()) > 0)
    ventas_filtradas() %>%
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
  
  output$grafico_boxplot <- renderPlot({
    datos <- ventas_filtradas() %>% filter(neto_gravado > 0)
    req(nrow(datos) > 0)
    # Escala logarítmica porque los montos por operación suelen ser muy asimétricos
    ggplot(datos, aes(x = unidad_negocio, y = neto_gravado, fill = unidad_negocio)) +
      geom_boxplot(show.legend = FALSE) +
      scale_y_log10(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ",")) +
      theme_minimal() +
      labs(x = NULL, y = "Monto neto por operación ($, escala log)")
  })
  
  output$tabla_desc_compras <- renderTable({
    req(nrow(compras_filtradas()) > 0)
    compras_filtradas() %>%
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
  
  output$grafico_equilibrio <- renderPlot({
    req(nrow(datos_equilibrio()) > 0)
    datos_equilibrio() %>%
      select(unidad_negocio, "Ventas Reales" = ventas_reales, "Punto de Equilibrio" = punto_equilibrio) %>%
      pivot_longer(-unidad_negocio, names_to = "concepto", values_to = "monto") %>%
      ggplot(aes(x = unidad_negocio, y = monto, fill = concepto)) +
      geom_col(position = "dodge") + theme_minimal() +
      scale_fill_manual(values = c("Ventas Reales" = "#18bc9c", "Punto de Equilibrio" = "#2c3e50")) +
      labs(x = NULL, y = "Monto Neto ($)", fill = NULL) +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
  })
  
}  # fin de server

# 5. LANZAMIENTO DE LA APLICACIÓN ----------------------------------------------
shinyApp(ui = ui, server = server)

