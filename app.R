# ==============================================================================
# TRABAJO FINAL: ANÁLISIS INTELIGENTE DE DATOS (2026)
# PRODUCTO: Aplicación Web Shiny Multi-Período para Rentabilidad Comercial
# DATASET: IVA Compras e IVA Ventas en Lote (Formatos Excel .xlsx)
# INTEGRANTES / AUTOR: CALVINO JOAQUIN PABLO
# ==============================================================================

# 1. CARGA DE PAQUETES Y CONFIGURACIÓN -----------------------------------------
library(shiny)
library(bslib)
library(tidyverse)
library(lubridate)
library(scales)
library(readxl)   # Paquete oficial del ecosistema Tidyverse para ingesta de Excel

# Configuración del tema visual moderno y profesional
mi_tema <- bs_theme(
  version = 5,
  bootswatch = "flatly",
  primary = "#2c3e50",
  secondary = "#18bc9c"
)

# 2. ETAPA: SELECCIÓN Y LIMPIEZA (PROCESAMIENTO EN LOTE) ------------------------

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

# Función para Limpiar Ventas
procesar_archivo_ventas <- function(ruta_archivo) {
  tryCatch({
    read_excel(ruta_archivo, skip = 3, col_types = "text") %>%
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
          punto_venta %in% c(3, 6) ~ "Retail (Mostrador)",
          punto_venta %in% c(5, 7) ~ "E-commerce",
          punto_venta == 8 ~ "Servicio Técnico",
          TRUE ~ "Otros"
        )
      ) %>%
      filter(!is.na(fecha))
  }, error = function(e) {
    message(paste("ERROR EN VENTAS ->", ruta_archivo, ":", conditionMessage(e)))
    tibble()
  })
}

# Función para Limpiar Compras (tolerante a columnas opcionales de percepciones)
procesar_archivo_compras <- function(ruta_archivo) {
  tryCatch({
    df_raw <- read_excel(ruta_archivo, skip = 3, col_types = "text") %>%
      filter(!is.na(Fecha), !is.na(Comprobante), Fecha != "Totales:")
    
    # Crear columnas de percepciones en 0 si no existen en ese mes
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

# Ingesta desde carpetas relativas
archivos_ventas <- list.files(path = "datos/ventas", pattern = "^[^~].*\\.xlsx$", full.names = TRUE)
archivos_compras <- list.files(path = "datos/compras", pattern = "^[^~].*\\.xlsx$", full.names = TRUE)

ventas_limpias <- map(archivos_ventas, procesar_archivo_ventas) %>% list_rbind()
compras_limpias <- map(archivos_compras, procesar_archivo_compras) %>% list_rbind()



# 3. INTERFAZ DE USUARIO (UI) --------------------------------------------------
ui <- page_sidebar(
  theme = mi_tema,
  title = "Control de Gestión: Rentabilidad Multi-Período",
  
  # Panel lateral con controles dinámicos (Componente c: Exploración)
  sidebar = sidebar(
    title = "Filtros de Exploración",
    uiOutput("selector_periodo"),
    uiOutput("selector_unidad")
  ),
  
  # Alerta en caso de carpetas vacías
  uiOutput("alerta_carga"),
  
  # Pestañas principales
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
          plotOutput("grafico_torta", height = "280px")
        ),
        card(
          card_header("Top 5 Proveedores por Monto Neto"),
          plotOutput("grafico_proveedores", height = "280px")
        )
      )
    ),
    nav_panel(
      title = "Metodología (Selección y Limpieza)",
      card(
        card_header("a. Selección de Datos y Problema en Estudio"),
        p("El objetivo de esta aplicación es evaluar la rentabilidad comercial y la evolución operativa multi-período de una empresa comercial."),
        p("Fuente de datos: Registros contables oficiales correspondientes a los Libros de IVA Ventas e IVA Compras (exportados desde sistema de gestión / AFIP), organizados por período mensual.")
      ),
      card(
        card_header("b. Limpieza y Transformación de Datos"),
        tags$ul(
          tags$li("Ingesta automatizada en lote desde rutas relativas (datos/ventas y datos/compras)."),
          tags$li("Omisión de encabezados institucionales superiores (skip = 3) y filtrado de filas de totales al pie de cada libro."),
          tags$li("Normalización de fechas seriales de Excel y formatos mixtos hacia estándar ISO (AAAA-MM-DD) para extracción de períodos mensuales."),
          tags$li("Conversión numérica de importes netos, IVA y percepciones, respetando el signo negativo de las Notas de Crédito."),
          tags$li("Segmentación de ventas por punto de venta (Emisor) para identificar las unidades de negocio.")
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
  
  output$grafico_torta <- renderPlot({
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
      ggplot(aes(x = reorder(proveedor, gasto), y = gasto)) +
      geom_col(fill = "#e74c3c") +
      coord_flip() +
      theme_minimal() +
      labs(x = NULL, y = "Monto Neto ($)") +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
  })
}

#5. LANZAMIENTO DE LA APLICACIÓN ----------------------------------------------
  
  shinyApp(ui = ui, server = server)