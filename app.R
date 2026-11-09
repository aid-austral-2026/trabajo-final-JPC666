# ==============================================================================
# TRABAJO FINAL: ANÁLISIS INTELIGENTE DE DATOS (2026)
# PRODUCTO: Aplicación Web Shiny Multi-Período para Rentabilidad Comercial
# DATASET: IVA Compras e IVA Ventas en Lote (Formatos Excel .xlsx)
# INTEGRANTES / AUTOR: CALVINO NICOLAS PEDRO
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

# 2. ETAPA: SELECCIÓN Y LIMPIEZA (PROCESAMIENTO EN LOTE SEGURO) ----------------

# --- Función para Limpiar Ventas con Validación de Errores ---
procesar_archivo_ventas <- function(ruta_archivo) {
  tryCatch({
    df <- read_excel(ruta_archivo, skip = 3) %>%
      filter(!is.na(Dia), Dia != "Totales:") %>%
      rename(
        dia = Dia, tipo = Tipo, punto_venta = Emisor,
        comprobante = Numero, cliente = Cliente,
        neto_gravado = `Neto Gravado`, iva_21 = `IVA 21.00%`, total = Total
      )
    
    # Manejo robusto de fechas (por si Excel lee números o texto/Date)
    df <- df %>% mutate(
      num_dia = suppressWarnings(as.numeric(dia)),
      fecha = if_else(!is.na(num_dia), 
                      as.Date(num_dia, origin = "1899-12-30"), 
                      as.Date(dia)),
      periodo = format(fecha, "%Y-%m"),
      neto_gravado = as.numeric(neto_gravado),
      iva_21 = as.numeric(iva_21),
      total = as.numeric(total),
      punto_venta = as.numeric(punto_venta),
      # Clasificación de unidades mediante Punto de Venta
      unidad_negocio = case_when(
        punto_venta == 3 ~ "Retail (Mostrador)",
        punto_venta == 5 ~ "E-commerce",
        punto_venta == 8 ~ "Servicio Técnico",
        TRUE ~ "Otros"
      )
    )
    return(df)
  }, error = function(e) {
    warning(paste("Error procesando archivo ventas:", ruta_archivo, "-", e$message))
    return(tibble()) # Retorna tabla vacía si hay error para no colgar la app
  })
}

# --- Función para Limpiar Compras con Validación de Errores ---
procesar_archivo_compras <- function(ruta_archivo) {
  tryCatch({
    df <- read_excel(ruta_archivo, skip = 3) %>%
      filter(!is.na(Fecha), Fecha != "Totales:") %>%
      rename(
        dia = Fecha, proveedor = Proveedor, rubro = Rubro,
        neto_gravado = `Neto Gravado`, total = Total
      )
    
    df <- df %>% mutate(
      num_dia = suppressWarnings(as.numeric(dia)),
      fecha = if_else(!is.na(num_dia), 
                      as.Date(num_dia, origin = "1899-12-30"), 
                      as.Date(dia)),
      periodo = format(fecha, "%Y-%m"),
      neto_gravado = as.numeric(neto_gravado),
      total = as.numeric(total)
    )
    return(df)
  }, error = function(e) {
    warning(paste("Error procesando archivo compras:", ruta_archivo, "-", e$message))
    return(tibble())
  })
}

# --- Inicialización y Creación de Carpetas en Caliente ---
if(!dir.exists("datos/ventas")) dir.create("datos/ventas", recursive = TRUE)
if(!dir.exists("datos/compras")) dir.create("datos/compras", recursive = TRUE)

# Escaneo automatizado de archivos locales
archivos_ventas <- list.files(path = "datos/ventas", pattern = "\\.xlsx$", full.names = TRUE)
archivos_compras <- list.files(path = "datos/compras", pattern = "\\.xlsx$", full.names = TRUE)

# Lectura controlada (si no hay archivos, inicializa dataframes vacíos correctos)
if(length(archivos_ventas) > 0) {
  ventas_limpias <- map_df(archivos_ventas, procesar_archivo_ventas)
} else {
  ventas_limpias <- tibble(fecha = Date(), periodo = character(), unidad_negocio = character(), neto_gravado = numeric(), iva_21 = numeric(), total = numeric())
}

if(length(archivos_compras) > 0) {
  compras_limpias <- map_df(archivos_compras, procesar_archivo_compras)
} else {
  compras_limpias <- tibble(fecha = Date(), periodo = character(), proveedor = character(), neto_gravado = numeric(), total = numeric())
}


# 3. INTERFAZ DE USUARIO (UI) --------------------------------------------------
ui <- page_sidebar(
  theme = mi_tema,
  title = "Control de Gestión: Rentabilidad Multi-Período",
  
  sidebar = sidebar(
    title = "Controles Dinámicos",
    uiOutput("selector_periodo"),
    uiOutput("selector_unidad"),
    hr(),
    p("Estructura en el Repositorio:", style = "font-weight: bold; font-size: 0.9em;"),
    code("├── app.R"), br(),
    code("└── datos/"), br(),
    code("    ├── compras/ (*.xlsx)"), br(),
    code("    └── ventas/  (*.xlsx)"), br(),
    hr(),
    p("Desarrollado para la materia Análisis Inteligente de Datos.", 
      style = "font-size: 0.8em; color: #7f8c8d;")
  ),
  
  navset_card_underline(
    nav_panel(
      title = "Dashboard Rentabilidad",
      uiOutput("alerta_carga"), # Mensaje dinámico si faltan datos
      layout_columns(
        value_box(
          title = "Ventas Netas (Ingresos)",
          value = textOutput("txt_ventas"),
          showcase = bsicons::bs_icon("arrow-up-right-circle-fill"),
          theme = "success"
        ),
        value_box(
          title = "Crédito Fiscal IVA Compras",
          value = textOutput("txt_iva_compras"),
          showcase = bsicons::bs_icon("receipt"),
          theme = "danger"
        ),
        value_box(
          title = "Transacciones Totales",
          value = textOutput("txt_ops"),
          showcase = bsicons::bs_icon("bag-check"),
          theme = "info"
        )
      ),
      card(
        card_header("Evolución de Ventas Diarias Netas por Unidad ($)"),
        plotOutput("grafico_lineas")
      )
    ),
    
    nav_panel(
      title = "Comparativa y Costos",
      layout_columns(
        card(
          card_header("Participación de Unidades en las Ventas Totales"),
          plotOutput("grafico_torta")
        ),
        card(
          card_header("Top 5 Proveedores Clave del Período"),
          plotOutput("grafico_proveedores")
        )
      )
    ),
    
    nav_panel(
      title = "Documentación del Proceso (AID)",
      markdown("
        ### Informe Técnico y Académico
        
        #### a. Selección
        La problemática aborda el control de gestión económica de un comercio unificando su facturación histórica. Las fuentes de información corresponden a los libros oficiales de **IVA Ventas** e **IVA Compras** de **CALVINO NICOLAS PEDRO**.
        
        #### b. Limpieza y Automatización No Trivial
        El pipeline fue diseñado bajo el principio de **reproducibilidad absoluta**:
        *   **Procesamiento en lote:** Utilizando `list.files()` y `map_df()`, el sistema consolida dinámicamente cualquier libro `.xlsx` en los directorios, descartando líneas del sistema (`Totales:`).
        *   **Tipado tolerante a fallos:** Se implementó una lógica dual para parsear las fechas de AFIP tanto si vienen codificadas como números de serie o cadenas estándar de fecha.
        *   **Aislamiento de Errores:** Mediante bloques `tryCatch`, si un archivo presenta fallas, la aplicación las salta permitiendo levantar la interfaz sin colapsar el entorno.
      ")
    )
  )
)


# 4. LÓGICA DEL SERVIDOR (SERVER) ----------------------------------------------
server <- function(input, output, session) {
  
  # Alerta visual en la App por si las carpetas aún no tienen archivos
  output$alerta_carga <- renderUI({
    if(nrow(ventas_limpias) == 0) {
      layout_column_wrap(
        div(class = "alert alert-warning", 
            "⚠️ No se detectaron archivos Excel en las carpetas. Por favor, asegúrese de guardar sus planillas .xlsx dentro de 'datos/ventas/' y 'datos/compras/' en la raíz de su repositorio.")
      )
    } else {
      NULL
    }
  })
  
  output$selector_periodo <- renderUI({
    if(nrow(ventas_limpias) == 0) return(selectInput("filtro_periodo", "Seleccionar Período Mensual:", choices = "Sin Datos"))
    periodos_disponibles <- unique(ventas_limpias$periodo) %>% sort(decreasing = TRUE)
    selectInput("filtro_periodo", "Seleccionar Período Mensual:", choices = c("Todos los Meses", periodos_disponibles), selected = "Todos los Meses")
  })
  
  output$selector_unidad <- renderUI({
    if(nrow(ventas_limpias) == 0) return(selectInput("filtro_unidad", "Filtrar por Unidad de Negocio:", choices = "Sin Datos"))
    selectInput("filtro_unidad", "Filtrar por Unidad de Negocio:", choices = c("Todas las Unidades", unique(ventas_limpias$unidad_negocio)), selected = "Todas las Unidades")
  })
  
  ventas_filtradas <- reactive({
    req(input$filtro_periodo, input$filtro_unidad)
    df <- ventas_limpias
    if(nrow(df) == 0 || input$filtro_periodo == "Sin Datos") return(df)
    
    if (input$filtro_periodo != "Todos los Meses") df <- df %>% filter(periodo == input$filtro_periodo)
    if (input$filtro_unidad != "Todas las Unidades") df <- df %>% filter(unidad_negocio == input$filtro_unidad)
    return(df)
  })
  
  compras_filtradas <- reactive({
    req(input$filtro_periodo)
    df <- compras_limpias
    if(nrow(df) == 0 || input$filtro_periodo == "Sin Datos") return(df)
    
    if (input$filtro_periodo != "Todos los Meses") df <- df %>% filter(periodo == input$filtro_periodo)
    return(df)
  })
  
  output$txt_ventas <- renderText({
    if(nrow(ventas_filtradas()) == 0) return("$0,00")
    dollar(sum(ventas_filtradas()$neto_gravado, na.rm = TRUE), prefix = "$", big.mark = ".", decimal.mark = ",")
  })
  
  output$txt_iva_compras <- renderText({
    if(nrow(compras_filtradas()) == 0) return("$0,00")
    c <- sum(compras_filtradas()$total - compras_filtradas()$neto_gravado, na.rm = TRUE)
    dollar(c, prefix = "$", big.mark = ".", decimal.mark = ",")
  })
  
  output$txt_ops <- renderText({ nrow(ventas_filtradas()) })
  
  output$grafico_lineas <- renderPlot({
    req(nrow(ventas_filtradas()) > 0)
    ventas_filtradas() %>%
      group_by(fecha, unidad_negocio) %>%
      summarise(v_diaria = sum(neto_gravado), .groups = "drop") %>%
      ggplot(aes(x = fecha, y = v_diaria, color = unidad_negocio, group = unidad_negocio)) +
      geom_line(size = 1.1) + geom_point(size = 2) +
      theme_minimal() + labs(x = "Fecha", y = "Monto Neto ($)", color = "Unidad de Negocio") +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
  })
  output$grafico_torta <- renderPlot({
    req(nrow(ventas_filtradas()) > 0)
    ventas_filtradas() %>%
      group_by(unidad_negocio) %>%
      summarise(total_neto = sum(neto_gravado)) %>%
      ggplot(aes(x = reorder(unidad_negocio, total_neto), y = total_neto, fill = unidad_negocio)) +
      geom_col(show.legend = FALSE) + coord_flip() + theme_minimal() + labs(x = NULL, y = "Ventas Netas ($)") +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
  })
  output$grafico_proveedores <- renderPlot({
    req(nrow(compras_filtradas()) > 0)
    compras_filtradas() %>%
      group_by(proveedor) %>%
      summarise(gasto = sum(neto_gravado)) %>%
      slice_max(gasto, n = 5, with_ties = FALSE) %>%
      ggplot(aes(x = reorder(proveedor, gasto), y = gasto)) +
      geom_col(fill = "#e74c3c") + coord_flip() + theme_minimal() + labs(x = NULL, y = "Monto Neto ($)") +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
  })
}
    

# 5. LANZAMIENTO DE LA APLICACIÓN ----------------------------------------------
  
  shinyApp(ui = ui, server = server)
      