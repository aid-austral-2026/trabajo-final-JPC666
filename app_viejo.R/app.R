# ==============================================================================
# TRABAJO FINAL: ANÁLISIS INTELIGENTE DE DATOS (2026)
# PRODUCTO: Aplicación Web Shiny Multi-Período para Rentabilidad Comercial
# DATASET: IVA Compras e IVA Ventas en Lote (Formatos Excel .xlsx)
# INTEGRANTES / AUTOR: [CALVIÑO JOAQUIN]
# ==============================================================================

# 1. CARGA DE PAQUETES Y CONFIGURACIÓN -----------------------------------------
library(shiny)
library(bslib)
library(tidyverse)
library(lubridate)
library(scales)
library(readxl)   # Paquete del ecosistema Tidyverse para la ingesta de Excel

# Configuración del tema visual moderno y profesional
mi_tema <- bs_theme(
  version = 5,
  bootswatch = "flatly",
  primary = "#2c3e50",
  secondary = "#18bc9c"
)

# 2. ETAPA: SELECCIÓN Y LIMPIEZA (PROCESAMIENTO EN LOTE) -----------------------

# --- Función para Limpiar y Estandarizar Ventas desde Excel ---
procesar_archivo_ventas <- function(ruta_archivo) {
  read_excel(ruta_archivo, skip = 3) %>%
    filter(!is.na(Dia), Dia != "Totales:") %>%
    rename(
      dia = Dia, tipo = Tipo, punto_venta = Emisor,
      comprobante = Numero, cliente = Cliente,
      neto_gravado = `Neto Gravado`, iva_21 = `IVA 21.00%`, total = Total
    ) %>%
    mutate(
      neto_gravado = as.numeric(neto_gravado),
      iva_21 = as.numeric(iva_21),
      total = as.numeric(total),
      # Los archivos Excel interpretan internamente las fechas como números de serie
      fecha = as.Date(as.numeric(dia), origin = "1899-12-30"),
      periodo = format(fecha, "%Y-%m"),
      # CRITERIO DE NEGOCIO: Clasificación de unidades mediante Punto de Venta
      unidad_negocio = case_when(
        punto_venta == 3 ~ "Retail (Mostrador)",
        punto_venta == 5 ~ "E-commerce",
        punto_venta == 8 ~ "Servicio Técnico",
        TRUE ~ "Otros"
      )
    )
}

# --- Función para Limpiar y Estandarizar Compras desde Excel ---
procesar_archivo_compras <- function(ruta_archivo) {
  read_excel(ruta_archivo, skip = 3) %>%
    filter(!is.na(Fecha), Fecha != "Totales:") %>%
    rename(
      dia = Fecha, proveedor = Proveedor, rubro = Rubro,
      neto_gravado = `Neto Gravado`, total = Total
    ) %>%
    mutate(
      neto_gravado = as.numeric(neto_gravado),
      total = as.numeric(total),
      fecha = as.Date(as.numeric(dia), origin = "1899-12-30"),
      periodo = format(fecha, "%Y-%m")
    )
}

# --- Ingesta Automatizada por Carpetas (Rutas Relativas) ---
# Escanea y concatena dinámicamente todos los archivos mensuales cargados en el repositorio
archivos_ventas <- list.files(path = "datos/ventas", pattern = "\\.xlsx$", full.names = TRUE)
ventas_limpias  <- map_df(archivos_ventas, procesar_archivo_ventas)

archivos_compras <- list.files(path = "datos/compras", pattern = "\\.xlsx$", full.names = TRUE)
compras_limpias  <- map_df(archivos_compras, procesar_archivo_compras)


# 3. INTERFAZ DE USUARIO (UI) --------------------------------------------------
ui <- page_sidebar(
  theme = mi_tema,
  title = "Control de Gestión: Rentabilidad Multi-Período",
  
  sidebar = sidebar(
    title = "Controles Dinámicos",
    # Los filtros se generan de manera reactiva desde el Server según los datos reales
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
    # Pestaña 1: Cuadro de Mando Principal
    nav_panel(
      title = "Dashboard Rentabilidad",
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
    
    # Pestaña 2: Análisis de Composición y Costos
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
    
    # Pestaña 3: Memoria Técnica de las etapas del AID (Requisito Obligatorio)
    nav_panel(
      title = "Documentación del Proceso (AID)",
      markdown("
        ### Informe Técnico y Académico
        
        #### a. Selección
        La problemática aborda el control de gestión económica de un comercio unificando su facturación histórica. Las fuentes de información corresponden a los libros oficiales de **IVA Ventas** e **IVA Compras** exportados mensualmente.
        
        #### b. Limpieza y Automatización No Trivial
        El pipeline fue diseñado bajo el principio de **reproducibilidad absoluta**:
        *   **Procesamiento en lote:** Utilizando `list.files()` y `map_df()` de `purrr`, el sistema detecta de forma dinámica cualquier nuevo archivo de Excel (`.xlsx`) en los directorios, los parsea, descarta líneas vacías o de sumatorias del sistema (`Totales:`) y los acopla en un solo dataframe.
        *   **Tipado nativo:** Mediante el uso de `readxl`, se realiza una lectura robusta de caracteres especiales, traduciendo de forma directa el formato numérico interno de fechas de Excel a objetos cronológicos legibles `Date` en R.
        *   **Clasificación operativa:** Las transacciones se asocian de manera unívoca a sus tres unidades de negocio reales (**Retail, E-commerce y Servicio Técnico**) mapeando el punto de venta emisor.
        
        #### c. Exploración y Comunicación
        Los inputs de la barra lateral se configuran de manera reactiva en el servidor analizando el espectro temporal cargado. La presentación de resultados aprovecha el ecosistema interactivo de Shiny y componentes adaptables de `bslib` para agilizar la interpretación gerencial.
      ")
    )
  )
)


# 4. LÓGICA DEL SERVIDOR (SERVER) ----------------------------------------------
server <- function(input, output, session) {
  
  # --- UI Dinámica (Poblada automáticamente con los meses encontrados en la carpeta) ---
  output$selector_periodo <- renderUI({
    req(nrow(ventas_limpias) > 0)
    periodos_disponibles <- unique(ventas_limpias$periodo) %>% sort(decreasing = TRUE)
    selectInput(
      inputId = "filtro_periodo",
      label = "Seleccionar Período Mensual:",
      choices = c("Todos los Meses", periodos_disponibles),
      selected = "Todos los Meses"
    )
  })
  
  output$selector_unidad <- renderUI({
    selectInput(
      inputId = "filtro_unidad",
      label = "Filtrar por Unidad de Negocio:",
      choices = c("Todas las Unidades", unique(ventas_limpias$unidad_negocio)),
      selected = "Todas las Unidades"
    )
  })
  
  # --- Expresiones Reactivas para Filtrado Eficiente ---
  ventas_filtradas <- reactive({
    req(input$filtro_periodo, input$filtro_unidad)
    df <- ventas_limpias
    
    if (input$filtro_periodo != "Todos los Meses") {
      df <- df %>% filter(periodo == input$filtro_periodo)
    }
    if (input$filtro_unidad != "Todas las Unidades") {
      df <- df %>% filter(unidad_negocio == input$filtro_unidad)
    }
    return(df)
  })
  
  compras_filtradas <- reactive({
    req(input$filtro_periodo)
    df <- compras_limpias
    
    if (input$filtro_periodo != "Todos los Meses") {
      df <- df %>% filter(periodo == input$filtro_periodo)
    }
    return(df)
  })
  
  # --- Renderización de KPIs Financieros ---
  output$txt_ventas <- renderText({
    v <- sum(ventas_filtradas()$neto_gravado, na.rm = TRUE)
    dollar(v, prefix = "$", big.mark = ".", decimal.mark = ",")
  })
  
  output$txt_iva_compras <- renderText({
    df_c <- compras_filtradas()
    c <- sum(df_c$total - df_c$neto_gravado, na.rm = TRUE)
    dollar(c, prefix = "$", big.mark = ".", decimal.mark = ",")
  })
  
  output$txt_ops <- renderText({
    nrow(ventas_filtradas())
  })
  
  # Gráfico 1: Línea temporal de recaudación
  output$grafico_lineas <- renderPlot({
    req(nrow(ventas_filtradas()) > 0)
    ventas_filtradas() %>%
      group_by(fecha, unidad_negocio) %>%
      summarise(v_diaria = sum(neto_gravado), .groups = "drop") %>%
      ggplot(aes(x = fecha, y = v_diaria, color = unidad_negocio, group = unidad_negocio)) +
      geom_line(size = 1.1) +
      geom_point(size = 2) +
      theme_minimal() +
      labs(x = "Fecha", y = "Monto Neto ($)", color = "Unidad de Negocio") +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
  })
  
  # Gráfico 2: Composición interna de ingresos
  output$grafico_torta <- renderPlot({
    req(nrow(ventas_filtradas()) > 0)
    ventas_filtradas() %>%
      group_by(unidad_negocio) %>%
      summarise(total_neto = sum(neto_gravado)) %>%
      ggplot(aes(x = reorder(unidad_negocio, total_neto), y = total_neto, fill = unidad_negocio)) +
      geom_col(show.legend = FALSE) +
      coord_flip() +
      theme_minimal() +
      labs(x = NULL, y = "Ventas Netas Consolidadas ($)") +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
  })

  # Gráfico 3: Estructura de Costos por Proveedor
  
  output$grafico_proveedores <- renderPlot({
    req(nrow(compras_filtradas()) > 0)
    compras_filtradas() %>%
      group_by(proveedor) %>%
      summarise(gasto = sum(neto_gravado)) %>%
      slice_max(gasto, n = 5, with_ties = FALSE) %>%
      ggplot(aes(x = reorder(proveedor, gasto), y = gasto)) +
      geom_col(fill = "#e74c3c") +
      coord_flip() +
      theme_minimal() +
      labs(x = NULL, y = "Monto de Compra Neto ($)") +
      scale_y_continuous(labels = dollar_format(prefix = "$", big.mark = ".", decimal.mark = ","))
  })
}  
  

# 5. LANZAMIENTO DE LA APLICACIÓN ----------------------------------------------
  
  shinyApp(ui = ui, server = server)
