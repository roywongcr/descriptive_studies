# ============================================================
# Shiny App: Tablas descriptivas + Parámetros + Gemini (ellmer)
# + Gráficos ggplot2 (barras) publicables + integración Word + Gemini
# + Reacomodo UI: sub-tabs "Tabla" / "Gráfico" dentro de cada TAB 1–5
#
# AJUSTES solicitados:
# 1) Título de leyenda editable por usuario; por default = variable "by" seleccionada (label_map)
# 2) Color de barras por selector de colores (no hex manual) + paletas para grupos (SIN error de hcl.colors)
# 3) Eliminar líneas del fondo del gráfico (sin grid)
# 4) En pestaña Parámetros > Formato e idioma:
#    - Número de decimales a la izquierda
#    - Idioma a la derecha
# 5) En sección Estimaciones:
#    - Variables categóricas a la izquierda
#    - Variables cuantitativas a la derecha
# 6) En cada TAB:
#    - Crear sección "Datos del cuadro/gráfico"
#    - Ordenar:
#         * Título del cuadro
#         * Nombre de la columna matriz
#         * Nombre encabezado de grupo
#         * Fuente
#      en dos columnas y dos filas
# 7) En la sección Gráfico:
#    - Apartado "Ejes"
#    - Apartado "Título y leyenda"
#    - Apartado "Apariencia"
#    - Orden amigable para usuario no experto
# 8) Ajuste adicional en "Ejes":
#    - Eje X izquierda | Eje Y derecha
#    - Nombre eje Y izquierda | Nombre eje X derecha
#    - Tamaño base izquierda | Ángulo etiquetas derecha
#    - Ancho nombre etiquetas izquierda | Fuente derecha
# 9) Ajuste adicional visual:
#    - Subtítulos internos para todas las secciones
#
# FIXES IMPORTANTES:
# - build_tab se define UNA sola vez
# - Se ejecuta lapply(1:5, build_tab) para que aparezcan los selectores
# - Word incluye texto + tabla + figura (si existe)
# ============================================================

setwd("C:/R_ANALYSIS/Shiny_Analytics")

library(readxl)
library(tidyverse)
library(shiny)
library(gtsummary)
library(gt)
library(janitor)
library(DT)
library(ellmer)

library(officer)
library(flextable)

library(ggplot2)
library(scales)

# =========================
# Helpers UI
# =========================
section_box <- function(title, ...) {
    tags$div(
        style = "background:#f8f9fa; border:1px solid #e9ecef; border-radius:8px; padding:14px; margin-bottom:15px;",
        h4(style = "margin-top:0; margin-bottom:12px;", title),
        ...
    )
}

subsection_label <- function(txt) {
    tags$div(
        style = "font-weight:600; color:#495057; font-size:0.98em; margin-top:4px; margin-bottom:8px;",
        txt
    )
}

subsection_note <- function(txt) {
    tags$div(
        style = "color:#6c757d; font-size:0.92em; margin-top:-4px; margin-bottom:8px;",
        txt
    )
}

# =========================
# UI
# =========================

tab_parametros <- tabPanel(
    "Parámetros",
    
    section_box(
        "Renombrar variables",
        p("Renombre variables para su visualización. Si no edita un nombre, se mantiene el del archivo."),
        DTOutput("rename_table"),
        tags$br(),
        actionButton("reset_names", "Restablecer nombres del archivo")
    ),
    
    section_box(
        "Formato e idioma",
        subsection_label("Configuración general"),
        fluidRow(
            column(
                width = 6,
                numericInput("decimals", "Número de Decimales:", value = 1, min = 0, max = 4, step = 1)
            ),
            column(
                width = 6,
                selectInput("lang", "Idioma:", choices = c("Español" = "es", "Inglés" = "en"), selected = "es")
            )
        )
    ),
    
    section_box(
        "Estimaciones",
        fluidRow(
            column(
                width = 6,
                subsection_label("Variables categóricas"),
                subsection_note("(PE: Sexo, Provincia)"),
                checkboxInput("cat_ci95", "Incluir intervalo de confianza (IC) al 95%", value = FALSE)
            ),
            column(
                width = 6,
                subsection_label("Variables cuantitativas"),
                subsection_note("(PE: Edad, IMC, Presión arterial)"),
                radioButtons(
                    "q_desc", "Medida descriptiva:",
                    choices = c("Media (promedio)" = "mean", "Mediana" = "median"),
                    selected = "mean", inline = TRUE
                ),
                radioButtons(
                    "q_disp", "Medida de dispersión:",
                    choices = c("Desviación Estándar" = "sd",
                                "Rango" = "range",
                                "Intervalo Intercuartílico" = "iqr",
                                "IC95%" = "ci95"),
                    selected = "sd", inline = FALSE
                ),
                tags$div(
                    style="color:#6c757d; font-size:0.95em; margin-top:6px;",
                    "Nota: Si selecciona IC95% como dispersión, se presentará una columna adicional de IC95% para variables cuantitativas."
                )
            )
        )
    )
)

tab_list_1_5 <- lapply(1:5, function(i) {
    local({
        ii <- i
        tabPanel(
            paste0("TAB ", ii),
            
            uiOutput(paste0("col_selector_", ii)),
            uiOutput(paste0("by_selector_", ii)),
            uiOutput(paste0("cmp_selector_", ii)),
            uiOutput(paste0("overall_selector_", ii)),
            tags$hr(),
            
            section_box(
                "Datos del cuadro/gráfico",
                fluidRow(
                    column(
                        width = 6,
                        subsection_label("Identificación"),
                        textInput(paste0("title_", ii), "Título del cuadro:", value = "Título del cuadro")
                    ),
                    column(
                        width = 6,
                        subsection_label("Estructura"),
                        textInput(paste0("matrix_col_", ii), "Nombre de la Columna matriz:", value = "Characteristics")
                    )
                ),
                fluidRow(
                    column(
                        width = 6,
                        textInput(paste0("group_header_", ii), "Nombre encabezado de grupo:", value = "Grupo")
                    ),
                    column(
                        width = 6,
                        textInput(paste0("source_", ii), "Fuente:", value = "Null")
                    )
                )
            ),
            
            actionButton(paste0("gen_desc_", ii), "Generar descripción (Tabla)"),
            tags$br(), tags$br(),
            uiOutput(paste0("desc_", ii)),
            tags$hr(),
            
            tabsetPanel(
                id = paste0("subtab_", ii),
                
                tabPanel("Tabla",
                         gt::gt_output(paste0("gt_table_", ii))
                ),
                
                tabPanel("Gráfico",
                         tags$div(
                             style="max-width:1100px;",
                             uiOutput(paste0("plot_ui_", ii)),
                             tags$br(),
                             plotOutput(paste0("plot_", ii), height = "420px"),
                             tags$div(
                                 style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;",
                                 downloadButton(paste0("download_plot_", ii), "Descargar PNG (300 dpi)"),
                                 actionButton(paste0("gen_plot_desc_", ii), "Generar descripción (Figura)")
                             ),
                             tags$br(),
                             uiOutput(paste0("plot_desc_", ii))
                         )
                )
            )
        )
    })
})

tab_compilar <- tabPanel(
    "Compilar Resultados",
    section_box(
        "Compilar resultados (TAB 1 a TAB 5)",
        p("Visualice o descargue en Word todos los resultados (texto + tabla + figura) en orden de TAB."),
        actionButton("view_compiled", "Visualizador"),
        tags$span(style="margin-left:10px;"),
        downloadButton("download_word", "Descarga en Word"),
        tags$hr(),
        tags$div(
            style="color:#6c757d;",
            "Nota: Se compilan los TABs con la última tabla guardada y la figura si existe."
        )
    )
)

tab_metodos <- tabPanel(
    "Métodos",
    section_box(
        "Métodos de análisis",
        p("Presione el botón para generar la sección de Métodos basada en los parámetros y tablas configuradas en las TABs de resultados."),
        actionButton("gen_methods", "Generar Métodos"),
        tags$br(), tags$br(),
        uiOutput("methods_text")
    )
)

tabs_main <- c(list(tab_parametros), tab_list_1_5, list(tab_compilar), list(tab_metodos))

ui <- fluidPage(
    titlePanel("Análisis Descriptivo Automatizado"),
    sidebarLayout(
        sidebarPanel(
            radioButtons("fileType", "Tipo de archivo:", choices = c(".csv", ".xlsx"), selected = ".csv"),
            fileInput("uploadFile", "Subir base de datos", accept = c(".csv", ".xlsx")),
            tags$hr(),
            helpText("Use la pestaña 'Parámetros' para renombrar variables y ajustar estimaciones.")
        ),
        mainPanel(
            do.call(tabsetPanel, c(list(id = "tabs"), tabs_main))
        )
    )
)

# =========================
# Server
# =========================
server <- function(input, output, session) {
    
    # -------- Helpers generales --------
    get_decimals <- function() {
        d <- suppressWarnings(as.integer(input$decimals))
        if (is.na(d)) d <- 1L
        max(0L, min(4L, d))
    }
    
    get_lang <- function() {
        lang <- input$lang
        if (!lang %in% c("es", "en")) lang <- "es"
        lang
    }
    
    observeEvent(input$lang, {
        lang <- get_lang()
        dec_mark <- if (lang == "es") "," else "."
        tryCatch(
            gtsummary::theme_gtsummary_language(lang, decimal.mark = dec_mark),
            error = function(e) NULL
        )
    }, ignoreInit = FALSE)
    
    # -------- GEMINI --------
    chat <- reactiveVal(NULL)
    
    make_google_chat <- function(key) {
        if (!nzchar(key)) return(NULL)
        ns <- asNamespace("ellmer")
        if (exists("chat_google_gemini", where = ns, inherits = FALSE)) {
            return(ellmer::chat_google_gemini(model = "gemini-2.5-flash", api_key = key, echo = "none"))
        }
        if (exists("chat_google", where = ns, inherits = FALSE)) {
            return(ellmer::chat_google(model = "gemini-2.5-flash", api_key = key))
        }
        stop("Tu versión de 'ellmer' no soporta Gemini. Actualiza con install.packages('ellmer').")
    }
    
    observeEvent(input$uploadFile, {
        key <- Sys.getenv("GOOGLE_API_KEY")
        if (!nzchar(key)) {
            chat(NULL)
        } else {
            chat(tryCatch(make_google_chat(key), error = function(e) NULL))
        }
    }, ignoreInit = TRUE)
    
    gemini_chat_text <- function(prompt) {
        ch <- chat()
        if (is.null(ch)) {
            return("Gemini no está disponible. Verifique GOOGLE_API_KEY y vuelva a cargar el archivo.")
        }
        resp <- tryCatch(ch$chat(prompt), error = function(e) paste0("Error Gemini: ", e$message))
        if (is.null(resp)) return("")
        if (is.character(resp)) return(paste(resp, collapse = "\n"))
        if (is.list(resp)) return(paste(unlist(resp), collapse = "\n"))
        as.character(resp)
    }
    
    make_table_text <- function(tbl_gts) {
        tb <- tryCatch(
            gtsummary::as_tibble(tbl_gts, col_labels = TRUE),
            error = function(e) gtsummary::as_tibble(tbl_gts)
        )
        tb <- tb %>% mutate(across(everything(), as.character))
        paste(capture.output(print(tb, n = Inf, width = Inf)), collapse = "\n")
    }
    
    # -------- DATASET + RENOMBRADO --------
    label_map <- reactiveVal(NULL)
    ds_cache  <- reactiveVal(NULL)
    
    observeEvent(input$uploadFile, {
        req(input$uploadFile)
        
        df_raw <- switch(
            input$fileType,
            ".csv"  = readr::read_csv(input$uploadFile$datapath, show_col_types = FALSE),
            ".xlsx" = readxl::read_excel(input$uploadFile$datapath)
        ) %>% tibble::as_tibble()
        
        orig_names <- names(df_raw)
        internal_names <- make.unique(janitor::make_clean_names(orig_names), sep = "_")
        
        df <- df_raw
        names(df) <- internal_names
        
        ds_cache(list(df = df, orig_names = orig_names, internal_names = internal_names))
        label_map(setNames(orig_names, internal_names))
    }, ignoreInit = FALSE)
    
    observeEvent(input$reset_names, {
        req(ds_cache())
        ds <- ds_cache()
        label_map(setNames(ds$orig_names, ds$internal_names))
    })
    
    output$rename_table <- renderDT({
        req(ds_cache(), label_map())
        ds <- ds_cache()
        lm <- label_map()
        
        df_names <- tibble::tibble(
            variable_en_archivo = ds$orig_names,
            nombre_mostrado     = unname(lm[ds$internal_names]),
            .internal_name      = ds$internal_names
        )
        
        DT::datatable(
            df_names,
            rownames = FALSE,
            options = list(
                pageLength = 10,
                autoWidth = TRUE,
                columnDefs = list(list(targets = 2, visible = FALSE))
            ),
            editable = list(target = "cell", disable = list(columns = c(0, 2)))
        )
    })
    
    observeEvent(input$rename_table_cell_edit, {
        req(ds_cache(), label_map())
        info <- input$rename_table_cell_edit
        ds <- ds_cache()
        lm <- label_map()
        if (info$col != 1) return(NULL)
        internal <- ds$internal_names[info$row]
        new_label <- as.character(info$value)
        if (!nzchar(trimws(new_label))) new_label <- ds$orig_names[info$row]
        lm[internal] <- new_label
        label_map(lm)
    })
    
    # -------- Construcción de tablas --------
    build_tbl_summary <- function(df_use, byvar = NULL, label_list = NULL,
                                  do_compare = FALSE, include_overall = TRUE,
                                  group_header = "Grupo") {
        
        dec  <- get_decimals()
        lang <- get_lang()
        
        df_use <- df_use %>%
            mutate(across(where(~ is.character(.x) || is.logical(.x) || is.factor(.x)), as.factor))
        
        desc_choice <- input$q_desc
        disp_choice <- input$q_disp
        
        cont_stat <- {
            if (disp_choice == "ci95") {
                if (desc_choice == "mean") "{mean}" else "{median}"
            } else if (desc_choice == "mean" && disp_choice == "sd") {
                "{mean} ({sd})"
            } else if (desc_choice == "mean" && disp_choice == "range") {
                "{mean} ({min}, {max})"
            } else if (desc_choice == "mean" && disp_choice == "iqr") {
                "{mean} ({p25}, {p75})"
            } else if (desc_choice == "median" && disp_choice == "sd") {
                "{median} ({sd})"
            } else if (desc_choice == "median" && disp_choice == "range") {
                "{median} ({min}, {max})"
            } else if (desc_choice == "median" && disp_choice == "iqr") {
                "{median} ({p25}, {p75})"
            } else {
                "{mean} ({sd})"
            }
        }
        
        stat_list <- list(
            all_continuous()  ~ cont_stat,
            all_categorical() ~ "{n} ({p}%)"
        )
        digits_list <- list(
            all_continuous()  ~ dec,
            all_categorical() ~ dec
        )
        
        has_by <- !is.null(byvar) && nzchar(byvar)
        
        if (has_by) {
            tbl <- tbl_summary(
                data = df_use,
                by = byvar,
                statistic = stat_list,
                digits = digits_list,
                label = label_list,
                missing = "ifany"
            ) %>% add_n()
            
            if (isTRUE(include_overall)) {
                tbl <- tbl %>%
                    add_overall() %>%
                    modify_header(stat_0 = "Total")
            }
            
            if (isTRUE(do_compare)) {
                tbl <- tbl %>% add_p(pvalue_fun = function(x) gtsummary::style_pvalue(x, digits = dec))
            }
            
            if (!is.null(group_header) && nzchar(trimws(group_header))) {
                tbl <- tbl %>% modify_spanning_header(matches("^stat_[1-9]") ~ group_header)
            }
            
        } else {
            tbl <- tbl_summary(
                data = df_use,
                statistic = stat_list,
                digits = digits_list,
                label = label_list,
                missing = "ifany"
            ) %>% add_n()
        }
        
        ci_include <- NULL
        if (isTRUE(input$cat_ci95)) ci_include <- c(ci_include, all_categorical())
        if (identical(disp_choice, "ci95")) ci_include <- c(ci_include, all_continuous())
        
        if (!is.null(ci_include)) {
            tbl <- tbl %>%
                add_ci(include = ci_include) %>%
                modify_header(ci = if (lang == "es") "IC95%" else "95% CI")
        }
        
        if (!isTRUE(do_compare) && "p.value" %in% names(tbl$table_body)) {
            tbl$table_body <- tbl$table_body %>% select(-p.value)
            if (!is.null(tbl$table_styling$header)) {
                tbl$table_styling$header <- tbl$table_styling$header %>% filter(column != "p.value")
            }
        }
        
        if ("p.value" %in% names(tbl$table_body)) {
            tbl <- tbl %>% modify_header(p.value = if (lang == "es") "Valor p" else "p-value")
        }
        
        tbl %>% bold_labels()
    }
    
    # -------- Stores --------
    desc_vals <- lapply(1:5, function(i) reactiveVal(""))
    tbl_store  <- lapply(1:5, function(i) reactiveVal(NULL))
    
    plot_store <- lapply(1:5, function(i) reactiveVal(NULL))
    plot_data_store <- lapply(1:5, function(i) reactiveVal(NULL))
    plot_desc_vals <- lapply(1:5, function(i) reactiveVal(""))
    
    # -------- GT con meta --------
    make_gt_with_meta <- function(i, tbl_gts) {
        lang <- get_lang()
        
        title_txt <- input[[paste0("title_", i)]]
        if (is.null(title_txt) || !nzchar(trimws(title_txt))) title_txt <- "Título del cuadro"
        
        source_txt <- input[[paste0("source_", i)]]
        if (is.null(source_txt) || !nzchar(trimws(source_txt))) source_txt <- "Null"
        
        prefix_source <- if (lang == "es") "Fuente: " else "Source: "
        
        gtsummary::as_gt(tbl_gts) %>%
            gt::tab_header(title = gt::md(title_txt)) %>%
            gt::tab_source_note(source_note = gt::md(paste0(prefix_source, source_txt)))
    }
    
    # ==========================================================
    # GRÁFICOS: helpers
    # ==========================================================
    `%||%` <- function(a, b) if (!is.null(a) && nzchar(as.character(a))) a else b
    
    wrap_text <- function(x, width = 30) {
        x <- as.character(x)
        vapply(x, function(s) paste(strwrap(s, width = width), collapse = "\n"), character(1))
    }
    
    get_palette <- function(name, n) {
        name <- as.character(name)
        n <- as.integer(n)
        if (is.na(n) || n <= 1) return("#2C7FB8")
        
        okabe_ito <- c(
            "#E69F00", "#56B4E9", "#009E73", "#F0E442",
            "#0072B2", "#D55E00", "#CC79A7", "#000000"
        )
        
        tableau10 <- c(
            "#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F",
            "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC"
        )
        
        expand_pal <- function(pal, n) {
            if (n <= length(pal)) return(pal[seq_len(n)])
            grDevices::colorRampPalette(pal)(n)
        }
        
        if (name %in% c("Okabe-Ito", "okabe-ito", "OkabeIto")) {
            return(expand_pal(okabe_ito, n))
        }
        
        if (name %in% c("Tableau 10", "Tableau10", "tableau10")) {
            return(expand_pal(tableau10, n))
        }
        
        if (name %in% c("Set1", "Set2")) {
            if (requireNamespace("RColorBrewer", quietly = TRUE)) {
                max_n <- if (name == "Set1") 9 else 8
                base <- RColorBrewer::brewer.pal(min(n, max_n), name)
                return(expand_pal(base, n))
            } else {
                base <- if (name == "Set1") tableau10 else okabe_ito
                return(expand_pal(base, n))
            }
        }
        
        if (name %in% c("Viridis", "viridis")) {
            if (requireNamespace("viridisLite", quietly = TRUE)) {
                return(viridisLite::viridis(n))
            }
        }
        
        return(grDevices::hcl.colors(n, palette = "Zissou 1"))
    }
    
    make_bar_data <- function(df, x, fill = NULL,
                              y_mode = c("count","percent"),
                              percent_den = c("overall","within_fill"),
                              top_n = NULL, include_others = TRUE,
                              others_label = "Otros") {
        y_mode <- match.arg(y_mode)
        percent_den <- match.arg(percent_den)
        
        df <- df %>%
            mutate(
                .x = as.factor(.data[[x]]),
                .fill = if (!is.null(fill) && nzchar(fill)) as.factor(.data[[fill]]) else factor("Total")
            ) %>%
            filter(!is.na(.x), .x != "")
        
        out <- df %>% count(.x, .fill, name = "n")
        
        if (!is.null(top_n) && is.finite(top_n) && top_n > 0) {
            top_n <- as.integer(top_n)
            
            totals_x <- out %>%
                group_by(.x) %>% summarise(n_total = sum(n), .groups = "drop") %>%
                arrange(desc(n_total))
            
            keep_levels <- head(totals_x$.x, top_n)
            
            out <- out %>%
                mutate(.x2 = ifelse(.x %in% keep_levels, as.character(.x), others_label)) %>%
                mutate(.x2 = factor(.x2, levels = c(as.character(keep_levels),
                                                    if (include_others) others_label else NULL)))
            
            if (!include_others) {
                out <- out %>% filter(.x %in% keep_levels) %>% mutate(.x2 = factor(as.character(.x)))
            }
            
            out <- out %>%
                group_by(.x2, .fill) %>% summarise(n = sum(n), .groups = "drop") %>%
                rename(.x = .x2)
        }
        
        if (y_mode == "percent") {
            if (percent_den == "within_fill") {
                out <- out %>% group_by(.fill) %>% mutate(p = n / sum(n)) %>% ungroup()
            } else {
                out <- out %>% mutate(p = n / sum(n))
            }
        } else {
            out <- out %>% mutate(p = NA_real_)
        }
        
        out
    }
    
    reorder_x_levels <- function(d, y_mode = c("count","percent"), decreasing = TRUE) {
        y_mode <- match.arg(y_mode)
        ord <- d %>%
            group_by(.x) %>%
            summarise(val = sum(if (y_mode == "percent") p else n, na.rm = TRUE), .groups = "drop") %>%
            arrange(if (decreasing) desc(val) else val)
        d$.x <- factor(d$.x, levels = as.character(ord$.x))
        d
    }
    
    make_bar_plot <- function(d, y_mode = c("count","percent"),
                              orientation = c("vertical","horizontal"),
                              title = NULL, xlab = NULL, ylab = NULL,
                              font_family = "sans", base_size = 12,
                              single_color = "#2C7FB8",
                              fill_by_groups = FALSE,
                              palette_name = "Set1",
                              legend_title = NULL,
                              x_text_angle = 0, x_wrap = 0,
                              show_labels = FALSE, label_size = 3.5) {
        
        y_mode <- match.arg(y_mode)
        orientation <- match.arg(orientation)
        
        has_groups <- fill_by_groups && nlevels(d$.fill) > 1
        
        d2 <- d
        if (x_wrap > 0) {
            d2$.x <- factor(
                wrap_text(d2$.x, width = x_wrap),
                levels = wrap_text(levels(d2$.x), width = x_wrap)
            )
        }
        
        p <- ggplot(d2, aes(
            x = .x,
            y = if (y_mode == "percent") p else n,
            fill = if (has_groups) .fill else NULL
        )) +
            geom_col(width = 0.78, color = NA) +
            theme_minimal(base_family = font_family, base_size = base_size) +
            theme(
                panel.grid = element_blank(),
                plot.title = element_text(face = "bold"),
                axis.title = element_text(face = "bold"),
                axis.text.x = element_text(angle = x_text_angle, hjust = ifelse(x_text_angle == 0, 0.5, 1))
            ) +
            labs(title = title, x = xlab, y = ylab)
        
        if (!has_groups) {
            p <- p + scale_fill_manual(values = single_color, guide = "none")
        } else {
            pal <- get_palette(palette_name, nlevels(d2$.fill))
            p <- p + scale_fill_manual(values = pal, name = legend_title %||% "")
        }
        
        if (y_mode == "percent") {
            p <- p + scale_y_continuous(labels = scales::percent_format(accuracy = 1))
        }
        
        if (isTRUE(show_labels)) {
            lab <- if (y_mode == "percent") scales::percent(d2$p, accuracy = 0.1) else as.character(d2$n)
            p <- p + geom_text(aes(label = lab), vjust = -0.25, size = label_size, show.legend = FALSE)
        }
        
        if (orientation == "horizontal") {
            p <- p + coord_flip() +
                theme(axis.text.x = element_text(angle = 0, hjust = 0.5))
        }
        
        p
    }
    
    make_plot_text_for_gemini <- function(i) {
        obj <- plot_data_store[[i]]()
        if (is.null(obj)) return(NULL)
        
        s <- obj$settings
        d <- obj$data
        tbl_txt <- paste(capture.output(print(d, n = Inf, width = Inf)), collapse = "\n")
        
        paste0(
            "CONFIGURACIÓN DEL GRÁFICO:\n",
            "- Eje X: ", s$x, "\n",
            if (!is.null(s$fill)) paste0("- Grupos (relleno/leyenda): ", s$fill, "\n") else "",
            if (!is.null(s$legend_title)) paste0("- Título leyenda: ", s$legend_title, "\n") else "",
            "- Métrica eje Y: ", if (s$y_mode == "percent") "Porcentaje" else "Frecuencia", "\n",
            if (!is.null(s$percent_den)) paste0("- Denominador %: ", s$percent_den, "\n") else "",
            "- Top N: ", s$top_n, "\n",
            "- Incluir 'Otros': ", s$include_others, "\n",
            "- Orden: ", s$order_mode, "\n",
            "- Orientación: ", s$orientation, "\n",
            "- Título: ", s$title, "\n",
            "- X label: ", s$xlab, "\n",
            "- Y label: ", s$ylab, "\n",
            "- Fuente: ", s$font, " | Tamaño base: ", s$size, "\n\n",
            "DATOS RESUMIDOS (n y %):\n",
            tbl_txt
        )
    }
    
    # ==========================================================
    # TABs 1–5
    # ==========================================================
    tab_tbls <- vector("list", 5)
    
    build_tab <- function(i) {
        
        output[[paste0("col_selector_", i)]] <- renderUI({
            req(ds_cache(), label_map())
            ds <- ds_cache()
            lm <- label_map()
            choices_named <- setNames(ds$internal_names, unname(lm[ds$internal_names]))
            
            selectizeInput(
                inputId = paste0("cols_", i),
                label   = paste("Columnas a describir (TAB", i, "):"),
                choices = choices_named,
                selected = ds$internal_names,
                multiple = TRUE,
                options = list(
                    plugins = list("remove_button"),
                    placeholder = "Seleccione una o más columnas…",
                    maxItems = length(ds$internal_names)
                )
            )
        })
        
        output[[paste0("by_selector_", i)]] <- renderUI({
            req(ds_cache(), label_map())
            ds <- ds_cache()
            lm <- label_map()
            choices_named <- setNames(ds$internal_names, unname(lm[ds$internal_names]))
            
            selectInput(
                inputId = paste0("by_", i),
                label   = "Variable de agrupación (by, opcional):",
                choices = c("— Sin 'by' —" = "", choices_named),
                selected = ""
            )
        })
        
        output[[paste0("cmp_selector_", i)]] <- renderUI({
            byv <- input[[paste0("by_", i)]]
            if (is.null(byv) || !nzchar(byv)) return(NULL)
            checkboxInput(
                paste0("do_compare_", i),
                "Comparación con prueba estadística (valor p)",
                value = TRUE
            )
        })
        
        output[[paste0("overall_selector_", i)]] <- renderUI({
            byv <- input[[paste0("by_", i)]]
            if (is.null(byv) || !nzchar(byv)) return(NULL)
            checkboxInput(
                paste0("include_overall_", i),
                "Incluir totales",
                value = TRUE
            )
        })
        
        tab_tbls[[i]] <- reactive({
            req(ds_cache(), label_map())
            ds <- ds_cache()
            df_full <- ds$df
            lm <- label_map()
            
            sel <- input[[paste0("cols_", i)]]
            validate(need(!is.null(sel) && length(sel) > 0, "Seleccione al menos una columna."))
            
            sel <- sel[sel %in% names(df_full)]
            validate(need(length(sel) > 0, "Las columnas seleccionadas no existen en el dataset."))
            
            byvar <- input[[paste0("by_", i)]]
            has_by <- !is.null(byvar) && nzchar(byvar) && byvar %in% names(df_full)
            
            do_compare <- if (has_by) isTRUE(input[[paste0("do_compare_", i)]]) else FALSE
            include_overall <- if (has_by) isTRUE(input[[paste0("include_overall_", i)]]) else FALSE
            
            cols_to_use <- unique(c(sel, if (has_by) byvar))
            df_use <- df_full[, cols_to_use, drop = FALSE]
            
            label_list <- as.list(unname(lm[names(df_use)]))
            names(label_list) <- names(df_use)
            
            group_header <- input[[paste0("group_header_", i)]]
            if (is.null(group_header) || !nzchar(trimws(group_header))) {
                group_header <- if (get_lang() == "es") "Grupo" else "Group"
            }
            
            tbl <- build_tbl_summary(
                df_use,
                byvar = if (has_by) byvar else NULL,
                label_list = label_list,
                do_compare = do_compare,
                include_overall = include_overall,
                group_header = group_header
            )
            
            matrix_name <- input[[paste0("matrix_col_", i)]]
            if (is.null(matrix_name) || !nzchar(trimws(matrix_name))) matrix_name <- "Characteristics"
            
            tbl %>% modify_header(label = matrix_name)
        })
        
        observe({
            req(ds_cache(), label_map())
            tbl_try <- tryCatch(tab_tbls[[i]](), error = function(e) NULL)
            if (!is.null(tbl_try)) tbl_store[[i]](tbl_try)
        })
        
        output[[paste0("gt_table_", i)]] <- gt::render_gt({
            req(tab_tbls[[i]]())
            make_gt_with_meta(i, tab_tbls[[i]]())
        })
        
        output[[paste0("desc_", i)]] <- renderUI({
            txt <- desc_vals[[i]]()
            if (!nzchar(trimws(txt))) return(NULL)
            tags$div(
                style = "background:#f8f9fa; border:1px solid #e9ecef; padding:12px; border-radius:8px;",
                tags$b("Descripción (Gemini):"),
                tags$br(), tags$br(),
                tags$div(style = "white-space:pre-wrap;", txt)
            )
        })
        
        observeEvent(input[[paste0("gen_desc_", i)]], {
            tbl_now <- tryCatch(tab_tbls[[i]](), error = function(e) NULL)
            if (is.null(tbl_now)) tbl_now <- tbl_store[[i]]()
            req(tbl_now)
            
            desc_vals[[i]]("Generando descripción...")
            tabla_texto <- make_table_text(tbl_now)
            
            prompt <- paste0(
                "Actúe como epidemiólogo y estadístico. ",
                "Redacte un párrafo profesional y claro basado en la siguiente tabla de resultados descriptivos. ",
                "Si hay comparación por grupos, mencione diferencias relevantes y valores p si aparecen. ",
                "No incluya títulos ni introducción.\n\n",
                "TABLA:\n", tabla_texto
            )
            desc_vals[[i]](gemini_chat_text(prompt))
        }, ignoreInit = TRUE)
        
        # ==========================================================
        # UI gráfico reorganizada
        # ==========================================================
        output[[paste0("plot_ui_", i)]] <- renderUI({
            req(ds_cache(), label_map())
            ds <- ds_cache()
            lm <- label_map()
            
            sel <- input[[paste0("cols_", i)]]
            if (is.null(sel) || length(sel) == 0) {
                return(tags$div(style="color:#6c757d;", "Primero seleccione columnas a describir en este TAB."))
            }
            
            choices_named <- setNames(sel, unname(lm[sel]))
            byv <- input[[paste0("by_", i)]]
            has_by <- !is.null(byv) && nzchar(byv)
            
            color_choices <- c(
                "Azul" = "#2C7FB8",
                "Rojo" = "#D73027",
                "Verde" = "#1A9850",
                "Naranja" = "#F46D43",
                "Morado" = "#7B3294",
                "Gris" = "#666666",
                "Negro" = "#000000"
            )
            
            palette_choices <- c("Set1", "Set2", "Okabe-Ito", "Tableau 10", "Viridis")
            by_label_default <- if (has_by) unname(lm[byv]) else ""
            
            tagList(
                
                section_box(
                    "Ejes",
                    subsection_label("Selección de ejes"),
                    fluidRow(
                        column(
                            width = 6,
                            selectInput(
                                paste0("plot_x_", i),
                                "Eje X (categoría):",
                                choices = choices_named,
                                selected = sel[1]
                            )
                        ),
                        column(
                            width = 6,
                            radioButtons(
                                paste0("plot_y_", i),
                                "Eje Y:",
                                choices = c("Frecuencias (n)" = "count", "Porcentajes (%)" = "percent"),
                                selected = "count",
                                inline = TRUE
                            ),
                            conditionalPanel(
                                condition = sprintf("input.%s == 'percent'", paste0("plot_y_", i)),
                                radioButtons(
                                    paste0("plot_pct_den_", i),
                                    "Denominador de %:",
                                    choices = c("Total global" = "overall",
                                                "Dentro de cada grupo (by)" = "within_fill"),
                                    selected = "overall",
                                    inline = TRUE
                                )
                            )
                        )
                    ),
                    
                    subsection_label("Etiquetas de ejes"),
                    fluidRow(
                        column(
                            width = 6,
                            textInput(paste0("plot_ylab_", i), "Nombre del eje Y:", value = "")
                        ),
                        column(
                            width = 6,
                            textInput(paste0("plot_xlab_", i), "Nombre del eje X:", value = "")
                        )
                    ),
                    
                    subsection_label("Formato de texto"),
                    fluidRow(
                        column(
                            width = 6,
                            numericInput(paste0("plot_size_", i), "Tamaño base:", value = 12, min = 8, max = 24)
                        ),
                        column(
                            width = 6,
                            sliderInput(paste0("plot_angle_", i), "Ángulo de etiquetas:", min = 0, max = 90, value = 0, step = 5)
                        )
                    ),
                    fluidRow(
                        column(
                            width = 6,
                            numericInput(paste0("plot_wrap_", i), "Ancho del nombre de etiquetas:", value = 0, min = 0, max = 80, step = 5)
                        ),
                        column(
                            width = 6,
                            textInput(paste0("plot_font_", i), "Fuente:", value = "sans")
                        )
                    )
                ),
                
                section_box(
                    "Título y leyenda",
                    fluidRow(
                        column(
                            width = 6,
                            subsection_label("Título"),
                            textInput(paste0("plot_title_", i), "Título del gráfico:", value = ""),
                            if (has_by) checkboxInput(
                                paste0("plot_fill_by_", i),
                                "Usar grupos (by) como leyenda y color",
                                value = TRUE
                            )
                        ),
                        column(
                            width = 6,
                            subsection_label("Leyenda"),
                            if (has_by) textInput(
                                paste0("plot_legend_title_", i),
                                "Nombre de la leyenda:",
                                value = by_label_default
                            )
                        )
                    )
                ),
                
                section_box(
                    "Apariencia",
                    fluidRow(
                        column(
                            width = 6,
                            subsection_label("Organización"),
                            selectInput(
                                paste0("plot_order_", i),
                                "Orden de las barras:",
                                choices = c("Ninguno"="none",
                                            "Descendente"="desc",
                                            "Ascendente"="asc"),
                                selected = "desc"
                            ),
                            numericInput(
                                paste0("plot_topn_", i),
                                "Número máximo de categorías (Top N):",
                                value = 10, min = 1, max = 200, step = 1
                            ),
                            checkboxInput(
                                paste0("plot_others_", i),
                                "Agrupar restantes como 'Otros'",
                                value = TRUE
                            ),
                            radioButtons(
                                paste0("plot_orient_", i),
                                "Orientación:",
                                choices = c("Vertical" = "vertical", "Horizontal" = "horizontal"),
                                selected = "vertical",
                                inline = TRUE
                            )
                        ),
                        column(
                            width = 6,
                            subsection_label("Colores"),
                            conditionalPanel(
                                condition = sprintf("input.%s", paste0("plot_fill_by_", i)),
                                selectInput(
                                    paste0("plot_palette_", i),
                                    "Paleta para grupos:",
                                    choices = palette_choices,
                                    selected = "Set1"
                                )
                            ),
                            conditionalPanel(
                                condition = sprintf("!input.%s", paste0("plot_fill_by_", i)),
                                selectInput(
                                    paste0("plot_color_pick_", i),
                                    "Color de barras:",
                                    choices = color_choices,
                                    selected = "#2C7FB8"
                                )
                            ),
                            subsection_label("Etiquetas de valor"),
                            checkboxInput(
                                paste0("plot_labels_", i),
                                "Mostrar etiquetas de valor",
                                value = FALSE
                            ),
                            numericInput(
                                paste0("plot_labelsize_", i),
                                "Tamaño de etiquetas de valor:",
                                value = 3.5, min = 2, max = 8, step = 0.5
                            )
                        )
                    )
                )
            )
        })
        
        observeEvent(input[[paste0("by_", i)]], {
            req(label_map())
            lm <- label_map()
            byv <- input[[paste0("by_", i)]]
            has_by <- !is.null(byv) && nzchar(byv)
            if (has_by) {
                updateTextInput(session, paste0("plot_legend_title_", i), value = unname(lm[byv]))
            }
        }, ignoreInit = TRUE)
        
        plot_reactive_i <- reactive({
            req(ds_cache(), label_map())
            ds <- ds_cache()
            df_full <- ds$df
            lm <- label_map()
            
            x <- input[[paste0("plot_x_", i)]]
            validate(need(!is.null(x) && nzchar(x), "Seleccione una variable para el eje X."))
            
            byv <- input[[paste0("by_", i)]]
            has_by <- !is.null(byv) && nzchar(byv) && byv %in% names(df_full)
            
            fill_by <- isTRUE(input[[paste0("plot_fill_by_", i)]]) && has_by
            fill_var <- if (fill_by) byv else NULL
            
            legend_title <- NULL
            if (fill_by) {
                lt <- input[[paste0("plot_legend_title_", i)]]
                legend_title <- if (!is.null(lt) && nzchar(trimws(lt))) lt else unname(lm[byv])
            }
            
            y_mode <- input[[paste0("plot_y_", i)]]
            if (is.null(y_mode)) y_mode <- "count"
            
            pct_den <- input[[paste0("plot_pct_den_", i)]]
            if (is.null(pct_den) || !nzchar(pct_den)) pct_den <- "overall"
            
            orient <- input[[paste0("plot_orient_", i)]]
            if (is.null(orient)) orient <- "vertical"
            
            title <- input[[paste0("plot_title_", i)]]
            xlab  <- input[[paste0("plot_xlab_", i)]]
            ylab  <- input[[paste0("plot_ylab_", i)]]
            font  <- input[[paste0("plot_font_", i)]]
            size  <- input[[paste0("plot_size_", i)]]
            
            top_n <- input[[paste0("plot_topn_", i)]]
            include_others <- isTRUE(input[[paste0("plot_others_", i)]])
            order_mode <- input[[paste0("plot_order_", i)]]
            
            x_angle <- input[[paste0("plot_angle_", i)]]
            x_wrap  <- input[[paste0("plot_wrap_", i)]]
            show_lbls <- isTRUE(input[[paste0("plot_labels_", i)]])
            lbl_size <- input[[paste0("plot_labelsize_", i)]]
            
            single_color <- input[[paste0("plot_color_pick_", i)]]
            if (is.null(single_color) || !nzchar(single_color)) single_color <- "#2C7FB8"
            palette_name <- input[[paste0("plot_palette_", i)]]
            if (is.null(palette_name) || !nzchar(palette_name)) palette_name <- "Set1"
            
            if (!nzchar(trimws(xlab))) xlab <- unname(lm[x])
            if (!nzchar(trimws(ylab))) ylab <- if (y_mode == "percent") "Porcentaje" else "Frecuencia"
            
            cols_needed <- unique(c(x, if (!is.null(fill_var)) fill_var))
            d0 <- df_full[, cols_needed, drop = FALSE]
            
            dbar <- make_bar_data(
                df = d0,
                x = x,
                fill = fill_var,
                y_mode = y_mode,
                percent_den = pct_den,
                top_n = top_n,
                include_others = include_others,
                others_label = "Otros"
            )
            
            if (!is.null(order_mode) && order_mode %in% c("desc", "asc")) {
                dbar <- reorder_x_levels(dbar, y_mode = y_mode, decreasing = (order_mode == "desc"))
            }
            
            p <- make_bar_plot(
                d = dbar,
                y_mode = y_mode,
                orientation = orient,
                title = if (nzchar(trimws(title))) title else NULL,
                xlab = xlab,
                ylab = ylab,
                font_family = if (nzchar(trimws(font))) font else "sans",
                base_size = ifelse(is.null(size) || is.na(size), 12, size),
                single_color = single_color,
                fill_by_groups = fill_by,
                palette_name = palette_name,
                legend_title = legend_title,
                x_text_angle = ifelse(is.null(x_angle), 0, x_angle),
                x_wrap = ifelse(is.null(x_wrap), 0, x_wrap),
                show_labels = show_lbls,
                label_size = ifelse(is.null(lbl_size) || is.na(lbl_size), 3.5, lbl_size)
            )
            
            d_for_text <- dbar %>%
                mutate(
                    categoria = as.character(.x),
                    grupo = as.character(.fill),
                    n = n,
                    pct = if (!all(is.na(p))) round(100 * p, 1) else NA_real_
                ) %>%
                select(categoria, grupo, n, pct)
            
            plot_data_store[[i]](list(
                settings = list(
                    x = unname(lm[x]),
                    fill = if (fill_by) unname(lm[fill_var]) else NULL,
                    legend_title = legend_title,
                    y_mode = y_mode,
                    percent_den = if (y_mode == "percent") pct_den else NULL,
                    top_n = top_n,
                    include_others = include_others,
                    order_mode = order_mode,
                    orientation = orient,
                    title = title,
                    xlab = xlab,
                    ylab = ylab,
                    font = font,
                    size = size
                ),
                data = d_for_text
            ))
            
            p
        })
        
        output[[paste0("plot_", i)]] <- renderPlot({
            req(plot_reactive_i())
            p <- plot_reactive_i()
            plot_store[[i]](p)
            p
        })
        
        output[[paste0("download_plot_", i)]] <- downloadHandler(
            filename = function() paste0("TAB", i, "_grafico_", format(Sys.Date(), "%Y%m%d"), ".png"),
            content = function(file) {
                p <- plot_store[[i]]()
                req(p)
                
                if (requireNamespace("ragg", quietly = TRUE)) {
                    ragg::agg_png(filename = file, width = 2400, height = 1600, res = 300)
                    print(p)
                    dev.off()
                } else {
                    png(file, width = 2400, height = 1600, res = 300)
                    print(p)
                    dev.off()
                }
            }
        )
        
        output[[paste0("plot_desc_", i)]] <- renderUI({
            txt <- plot_desc_vals[[i]]()
            if (!nzchar(trimws(txt))) return(NULL)
            tags$div(
                style = "background:#f8f9fa; border:1px solid #e9ecef; padding:12px; border-radius:8px;",
                tags$b("Descripción de la figura (Gemini):"),
                tags$br(), tags$br(),
                tags$div(style = "white-space:pre-wrap;", txt)
            )
        })
        
        observeEvent(input[[paste0("gen_plot_desc_", i)]], {
            txt_plot <- make_plot_text_for_gemini(i)
            req(txt_plot)
            plot_desc_vals[[i]]("Generando descripción...")
            
            prompt <- paste0(
                "Actúe como estadístico y redactor científico. ",
                "Redacte un párrafo para describir una figura (gráfico de barras) en la sección de Resultados. ",
                "Use únicamente los datos resumidos proporcionados. ",
                "Si hay grupos (leyenda), resuma diferencias relevantes sin especular. ",
                "No incluya título ni leyenda literal; solo el texto del resultado.\n\n",
                txt_plot
            )
            plot_desc_vals[[i]](gemini_chat_text(prompt))
            updateTabsetPanel(session, paste0("subtab_", i), selected = "Gráfico")
        }, ignoreInit = TRUE)
    }
    
    # ====== INICIALIZAR TABs 1–5 ======
    lapply(1:5, build_tab)
    
    # ==========================================================
    # COMPILAR RESULTADOS (vista)
    # ==========================================================
    lapply(1:5, function(i) {
        output[[paste0("gt_comp_", i)]] <- gt::render_gt({
            req(ds_cache(), label_map())
            tbl <- tbl_store[[i]]()
            validate(need(!is.null(tbl), ""))
            make_gt_with_meta(i, tbl)
        })
    })
    
    output$compiled_view <- renderUI({
        req(ds_cache(), label_map())
        
        blocks <- lapply(1:5, function(i) {
            tbl_gts <- tbl_store[[i]]()
            if (is.null(tbl_gts)) return(NULL)
            
            title_txt <- input[[paste0("title_", i)]]
            if (is.null(title_txt) || !nzchar(trimws(title_txt))) title_txt <- paste0("TAB ", i)
            
            desc_txt <- desc_vals[[i]]()
            has_desc <- nzchar(trimws(desc_txt)) && !identical(desc_txt, "Generando descripción...")
            
            fig_desc <- plot_desc_vals[[i]]()
            has_fig_desc <- nzchar(trimws(fig_desc)) && !identical(fig_desc, "Generando descripción...")
            
            tags$div(
                style = "margin-bottom: 26px;",
                tags$h4(style = "margin-bottom:10px;", paste0("TAB ", i, ": ", title_txt)),
                
                if (has_desc) tags$div(
                    style = "background:#f8f9fa; border:1px solid #e9ecef; padding:12px; border-radius:8px; margin-bottom:10px;",
                    tags$b("Texto (tabla):"),
                    tags$br(), tags$br(),
                    tags$div(style = "white-space:pre-wrap;", desc_txt)
                ) else NULL,
                
                if (has_fig_desc) tags$div(
                    style = "background:#f8f9fa; border:1px solid #e9ecef; padding:12px; border-radius:8px; margin-bottom:10px;",
                    tags$b("Texto (figura):"),
                    tags$br(), tags$br(),
                    tags$div(style = "white-space:pre-wrap;", fig_desc)
                ) else NULL,
                
                gt::gt_output(paste0("gt_comp_", i)),
                tags$br(),
                tags$div(
                    style="color:#6c757d; font-size:0.9em;",
                    "Figura: disponible si fue generada en la pestaña Gráfico."
                ),
                tags$hr()
            )
        })
        
        blocks <- Filter(Negate(is.null), blocks)
        
        if (length(blocks) == 0) {
            return(tags$div(
                style = "color:#6c757d;",
                "No hay TABs para compilar. Abra al menos un TAB (1–5) para que se guarde la tabla y luego vuelva a Compilar Resultados."
            ))
        }
        
        do.call(tagList, blocks)
    })
    
    observeEvent(input$view_compiled, {
        showModal(modalDialog(
            title = "Compilación de Resultados (TAB 1 a TAB 5)",
            size = "l",
            easyClose = TRUE,
            footer = modalButton("Cerrar"),
            uiOutput("compiled_view")
        ))
    }, ignoreInit = TRUE)
    
    # ==========================================================
    # DESCARGA WORD
    # ==========================================================
    output$download_word <- downloadHandler(
        filename = function() paste0("Resultados_", format(Sys.Date(), "%Y%m%d"), ".docx"),
        contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        content = function(file) {
            
            if (!grepl("\\.docx$", file, ignore.case = TRUE)) file <- paste0(file, ".docx")
            
            lang <- get_lang()
            src_prefix <- if (lang == "es") "Fuente: " else "Source: "
            
            safe_docx <- function(msg) {
                d <- officer::read_docx()
                d <- officer::body_add_par(d, msg, style = "Normal")
                print(d, target = file)
            }
            
            tryCatch({
                doc <- officer::read_docx()
                included_any <- FALSE
                
                for (i in 1:5) {
                    tbl_gts <- tbl_store[[i]]()
                    if (is.null(tbl_gts)) next
                    
                    included_any <- TRUE
                    
                    title_txt <- input[[paste0("title_", i)]]
                    if (is.null(title_txt) || !nzchar(trimws(title_txt))) title_txt <- paste0("TAB ", i)
                    
                    source_txt <- input[[paste0("source_", i)]]
                    if (is.null(source_txt) || !nzchar(trimws(source_txt))) source_txt <- "Null"
                    
                    desc_txt <- desc_vals[[i]]()
                    has_desc <- nzchar(trimws(desc_txt)) && !identical(desc_txt, "Generando descripción...")
                    
                    fig_desc <- plot_desc_vals[[i]]()
                    has_fig_desc <- nzchar(trimws(fig_desc)) && !identical(fig_desc, "Generando descripción...")
                    
                    doc <- officer::body_add_par(doc, paste0("TAB ", i, ": ", title_txt), style = "heading 1")
                    
                    if (has_desc) {
                        doc <- officer::body_add_par(doc, "Texto (tabla):", style = "heading 2")
                        doc <- officer::body_add_par(doc, desc_txt, style = "Normal")
                    }
                    
                    if (has_fig_desc) {
                        doc <- officer::body_add_par(doc, "Texto (figura):", style = "heading 2")
                        doc <- officer::body_add_par(doc, fig_desc, style = "Normal")
                    }
                    
                    doc <- officer::body_add_par(doc, "Tabla:", style = "heading 2")
                    ft <- tryCatch(gtsummary::as_flex_table(tbl_gts), error = function(e) NULL)
                    
                    if (is.null(ft)) {
                        doc <- officer::body_add_par(
                            doc,
                            "No fue posible convertir la tabla a formato Word (flextable).",
                            style = "Normal"
                        )
                    } else {
                        ft <- flextable::autofit(ft)
                        doc <- tryCatch(
                            flextable::body_add_flextable(doc, value = ft),
                            error = function(e) officer::body_add_par(doc, paste0("Error al insertar tabla: ", e$message), style = "Normal")
                        )
                    }
                    
                    p <- plot_store[[i]]()
                    if (!is.null(p)) {
                        tmp_png <- tempfile(fileext = ".png")
                        if (requireNamespace("ragg", quietly = TRUE)) {
                            ragg::agg_png(filename = tmp_png, width = 2400, height = 1600, res = 300)
                            print(p); dev.off()
                        } else {
                            png(tmp_png, width = 2400, height = 1600, res = 300)
                            print(p); dev.off()
                        }
                        doc <- officer::body_add_par(doc, "Figura:", style = "heading 2")
                        doc <- officer::body_add_img(doc, src = tmp_png, width = 6.5, height = 4.3, style = "centered")
                    }
                    
                    doc <- officer::body_add_par(doc, paste0(src_prefix, source_txt), style = "Normal")
                    doc <- officer::body_add_par(doc, " ", style = "Normal")
                }
                
                if (!included_any) {
                    doc <- officer::body_add_par(doc, "No se encontraron TABs con resultados para compilar.", style = "Normal")
                }
                
                print(doc, target = file)
            }, error = function(e) {
                safe_docx(paste0("Error al generar el Word: ", e$message))
            })
        }
    )
    
    # ==========================================================
    # MÉTODOS
    # ==========================================================
    methods_val <- reactiveVal("")
    
    output$methods_text <- renderUI({
        txt <- methods_val()
        if (!nzchar(trimws(txt))) return(NULL)
        tags$div(
            style = "background:#f8f9fa; border:1px solid #e9ecef; padding:12px; border-radius:8px;",
            tags$b("Métodos (Gemini):"),
            tags$br(), tags$br(),
            tags$div(style = "white-space:pre-wrap;", txt)
        )
    })
    
    observeEvent(input$gen_methods, {
        req(ds_cache(), label_map())
        methods_val("Generando métodos...")
        
        lang <- get_lang()
        dec  <- get_decimals()
        desc_choice <- input$q_desc
        disp_choice <- input$q_disp
        cat_ci <- isTRUE(input$cat_ci95)
        
        cont_desc_es <- if (disp_choice == "ci95") {
            if (desc_choice == "mean") "media con IC95% (columna adicional de IC)" else "mediana con IC95% (columna adicional de IC)"
        } else if (desc_choice == "mean" && disp_choice == "sd") "media (desviación estándar)"
        else if (desc_choice == "mean" && disp_choice == "range") "media (mínimo, máximo)"
        else if (desc_choice == "mean" && disp_choice == "iqr") "media (percentiles 25 y 75)"
        else if (desc_choice == "median" && disp_choice == "sd") "mediana (desviación estándar)"
        else if (desc_choice == "median" && disp_choice == "range") "mediana (mínimo, máximo)"
        else if (desc_choice == "median" && disp_choice == "iqr") "mediana (rango intercuartílico; percentiles 25 y 75)"
        else "media (desviación estándar)"
        
        cont_desc_en <- if (disp_choice == "ci95") {
            if (desc_choice == "mean") "mean with 95% CI (additional CI column)" else "median with 95% CI (additional CI column)"
        } else if (desc_choice == "mean" && disp_choice == "sd") "mean (standard deviation)"
        else if (desc_choice == "mean" && disp_choice == "range") "mean (minimum, maximum)"
        else if (desc_choice == "mean" && disp_choice == "iqr") "mean (25th, 75th percentiles)"
        else if (desc_choice == "median" && disp_choice == "sd") "median (standard deviation)"
        else if (desc_choice == "median" && disp_choice == "range") "median (minimum, maximum)"
        else if (desc_choice == "median" && disp_choice == "iqr") "median (interquartile range; 25th, 75th percentiles)"
        else "mean (standard deviation)"
        
        cat_desc_es <- if (cat_ci) "frecuencias y porcentajes, con IC95% cuando fue seleccionado" else "frecuencias y porcentajes"
        cat_desc_en <- if (cat_ci) "counts and percentages, with 95% confidence intervals when selected" else "counts and percentages"
        
        prompt_methods <- if (lang == "es") {
            paste0(
                "Actúe como estadístico experto en redacción de artículos científicos.\n",
                "Redacte la sección de Métodos de análisis estadístico (solo métodos), en español, en tiempo pasado.\n",
                "Describa EXACTAMENTE:\n",
                "1) Tablas descriptivas en R con RStudio (última versión disponible al momento del análisis).\n",
                "2) Continuas: ", cont_desc_es, ".\n",
                "3) Categóricas: ", cat_desc_es, ".\n",
                "4) Redondeo uniforme de ", dec, " decimales para porcentajes, estimaciones, rangos, IC y valores p.\n",
                "5) Cuando hubo agrupación, se mostraron resultados por grupo; si se seleccionó, se incluyó columna Total.\n",
                "6) Cuando se activó comparación, se calcularon valores p con pruebas apropiadas sin especificar nombres.\n",
                "7) Se incluyeron conteos (n) y se manejaron datos faltantes cuando correspondió.\n",
                "8) Si se seleccionó, se reportaron IC95%.\n\n",
                "Salida: 1-2 párrafos, sin títulos, sin mencionar Shiny ni botones."
            )
        } else {
            paste0(
                "Act as an expert biostatistician writing for a scientific manuscript.\n",
                "Write the Statistical Analysis Methods section (methods only), in English, in past tense.\n",
                "Describe EXACTLY:\n",
                "1) Descriptive tables in R using RStudio (latest available version at time of analysis).\n",
                "2) Continuous: ", cont_desc_en, ".\n",
                "3) Categorical: ", cat_desc_en, ".\n",
                "4) Uniform rounding of ", dec, " decimals for percentages, estimates, ranges, CIs and p-values.\n",
                "5) When grouped, group-wise results were shown; if selected, a Total column was included.\n",
                "6) When comparisons were enabled, p-values were computed with appropriate tests without naming them.\n",
                "7) Counts (n) included and missing handled when present.\n",
                "8) If selected, 95% CIs were reported.\n\n",
                "Output: 1–2 paragraphs, no headings, do not mention Shiny/buttons."
            )
        }
        
        methods_val(gemini_chat_text(prompt_methods))
    }, ignoreInit = TRUE)
}

shinyApp(ui = ui, server = server)