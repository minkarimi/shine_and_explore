# app.R
# ============================================================
# General Excel Explorer (Shiny)
# - Upload an Excel file
# - Choose a sheet
# - Optional: apply multiple filters (categorical + numeric range)
# - Create basic plots: bar, pie, scatter, line
# - Custom colors (palette / manual mapping)
#
# Dependencies:
# install.packages(c("shiny","readxl","dplyr","tidyr","ggplot2","stringr"))
# ============================================================

suppressPackageStartupMessages({
  library(shiny)
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(stringr)
  library(ggplot2)
})

`%||%` <- function(a, b) if (!is.null(a)) a else b

# ----------------------------
# Helpers
# ----------------------------
is_numericish <- function(v) {
  # numeric OR character that contains many numeric values
  if (is.numeric(v)) return(TRUE)
  x <- suppressWarnings(as.numeric(v))
  # consider numeric-ish if enough non-NA numeric conversions
  sum(!is.na(x)) >= max(5, floor(0.2 * length(v)))
}

as_num_safe <- function(v) suppressWarnings(as.numeric(v))

apply_filters_df <- function(df, filters) {
  out <- df
  if (length(filters) == 0) return(out)
  
  for (f in filters) {
    col <- f$col
    if (!col %in% names(out)) next
    
    if (f$type == "cat") {
      vals <- f$values
      if (!is.null(vals) && length(vals) > 0) {
        out <- out %>% filter(as.character(.data[[col]]) %in% vals)
      }
    }
    
    if (f$type == "num") {
      rr <- f$range
      if (!is.null(rr) && length(rr) == 2) {
        x <- as_num_safe(out[[col]])
        out <- out[!is.na(x) & x >= rr[1] & x <= rr[2], , drop = FALSE]
      }
    }
  }
  
  out
}

# palette options (simple & nice)
palette_choices <- c(
  "Set2" = "Set2",
  "Set3" = "Set3",
  "Pastel1" = "Pastel1",
  "Pastel2" = "Pastel2",
  "Dark2" = "Dark2",
  "Paired" = "Paired",
  "Accent" = "Accent"
)

# ----------------------------
# UI
# ----------------------------
ui <- fluidPage(
  titlePanel("Excel Data Explorer (Portfolio Shiny App)"),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("xls", "Upload Excel file", accept = c(".xlsx", ".xls")),
      
      uiOutput("sheet_ui"),
      tags$hr(),
      
      h4("Filters"),
      checkboxInput("use_filters", "Apply filters?", value = TRUE),
      uiOutput("filter_cols_ui"),
      uiOutput("filters_ui"),
      
      tags$hr(),
      h4("Plot builder"),
      selectInput("plot_type", "Plot type",
                  choices = c("Barplot"="bar", "Pie chart"="pie", "Scatter"="scatter", "Line"="line")),
      
      uiOutput("plot_vars_ui"),
      
      tags$hr(),
      h4("Colors"),
      radioButtons("color_mode", "Color mode",
                   choices = c("Palette"="palette", "Manual (map levels)"="manual"),
                   inline = TRUE),
      selectInput("palette_name", "Palette (categorical)", choices = palette_choices, selected = "Set2"),
      uiOutput("manual_color_ui"),
      
      tags$hr(),
      actionButton("apply", "Apply / Refresh", class = "btn-primary")
    ),
    
    mainPanel(
      h4("Data preview (after filters)"),
      verbatimTextOutput("shape_txt"),
      tableOutput("preview_tbl"),
      tags$hr(),
      h4("Plot"),
      plotOutput("plot", height = "450px"),
      tags$hr(),
      h4("Summary"),
      verbatimTextOutput("summary_txt")
    )
  )
)

# ----------------------------
# Server
# ----------------------------
server <- function(input, output, session) {
  
  # list sheets
  sheets <- reactive({
    req(input$xls)
    readxl::excel_sheets(input$xls$datapath)
  })
  
  output$sheet_ui <- renderUI({
    if (is.null(input$xls)) return(tags$em("Upload an Excel file to choose sheet."))
    selectInput("sheet", "Choose sheet", choices = sheets())
  })
  
  # read selected sheet
  raw_df <- reactive({
    req(input$xls, input$sheet)
    readxl::read_excel(input$xls$datapath, sheet = input$sheet)
  })
  
  # update filter/plot choices when data changes
  observeEvent(raw_df(), {
    cols <- names(raw_df())
    updateSelectInput(session, "x_col", choices = cols, selected = cols[1] %||% "")
    updateSelectInput(session, "y_col", choices = cols, selected = cols[2] %||% cols[1] %||% "")
    updateSelectInput(session, "group_col", choices = c("", cols), selected = "")
    updateSelectInput(session, "filter_cols", choices = cols, selected = character(0))
  })
  
  output$filter_cols_ui <- renderUI({
    req(raw_df())
    if (!isTRUE(input$use_filters)) return(tags$em("Filters disabled."))
    selectizeInput(
      "filter_cols",
      "Select one or more columns to filter",
      choices = names(raw_df()),
      multiple = TRUE,
      options = list(placeholder = "Pick columns...")
    )
  })
  
  # build dynamic filter widgets
  output$filters_ui <- renderUI({
    req(raw_df())
    if (!isTRUE(input$use_filters)) return(NULL)
    if (is.null(input$filter_cols) || length(input$filter_cols) == 0) return(tags$em("No filter columns selected."))
    
    df <- raw_df()
    
    widgets <- lapply(input$filter_cols, function(col) {
      v <- df[[col]]
      
      if (is_numericish(v)) {
        x <- as_num_safe(v)
        x <- x[!is.na(x)]
        if (!length(x)) return(tags$div(strong(col), ": numeric but no numeric values found"))
        rng <- range(x)
        sliderInput(
          inputId = paste0("rng__", col),
          label = paste0(col, " (range)"),
          min = floor(rng[1]),
          max = ceiling(rng[2]),
          value = c(floor(rng[1]), ceiling(rng[2]))
        )
      } else {
        lev <- sort(unique(as.character(v)))
        lev <- lev[!is.na(lev) & trimws(lev) != ""]
        selectizeInput(
          inputId = paste0("cat__", col),
          label = paste0(col, " (values)"),
          choices = lev,
          multiple = TRUE,
          options = list(placeholder = "Select values...")
        )
      }
    })
    
    do.call(tagList, widgets)
  })
  
  # plot variables UI
  output$plot_vars_ui <- renderUI({
    req(raw_df())
    cols <- names(raw_df())
    
    # For bar/pie: x is category; optional group as fill
    if (input$plot_type %in% c("bar", "pie")) {
      tagList(
        selectInput("x_col", "Category (X)", choices = cols),
        selectInput("group_col", "Optional group (fill)", choices = c("", cols), selected = "")
      )
    } else {
      # scatter/line: x and y should be numeric-ish
      tagList(
        selectInput("x_col", "X axis", choices = cols),
        selectInput("y_col", "Y axis", choices = cols),
        selectInput("group_col", "Optional group (color)", choices = c("", cols), selected = "")
      )
    }
  })
  
  # manual colors UI depends on selected group_col
  output$manual_color_ui <- renderUI({
    req(raw_df())
    if (input$color_mode != "manual") return(NULL)
    if (is.null(input$group_col) || !nzchar(input$group_col)) {
      return(tags$em("Choose a group column to define manual colors."))
    }
    
    df <- raw_df()
    v <- df[[input$group_col]]
    lev <- sort(unique(as.character(v)))
    lev <- lev[!is.na(lev) & trimws(lev) != ""]
    if (!length(lev)) return(tags$em("No levels found in selected group column."))
    
    # One text input where user sets mappings: level=#RRGGBB, ...
    default_map <- paste0(lev, "=", "#", rep("1f77b4", length(lev)), collapse = ", ")
    tagList(
      tags$div(style="font-size:12px;color:#555;",
               "Manual mapping format: Level=#RRGGBB, Other=#RRGGBB"),
      textAreaInput("manual_map", "Manual color mapping", value = default_map, rows = 3)
    )
  })
  
  # parse manual color mapping: "A=#ff0000, B=#00ff00"
  parse_manual_map <- function(txt) {
    if (is.null(txt) || !nzchar(txt)) return(NULL)
    parts <- unlist(strsplit(txt, ","))
    parts <- trimws(parts)
    kv <- strsplit(parts, "=", fixed = TRUE)
    
    keys <- vapply(kv, function(x) if (length(x) >= 1) trimws(x[[1]]) else NA_character_, character(1))
    vals <- vapply(kv, function(x) if (length(x) >= 2) trimws(x[[2]]) else NA_character_, character(1))
    
    ok <- !is.na(keys) & keys != "" & !is.na(vals) & grepl("^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$", vals)
    if (!any(ok)) return(NULL)
    
    out <- vals[ok]
    names(out) <- keys[ok]
    out
  }
  
  # Build filters structure
  current_filters <- reactive({
    req(raw_df())
    if (!isTRUE(input$use_filters) || is.null(input$filter_cols) || length(input$filter_cols) == 0) return(list())
    
    df <- raw_df()
    filters <- list()
    
    for (col in input$filter_cols) {
      if (!col %in% names(df)) next
      v <- df[[col]]
      
      if (is_numericish(v)) {
        rr <- input[[paste0("rng__", col)]]
        if (!is.null(rr) && length(rr) == 2) {
          filters[[length(filters) + 1]] <- list(col = col, type = "num", range = rr)
        }
      } else {
        vv <- input[[paste0("cat__", col)]]
        if (!is.null(vv) && length(vv) > 0) {
          filters[[length(filters) + 1]] <- list(col = col, type = "cat", values = vv)
        }
      }
    }
    
    filters
  })
  
  # filtered df
  df_filtered <- eventReactive(input$apply, {
    req(raw_df())
    df <- raw_df()
    df2 <- apply_filters_df(df, current_filters())
    df2
  }, ignoreInit = FALSE)
  
  output$shape_txt <- renderPrint({
    req(df_filtered())
    df <- df_filtered()
    list(rows = nrow(df), cols = ncol(df), colnames = names(df))
  })
  
  output$preview_tbl <- renderTable({
    req(df_filtered())
    head(df_filtered(), 15)
  })
  
  output$summary_txt <- renderPrint({
    req(df_filtered())
    df <- df_filtered()
    
    # quick stats for numeric-ish columns
    num_cols <- names(df)[vapply(df, is_numericish, logical(1))]
    if (!length(num_cols)) {
      return(list(message = "No numeric-like columns detected."))
    }
    
    sm <- lapply(num_cols, function(nm) {
      x <- as_num_safe(df[[nm]])
      c(
        col = nm,
        n = sum(!is.na(x)),
        mean = mean(x, na.rm = TRUE),
        median = median(x, na.rm = TRUE),
        sd = sd(x, na.rm = TRUE),
        min = min(x, na.rm = TRUE),
        max = max(x, na.rm = TRUE)
      )
    })
    sm
  })
  
  output$plot <- renderPlot({
    req(df_filtered(), input$plot_type, input$x_col)
    df <- df_filtered()
    
    # if group col is blank, treat as no grouping
    gcol <- input$group_col
    has_group <- !is.null(gcol) && nzchar(gcol) && gcol %in% names(df)
    
    # build color scale
    manual_map <- if (input$color_mode == "manual") parse_manual_map(input$manual_map) else NULL
    
    # ---- bar / pie ----
    if (input$plot_type %in% c("bar", "pie")) {
      xcol <- input$x_col
      validate(need(xcol %in% names(df), "X column not found."))
      
      if (has_group) {
        dd <- df %>%
          mutate(
            x = as.character(.data[[xcol]]),
            g = as.character(.data[[gcol]])
          ) %>%
          mutate(
            x = ifelse(is.na(x) | trimws(x) == "", "(Missing)", x),
            g = ifelse(is.na(g) | trimws(g) == "", "(Missing)", g)
          ) %>%
          count(x, g, name = "n")
      } else {
        dd <- df %>%
          mutate(x = as.character(.data[[xcol]])) %>%
          mutate(x = ifelse(is.na(x) | trimws(x) == "", "(Missing)", x)) %>%
          count(x, name = "n")
      }
      
      if (input$plot_type == "pie") {
        # Pie only makes sense without group, or we facet by group
        if (has_group) {
          p <- ggplot(dd, aes(x = "", y = n, fill = x)) +
            geom_col(width = 1, color = "white") +
            coord_polar(theta = "y") +
            facet_wrap(~ g, scales = "free_y") +
            labs(title = paste("Pie by", xcol, "faceted by", gcol), x = NULL, y = NULL, fill = xcol) +
            theme_void()
        } else {
          p <- ggplot(dd, aes(x = "", y = n, fill = x)) +
            geom_col(width = 1, color = "white") +
            coord_polar(theta = "y") +
            labs(title = paste("Pie:", xcol), x = NULL, y = NULL, fill = xcol) +
            theme_void()
        }
      } else {
        if (has_group) {
          p <- ggplot(dd, aes(x = reorder(x, n), y = n, fill = g)) +
            geom_col(position = "stack") +
            coord_flip() +
            labs(title = paste("Barplot:", xcol, "filled by", gcol), x = xcol, y = "Count", fill = gcol) +
            theme_minimal()
        } else {
          p <- ggplot(dd, aes(x = reorder(x, n), y = n)) +
            geom_col() +
            coord_flip() +
            labs(title = paste("Barplot:", xcol), x = xcol, y = "Count") +
            theme_minimal()
        }
      }
      
      if (has_group && input$plot_type == "bar") {
        # fill by group
        if (!is.null(manual_map)) {
          p <- p + scale_fill_manual(values = manual_map, na.translate = TRUE)
        } else {
          p <- p + scale_fill_brewer(palette = input$palette_name, na.translate = TRUE)
        }
      } else {
        # fill by category (pie) or none
        if (input$plot_type == "pie") {
          if (!is.null(manual_map) && !has_group) {
            p <- p + scale_fill_manual(values = manual_map, na.translate = TRUE)
          } else {
            p <- p + scale_fill_brewer(palette = input$palette_name, na.translate = TRUE)
          }
        }
      }
      
      return(p)
    }
    
    # ---- scatter / line ----
    req(input$y_col)
    xcol <- input$x_col
    ycol <- input$y_col
    validate(need(xcol %in% names(df), "X column not found."))
    validate(need(ycol %in% names(df), "Y column not found."))
    
    x <- df[[xcol]]
    y <- df[[ycol]]
    validate(need(is_numericish(x), "X must be numeric-like for scatter/line."))
    validate(need(is_numericish(y), "Y must be numeric-like for scatter/line."))
    
    dd <- df %>%
      transmute(
        x = as_num_safe(.data[[xcol]]),
        y = as_num_safe(.data[[ycol]]),
        g = if (has_group) as.character(.data[[gcol]]) else NA_character_
      ) %>%
      filter(!is.na(x) & !is.na(y)) %>%
      mutate(g = ifelse(is.na(g) | trimws(g) == "", "(Missing)", g))
    
    if (input$plot_type == "scatter") {
      if (has_group) {
        p <- ggplot(dd, aes(x = x, y = y, color = g)) +
          geom_point() +
          labs(title = paste("Scatter:", ycol, "vs", xcol), x = xcol, y = ycol, color = gcol) +
          theme_minimal()
        
        if (!is.null(manual_map)) {
          p <- p + scale_color_manual(values = manual_map, na.translate = TRUE)
        } else {
          p <- p + scale_color_brewer(palette = input$palette_name, na.translate = TRUE)
        }
      } else {
        p <- ggplot(dd, aes(x = x, y = y)) +
          geom_point() +
          labs(title = paste("Scatter:", ycol, "vs", xcol), x = xcol, y = ycol) +
          theme_minimal()
      }
      return(p)
    }
    
    # line plot (sort by x)
    dd <- dd %>% arrange(x)
    
    if (has_group) {
      p <- ggplot(dd, aes(x = x, y = y, color = g, group = g)) +
        geom_line() +
        geom_point(size = 1) +
        labs(title = paste("Line:", ycol, "vs", xcol), x = xcol, y = ycol, color = gcol) +
        theme_minimal()
      
      if (!is.null(manual_map)) {
        p <- p + scale_color_manual(values = manual_map, na.translate = TRUE)
      } else {
        p <- p + scale_color_brewer(palette = input$palette_name, na.translate = TRUE)
      }
    } else {
      p <- ggplot(dd, aes(x = x, y = y)) +
        geom_line() +
        geom_point(size = 1) +
        labs(title = paste("Line:", ycol, "vs", xcol), x = xcol, y = ycol) +
        theme_minimal()
    }
    
    p
  })
}

shinyApp(ui, server)
