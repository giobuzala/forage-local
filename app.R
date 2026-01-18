library(shiny)
library(shinyjs)

# Environment loading ----

load_env <- function() {
  if (file.exists(".env")) {
    if (!requireNamespace("dotenv", quietly = TRUE)) {
      stop(
        "Package 'dotenv' is required to load .env. Please install it with install.packages('dotenv').",
        call. = FALSE
      )
    }
    dotenv::load_dot_env(file = ".env")
  }
}

load_env()

source("Functions/code_gpt.R")
source("Functions/theme_gpt.R")

# Helpers ----

read_input_file <- function(path) {
  ext <- tolower(tools::file_ext(path))
  if (ext %in% c("xlsx", "xls")) {
    return(readxl::read_excel(path))
  }
  stop(
    "Unsupported file type: .", ext,
    ". Please upload a .xlsx or .xls file.",
    call. = FALSE
  )
}

parse_n_themes <- function(x) {
  if (is.null(x)) return(NULL)
  x <- trimws(x)
  if (x == "") return(NULL)
  n <- suppressWarnings(as.integer(x))
  if (is.na(n) || n <= 0) return(NULL)
  n
}

# UI ----

ui <- fluidPage(
  useShinyjs(),
  
  tags$head(
    tags$style(HTML("

      /* ===== GLOBAL ===== */

      html {
        font-size: 16px;
      }

      body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        color: #1f2933;
        font-size: 1rem;
      }

      .container {
        max-width: 760px;
        margin: 56px auto;
      }

      /* ===== HEADINGS ===== */

      h2 {
        font-size: 3rem;
        font-weight: 700;
        letter-spacing: -0.02em;
        margin-bottom: 14px;
      }

      h4 {
        font-size: 1.55rem;
        font-weight: 600;
        margin-bottom: 20px;
      }

      .step-num {
        font-size: 1.55rem;
        font-weight: 600;
        margin-right: 8px;
      }

      /* ===== TEXT ===== */

      .lead {
        font-size: 1.25rem;
        color: #4b5563;
        margin-bottom: 24px;
      }

      .subtle {
        font-size: 0.95rem;
        color: #6b7280;
        line-height: 1.6;
      }

      /* ===== PANELS ===== */

      .panel {
        border: 1px solid #6b7280;
        padding: 18px 26px;
        border-radius: 14px;
        margin-top: 16px;
        background: #ffffff;
      }
      
      .spacer {
        height: 15px;
      }

      /* ===== FORM ELEMENTS ===== */

      label {
        display: block;
        font-size: 1.05rem;
        font-weight: 500;
        margin-bottom: 10px;
      }

      .shiny-input-container input,
      .shiny-input-container select {
        font-size: 0.95rem;
        padding: 10px 14px;
        height: auto;
      }

      input[type='file'] {
        font-size: 1rem;
      }

      /* Remove Bootstrap spacing under inputs inside panels */
      .panel .form-group {
        margin-bottom: 0;
      }

      /* ===== CHECKBOX ===== */

      .panel .checkbox {
        margin-top: 12px;
      }

      .panel .checkbox label {
        font-size: 1rem;
        font-weight: 500;
        margin-bottom: 0;
      }

      /* ===== BUTTONS ===== */

      .btn {
        font-size: 0.95rem;
        font-weight: 600;
        padding: 10px 22px;
      }

      .btn-primary,
      .btn-primary:hover,
      .btn-primary:focus,
      .btn-primary:active {
        background-color: #1f2933;
        border-color: #1f2933;
      }

      .btn-download {
        background-color: #065f46;
        border-color: #065f46;
        color: #ffffff;
      }

      /* ===== STATUS ===== */

      .notice {
        font-size: 0.95rem;
        font-weight: 500;
        background: #f9fafb;
        border: 1px solid #4b5563;
        padding: 12px 18px;
        border-radius: 14px;
        margin-top: 20px;
        display: inline-block;
      }

      /* ===== INSTRUCTION LIST ===== */

      .panel ol {
        margin: 4px 0;
        padding-left: 22px;
      }

    "))
  ),
  
  div(
    class = "container",
    
    h2("forage"),
    p(
      class = "lead",
      "AI-assisted open-ended coding agent"
    ),
    
    div(
      class = "panel",
      tags$ol(
        tags$li(
          strong("Upload your survey file:"),
          " Select the file containing your open-ended survey responses."
        ),
        tags$li(
          strong("Generate a theme list:"),
          " The app will analyze responses and generate themes with clear labels and descriptions.
          You can download the list for review or proceed to the next step using it as generated."
        ),
        tags$li(
          strong("Code responses and download results:"),
          " If you are not using the theme list generated in the previous step, upload your own theme list.
          Each response will be assigned one or more themes, and you can download the coded file."
        )
      )
    ),
    
    uiOutput("status_banner"),
    
    # Step 1 ----
    
    div(
      class = "panel",
      h4(tags$span("Step 1.", class = "step-num"), "Upload survey data"),
      tags$small(
        class = "subtle",
        "Upload an Excel file that includes an ID column and an open-ended response column."
      ),
      div(class = "spacer"),
      fileInput("data_file", "Survey file", accept = c(".xlsx", ".xls")),
      uiOutput("column_selectors")
    ),
    
    # Step 2 ----
    
    div(
      class = "panel",
      h4(tags$span("Step 2.", class = "step-num"), "Generate themes"),
      p(class = "subtle", "Generate a theme list from your open-ended responses."),
      textInput("n_themes", "Number of themes", value = ""),
      div(class = "spacer"),
      tags$small(class = "subtle", "Leave blank to let the app decide the number of themes."),
      div(class = "spacer"),
      actionButton("run_themes", "Generate theme list", class = "btn-primary"),
      div(class = "spacer"),
      uiOutput("download_themes_ui")
    ),
    
    # Step 3 ----
    
    div(
      class = "panel",
      h4(tags$span("Step 3.", class = "step-num"), "Code responses"),
      p(
        class = "subtle",
        "Use the generated themes or upload an existing theme list.
         As a best practice, download and review the generated themes before uploading them for coding."
      ),
      p(
        class = "subtle",
        "Results are returned as an Excel file with coded responses and a theme list."
      ),
      fileInput("theme_file", "Theme list", accept = c(".xlsx", ".xls")),
      tags$small(class = "subtle", "Leave empty to let the app generate themes."),
      checkboxInput("use_generated", "Use generated themes", value = TRUE),
      div(class = "spacer"),
      actionButton("run", "Code responses", class = "btn-primary"),
      div(class = "spacer"),
      uiOutput("download_coded_ui")
    )
  )
)

# Server ----

server <- function(input, output, session) {
  
  coded_data <- reactiveVal(NULL)
  generated_themes <- reactiveVal(NULL)
  theme_used <- reactiveVal(NULL)
  status_text <- reactiveVal("Waiting for input.")
  
  data_df <- reactive({
    req(input$data_file)
    read_input_file(input$data_file$datapath)
  })
  
  theme_df <- reactive({
    if (is.null(input$theme_file)) return(NULL)
    read_input_file(input$theme_file$datapath)
  })
  
  output$column_selectors <- renderUI({
    req(data_df())
    cols <- names(data_df())
    tagList(
      selectInput("id_col", "ID column", choices = cols),
      selectInput("response_col", "Open-ended response column", choices = cols)
    )
  })
  
  output$status_banner <- renderUI({
    div(class = "notice", tags$strong("Status: "), status_text())
  })
  
  observe({
    toggleState("run_themes", !is.null(input$data_file))
    toggleState("run", !is.null(generated_themes()) || !is.null(theme_df()))
  })
  
  observeEvent(input$run_themes, {
    req(data_df(), input$response_col)
    
    withProgress(
      message = "Analyzing responses",
      detail = "Generating theme list...",
      {
        result <- tryCatch(
          theme_gpt(
            data = data_df(),
            x = input$response_col,
            n = parse_n_themes(input$n_themes),
            sample = NULL
          ),
          error = function(e) {
            status_text(e$message)
            NULL
          }
        )
        
        if (!is.null(result)) {
          generated_themes(result)
          status_text("Themes generated. You can download them or proceed to coding.")
        }
      }
    )
  })
  
  observeEvent(input$run, {
    req(data_df(), input$id_col, input$response_col)
    
    withProgress(
      message = "Coding responses",
      detail = "Assigning themes to each response...",
      {
        result <- tryCatch(
          {
            theme_list <- if (isTRUE(input$use_generated)) generated_themes() else theme_df()
            
            if (is.null(theme_list)) {
              stop("Please generate or upload a theme list first.", call. = FALSE)
            }
            
            if (!all(c("Code", "Bin") %in% names(theme_list))) {
              stop("Theme list format is invalid.", call. = FALSE)
            }
            
            theme_used(theme_list)
            
            code_gpt(
              data = data_df(),
              x = input$response_col,
              id_var = input$id_col,
              theme_list = theme_list
            )
          },
          error = function(e) {
            status_text(e$message)
            NULL
          }
        )
        
        if (!is.null(result)) {
          coded_data(result)
          status_text("Coding complete. Download your results below.")
        }
      }
    )
  })
  
  output$download_themes_ui <- renderUI({
    req(generated_themes())
    downloadButton("download_themes", "Download theme list", class = "btn-download")
  })
  
  output$download_themes <- downloadHandler(
    filename = function() "Theme List.xlsx",
    content = function(file) {
      writexl::write_xlsx(generated_themes(), file)
    }
  )
  
  output$download_coded_ui <- renderUI({
    req(coded_data())
    downloadButton("download", "Download coded file", class = "btn-download")
  })
  
  output$download <- downloadHandler(
    filename = function() "Coded Responses.xlsx",
    content = function(file) {
      writexl::write_xlsx(
        list(
          "Coded Responses" = coded_data(),
          "Theme List" = theme_used()
        ),
        file
      )
    }
  )
}

shinyApp(ui, server)
