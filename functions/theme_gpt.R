#' **Generate a thematic code list using a local LLM**
#'
#' @description
#' This function analyzes open-ended survey responses and automatically generates a set of thematic codes with descriptions using a local Ollama server.
#'
#' @param data A data frame containing the survey data, or a path to a `.xlsx` or `.xls` file.
#' @param x The open-ended variable to analyze.
#' @param n Integer; number of themes to return. If `NULL` (default), the model determines an appropriate number of themes based on the responses.
#' @param sample Optional integer specifying the number of responses to sample for analysis. If `NULL`, all valid responses are used.
#' @param model Character string; the Ollama model to use. Defaults to `llama3.2:3b`.
#' @param instructions Optional string; additional instructions for coding.
#'
#' @details
#' The output is a tibble with three columns:
#' - `Code`: A unique numeric code for each theme (standard codes 97–99 are added automatically).
#' - `Bin`: Short label for the theme, written in sentence case.
#' - `Description`: A one-sentence summary describing the theme's content.
#'
#' Standard codes are included automatically:
#' - `97` = `Other`
#' - `98` = `None`
#' - `99` = `Don't know`
#' 
#' Use this function to create a `theme_list` for input into `code_gpt()`.
#'
#' **Note:** It's best to review and refine the generated codes before using them in `code_gpt()`.
#'
#' @return
#' A table containing the generated thematic code list and their description. Standard codes (`Other`, `None`, `Don't know`) are included automatically.
#' 
#' @export

theme_gpt <- function(data, x, n = NULL, sample = NULL, model = "llama3.2:3b", instructions = NULL) {
  # Check required packages ----
  
  required_pkgs <- c("tibble", "dplyr", "readxl", "purrr", "stringr", "httr2")
  missing_pkgs <- required_pkgs[!vapply(required_pkgs, requireNamespace, logical(1), quietly = TRUE)]
  if (length(missing_pkgs) > 0) {
    stop("These packages are required but not installed: ",
         paste(missing_pkgs, collapse = ", "), call. = FALSE)
  }
  
  # Set up ----
  
  `%>%` <- dplyr::`%>%`
  
  # Ollama endpoint config
  base_url <- Sys.getenv("LLM_BASE_URL", "http://localhost:11434/v1")
  base_url <- sub("/+$", "", base_url)
  model <- Sys.getenv("LLM_MODEL", model)
  
  # Load data (data frame or file path)
  read_input_data <- function(data) {
    if (is.data.frame(data)) return(data)
    if (is.character(data) && length(data) == 1) {
      ext <- tolower(tools::file_ext(data))
      if (ext %in% c("xlsx", "xls")) return(readxl::read_excel(data))
      stop("Unsupported file type: .", ext, ". Please upload a .xlsx or .xls file.", call. = FALSE)
    }
    stop("`data` must be a data frame or a file path.", call. = FALSE)
  }
  
  data <- read_input_data(data)
  
  # Validate that x exists in dataset
  if (!(x %in% names(data))) stop("Variable ", x, " was not found in the dataset.", call. = FALSE)
  
  # Identify question label
  q_label <- attr(data[[x]], "label", exact = TRUE)
  if (is.null(q_label)) q_label <- x
  
  # Extract valid responses
  responses <- data[[x]]
  responses <- responses[!is.na(responses) & stringr::str_trim(responses) != ""]
  if (length(responses) == 0) stop("No valid responses found in variable ", x, ".", call. = FALSE)
  
  # Determine sample size
  if (is.null(sample) || sample >= length(responses)) {
    sample_resps <- responses
    sample_n <- length(responses)
  } else {
    sample_resps <- sample(responses, sample)
    sample_n <- sample
  }
  
  sample_text <- paste0(seq_along(sample_resps), ". ", sample_resps, collapse = "\n")
  
  # Theme count instruction
  if (is.null(n)) {
    n_text <- "an appropriate set of"
  } else {
    n_text <- paste(n)
  }
  
  # Ollama API request ----
  
  req <- httr2::request(paste0(base_url, "/chat/completions")) %>%
    httr2::req_headers("Content-Type" = "application/json") %>%
    httr2::req_body_json(list(
      model = model,
      temperature = 0,
      seed = 123,
      messages = list(
        list(
          role = "system",
          content = paste(
            "You are a survey researcher designing thematic code frames for qualitative survey data.",
            "You will be shown real open-ended responses. Identify recurring ideas and group them into clear, distinct themes.",
            "Each theme should have a short, descriptive name in *sentence case* (e.g., 'Situational or environmental constraints').",
            "Each description must begin with 'This theme...' and be one concise sentence describing what the theme includes.",
            "Return the output in plain CSV-style text with two columns: Bin, Description.",
            "Do not number or bullet the themes."
          )
        ),
        list(
          role = "user",
          content = paste0(
            "Here are ", sample_n, " example responses to the survey question: '", q_label, "'.\n\n",
            sample_text,
            "\n\nBased on these, create ", n_text,
            " thematic codes with one-sentence descriptions.\nFormat as:\nBin, Description",
            if (!is.null(instructions)) paste0("\n\nAdditional instructions: ", instructions) else ""
          )
        )
      )
    )) %>%
    httr2::req_options(timeout = 120)
  
  resp <- tryCatch(
    httr2::req_perform(req),
    error = function(e) {
      msg <- conditionMessage(e)
      if (grepl("Failed to connect|Connection refused|could not connect|timeout", msg, ignore.case = TRUE)) {
        stop(
          paste0(
            "Could not reach Ollama at ", base_url, ".\n\n",
            "Make sure Ollama is running and the model is pulled:\n",
            "  ollama run ", model, "\n\n",
            "Then retry."
          ),
          call. = FALSE
        )
      }
      stop(msg, call. = FALSE)
    }
  )
  
  status <- httr2::resp_status(resp)
  if (status >= 400) {
    body <- tryCatch(httr2::resp_body_json(resp), error = function(...) NULL)
    detail <- if (!is.null(body$error$message)) body$error$message else ""
    if (grepl("model.*not found", detail, ignore.case = TRUE)) {
      stop(paste0("Model '", model, "' not found. Run: ollama run ", model), call. = FALSE)
    }
    stop(paste0("LLM request failed (HTTP ", status, "): ", detail), call. = FALSE)
  }
  
  result <- httr2::resp_body_json(resp)
  text <- NULL
  if (!is.null(result$choices) && length(result$choices) > 0) {
    text <- result$choices[[1]]$message$content
  }
  if (is.null(text) || text == "") {
    stop(
      paste0("LLM returned empty response. Ensure Ollama is running:\n  ollama run ", model),
      call. = FALSE
    )
  }
  
  # Output ----
  
  # Parse LLM output
  lines <- stringr::str_split(text, "\n")[[1]]
  lines <- stringr::str_trim(lines)
  lines <- lines[lines != ""]
  
  # Build data frame row by row
  rows <- vector("list", length(lines))
  for (i in seq_along(lines)) {
    line <- lines[[i]]
    parts <- strsplit(line, ",", fixed = TRUE)[[1]]
    if (length(parts) >= 2) {
      rows[[i]] <- tibble::tibble(
        Bin = stringr::str_trim(parts[1]),
        Description = stringr::str_trim(paste(parts[-1], collapse = ","))
      )
    } else {
      rows[[i]] <- tibble::tibble(Bin = stringr::str_trim(parts[1]), Description = NA_character_)
    }
  }
  
  df <- dplyr::bind_rows(rows)
  
  # Add codes
  df <- df %>%
    dplyr::filter(!(Bin %in% c("Bin", "Description") | Description %in% c("Bin", "Description"))) %>%
    dplyr::filter(!stringr::str_detect(Bin, "^Here are the")) %>%
    dplyr::mutate(Code = dplyr::row_number()) %>%
    dplyr::select(Code, Bin, Description)
  
  # Prepend standard codes
  standard <- tibble::tibble(
    Code = c(97, 98, 99),
    Bin = c("Other", "None", "Don't know"),
    Description = c(
      "Response does not fit into any existing categories or represents a unique situation not captured by other codes.",
      "Response is irrelevant, nonsensical, or provides no meaningful information (e.g., gibberish, off-topic, or empty text).",
      "Respondent expresses uncertainty, confusion, or lack of an opinion or knowledge about the topic."
    )
  )
  
  result <- dplyr::bind_rows(standard, df) %>%
    dplyr::mutate(Code = as.integer(Code)) %>%
    dplyr::relocate(Code, Bin, Description)
  
  invisible(result)
}
