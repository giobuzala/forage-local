#' **Code open-ended survey responses using a local LLM**
#'
#' @description
#' This function submits open-ended survey responses to a local Ollama server and returns a dataset with the original responses, assigned codes, and bins from a provided code list.
#' 
#' @param data A data frame containing the survey data, or a path to a `.xlsx` or `.xls` file.
#' @param x The open-ended variable to be coded.
#' @param theme_list A data frame with at least two columns: `Code` and `Bin`.
#' @param id_var The respondent ID variable.
#' @param n Optional integer; number of responses to code. Defaults to all rows. Useful for testing coding quality without coding every response.
#' @param batch_size Integer; number of responses per API call. Defaults to `100`.
#' @param model Character string; the Ollama model to use. Defaults to `llama3.2:3b`.
#' @param instructions Optional string; additional instructions for coding.
#'
#' @return
#' A table with respondent IDs, responses, codes, and bins.
#'
#' @export

code_gpt <- function(data, x, theme_list, id_var, n = NULL, batch_size = 100, model = "llama3.2:3b", instructions = NULL) {
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
  
  # Validate that x and id_var exist in dataset
  if (!(x %in% names(data))) stop("Variable ", x, " was not found in the dataset.", call. = FALSE)
  if (!(id_var %in% names(data))) stop("ID variable ", id_var, " was not found in the dataset.", call. = FALSE)
  
  # Identify question label
  q_label <- attr(data[[x]], "label", exact = TRUE)
  if (is.null(q_label)) q_label <- x
  
  # Build code list
  if (!("Description" %in% names(theme_list))) {
    theme_list <- dplyr::mutate(theme_list, Description = NA_character_)
  }

  code_text <- paste0(
    "Available codes:\n",
    paste(
      purrr::pmap_chr(theme_list, function(Code, Bin, Description, ...) {
        if (!is.null(Description) && !is.na(Description) && trimws(Description) != "") {
          paste0(Code, ": ", Bin, " — ", Description)
        } else {
          paste0(Code, ": ", Bin)
        }
      }),
      collapse = "\n"
    )
  )
  
  # Limit to first n responses
  if (is.null(n)) {n <- nrow(data)}
  responses <- head(data[[x]], n)
  ids <- head(data[[id_var]], n)
  
  # Identify valid responses
  valid_idx <- which(!is.na(responses) & stringr::str_trim(responses) != "")
  valid_responses <- responses[valid_idx]
  
  # Ollama API request ----
  
  batches <- split(seq_along(valid_responses), ceiling(seq_along(valid_responses) / batch_size))
  
  batch_results <- vector("list", length(batches))
  
  for (batch_num in seq_along(batches)) {
    idx <- batches[[batch_num]]
    these_resps <- valid_responses[idx]
    
    resp_text <- paste0(seq_along(these_resps), ". ", these_resps, collapse = "\n")
    instr_text <- if (!is.null(instructions)) paste0("\n\nAdditional instructions: ", instructions) else ""
    
    req <- httr2::request(paste0(base_url, "/chat/completions")) %>%
      httr2::req_headers("Content-Type" = "application/json") %>%
      httr2::req_body_json(
        list(
          model = model,
          temperature = 0,
          seed = 123,
          messages = list(
            list(
              role = "system",
              content = paste(
                "You are a survey researcher coding open-ended responses based on the detailed code descriptions.",
                "Follow these rules strictly:",
                "1. Do not skip any rows — each response must receive at least one code.",
                "2. Only use numeric codes from the provided list; no text, symbols, or letters.",
                "3. If the same code number appears more than once for a response, list it only once.",
                "4. If multiple codes apply, separate them with commas.",
                "If a response is written in a language other than English, translate it into English first."
              )
            ),
            list(
              role = "user",
              content = paste(
                "Question label: ", q_label, "\n",
                code_text,
                "\n\nResponses:\n", resp_text,
                "\n\nReturn one line per response in the same order, in the format:\n1. codes\n2. codes\n3. codes\n...",
                instr_text
              )
            )
          )
        )
      ) %>%
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
    raw_result <- NULL
    if (!is.null(result$choices) && length(result$choices) > 0) {
      raw_result <- result$choices[[1]]$message$content
    }
    if (is.null(raw_result) || raw_result == "") {
      stop(
        paste0("LLM returned empty response. Ensure Ollama is running:\n  ollama run ", model),
        call. = FALSE
      )
    }
    
    codes <- stringr::str_split(raw_result, "\n")[[1]]
    codes <- stringr::str_trim(codes)
    codes <- stringr::str_remove(codes, "^[0-9]+\\.\\s*")
    batch_results[[batch_num]] <- codes
  }
  
  coded_valid <- unlist(batch_results)
  
  # Handle length mismatch - pad or truncate to match valid_idx length
  if (length(coded_valid) < length(valid_idx)) {
    coded_valid <- c(coded_valid, rep(NA_character_, length(valid_idx) - length(coded_valid)))
  } else if (length(coded_valid) > length(valid_idx)) {
    coded_valid <- coded_valid[seq_along(valid_idx)]
  }
  
  # Insert NA back for invalid rows
  coded_raw <- rep(NA_character_, length(responses))
  coded_raw[valid_idx] <- coded_valid
  
  # Output ----
  
  # Valid codes from theme list

  valid_codes <- theme_list$Code
  
  # Clean codes - filter to only valid codes
  clean_codes <- function(x) {
    if (is.na(x) || trimws(x) == "") return(NA_character_)
    codes <- stringr::str_extract_all(x, "\\d+")[[1]]
    codes <- suppressWarnings(as.numeric(codes))
    codes <- codes[!is.na(codes)]
    codes <- codes[codes %in% valid_codes]
    if (length(codes) == 0) return(NA_character_)
    codes <- unique(codes)
    codes <- sort(codes)
    paste(codes, collapse = ", ")
  }
  
  # Get bins for codes
  get_bins <- function(c) {
    if (is.na(c) || c == "") return(NA_character_)
    codes_vec <- as.numeric(stringr::str_split(c, ",\\s*")[[1]])
    bins_vec <- theme_list$Bin[match(codes_vec, theme_list$Code)]
    bins_vec <- bins_vec[!is.na(bins_vec)]
    if (length(bins_vec) == 0) return(NA_character_)
    paste(bins_vec, collapse = "; ")
  }
  
  # Apply cleaning
  cleaned_codes <- vapply(coded_raw, clean_codes, character(1), USE.NAMES = FALSE)
  cleaned_bins <- vapply(cleaned_codes, get_bins, character(1), USE.NAMES = FALSE)
  
  # Create table
  result <- tibble::tibble(
    !!id_var := ids,
    !!q_label := responses,
    `Code(s)` = cleaned_codes,
    `Bin(s)` = cleaned_bins
  )
  
  invisible(result)
}
