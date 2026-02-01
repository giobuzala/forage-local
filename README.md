# forage

A locally hosted, interactive Shiny app for secure, AI-assisted coding of open-ended survey responses.

## Overview

forage is an interactive Shiny application for coding open-ended survey responses into structured qualitative themes using a **locally hosted LLM via Ollama**, with all processing performed on the user’s machine.

It is designed for survey researchers and analysts who want a faster, more consistent, and more transparent way to generate code frames and apply them to large volumes of verbatim responses without losing human control over the process.

The workflow is explicitly human-in-the-loop: generated theme lists can be reviewed, edited, or replaced before coding begins.

**Data stays local**: forage is fully self-hosted. All processing occurs on your machine, and no survey data is transmitted to external APIs or cloud services.

## Data security

forage is **self-hosted and locally executed**. The Shiny app, language model, and all data processing run entirely on your machine via Ollama.

   - No external APIs – responses are never sent to OpenAI, cloud services, or third-party servers
   - No data egress – survey files, verbatims, and coded outputs remain on disk under your control
   - Network-isolatable – the app can be run offline or behind a firewall
   - Model transparency – the exact local model used for theme generation and coding is fully known and configurable

This makes forage suitable for working with **confidential, proprietary, or pre-release survey data**, and for environments with strict data-handling or client privacy requirements.

## How it works

1. Upload a survey Excel file (`.xlsx` or `.xls`) that contains:
   - a **respondent ID** column
   - an **open-ended response** column

2. **Theme generation** creates a theme list:
   - **Code** - a unique numeric code for each theme (e.g., 1, 2, 3)
   - **Bin** - short descriptive label for the theme (e.g., "Trust", "Price", "Service")
   - **Description** - a brief explanation of what the theme represents

3. **Coding responses** assigns all applicable codes from the theme list to each response. The output is an Excel file with two sheets:
   - `Coded Responses`
   - `Theme List`

This structure keeps the coding transparent and auditable, making it easy to review results, share with collaborators, or feed into downstream analysis workflows.

## Methodology

This tool uses AI to assist with generating themes and coding open-ended survey responses. The resulting themes and codes are based on probabilistic interpretations of text and may differ from how a human researcher would code the same responses.

As a best practice, generated themes should be reviewed before coding, with results validated against research objectives and human judgment applied where nuance or ambiguity exists.

This tool is designed to accelerate qualitative analysis, not to replace expert review or methodological oversight.

## Files

- `app.R` Shiny UI with two steps: generate themes, then code responses.
- `Functions/theme_gpt.R` Generates a theme list from open-ended responses.
- `Functions/code_gpt.R` Codes responses using a provided theme list.

## To run locally

1. Install R packages:

```r
install.packages(c(
  "dplyr", "httr2", "openxlsx", "purrr", "readxl", "shiny", "shinyjs", "stringr", "tibble"
))
```

2. Install and start Ollama, then pull a model:

```bash
ollama run llama3.2:3b
```

3. Start the app:

```r
shiny::runApp("app.R")
```

## License

This project is released under a **custom non-commercial license** authored by the project owner.
