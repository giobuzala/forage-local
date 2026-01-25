# forage

forage is a Shiny app that turns open-ended survey responses into structured codes using OpenAI.  
It supports two steps: **theme generation** and **coding responses**.

## Hosted app

This project is published on shinyapps.io.  

```
https://giobuzala.shinyapps.io/forage/
```

## How it works

1. Upload a survey Excel file (`.xlsx` or `.xls`) that contains:
   - an **ID** column
   - an **open-ended response** column
2. **Theme generation** creates a theme list (Code, Bin, Description).
3. **Coding** assigns one or more codes to each response.
4. The output is an Excel file with two sheets:
   - `Coded responses`
   - `Theme list`

## Run locally

1. Install packages:

```
install.packages(c(
  "shiny", "shinyjs", "dotenv", "tidyverse", "openxlsx", "readxl",
  "writexl", "tibble", "dplyr", "purrr", "stringr", "httr2", "rlang"
))
```

2. Set your OpenAI API key (one of these methods):

```
Sys.setenv(OPENAI_API_KEY = "your_api_key")
```

Or create a local `.env` file in the project root:

```
OPENAI_API_KEY="your_api_key"
```

3. Start the app:

```
shiny::runApp("App")
```

## Files

- `App/app.R` Shiny UI with two steps: generate themes, then code responses.
- `Functions/theme_gpt.R` Generates a theme list from open-ended responses.
- `Functions/code_gpt.R` Codes responses using a provided theme list.