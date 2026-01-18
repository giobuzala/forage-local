# forage

Minimal Shiny app to code open-ended survey responses with a theme list.

## Run locally

1. Install packages (if needed):

```
install.packages(c("shiny", "dotenv", "readxl", "writexl", "tibble", "dplyr", "purrr", "stringr", "httr2", "rlang"))
```

2. Set your OpenAI API key:

```
Sys.setenv(OPENAI_API_KEY = "your_api_key")
```

Or create a local `.env` file in the project root:

```
OPENAI_API_KEY="your_api_key"
```

3. Start the app:

```
shiny::runApp()
```

## Files

- `app.R` Shiny UI with two steps: generate themes, then code responses.
- `Functions/code_gpt.R` and `Functions/theme_gpt.R` accept data frames or `.xlsx`/`.xls` file paths.