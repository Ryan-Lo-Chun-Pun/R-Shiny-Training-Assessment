# Install and Load required packages
# if(!require(pacman)) install.packages("pacman")
# pacman::p_load(shiny, dplyr, ggplot2, tidyr, readxl, writexl, zip, DT)

ui <- fluidPage(

  # Title for the Shiny app
  titlePanel("R Shiny Training Assessment"),

  sidebarLayout(
    sidebarPanel(
      fileInput("upload", "Upload Excel", accept = ".xlsx"),
      textOutput("file_format"),
      sliderInput("tail_parameter", "Tail Parameter",
                  min = 0, max = 3, value = 1.1, step = 0.01),
      textOutput("tail_value"),

      # Buttons to download the outputs
      downloadButton("download_table", "Download Cumulative Table"),
      downloadButton("download_plot", "Download Plot"),
      downloadButton("download_all", "Download Table and Plot")
    ),

    mainPanel(
      textOutput("error"), # error message if file wrong
      tabsetPanel(
        tabPanel("Claim Data", DT::DTOutput("dynamic"),
                 textOutput("definitions")),
        tabPanel("Cumulative Table", tableOutput("cumulative_table")),
        tabPanel("Plot", plotOutput("cumulative_plot"))
      )
    )
  )
)

server <- function(input, output, session) {

  # Store claim data that can be edited
  claims_data <- reactiveVal(NULL)
  proxy <- DT::dataTableProxy("dynamic")

  # Validating file extension
  output$error <- renderText({
    req(input$upload)
    ext <- tools::file_ext(input$upload$name)
    if(!ext %in% "xlsx") {
      return("Error: Please upload a valid excel file")
    }
    NULL
  })

  # Show the file name, tail factor and definitions
  output$file_format <- renderText({
    req(input$upload)
    paste("Uploaded file:", input$upload$name)
  })

  output$tail_value <- renderText({
    paste0("Selected Tail Parameter = ", input$tail_parameter)
  })

  output$definitions <- renderText({
    paste(
      "Loss Year: Year claim occurred. \n",
      "Development Year: Year since the loss occurred. \n",
      "Amount of Claims Paid: Total paid for claims."
    )
  })

  # Read excel data
  observeEvent(input$upload, {
    req(input$upload)
    ext <- tools::file_ext(input$upload$name)
    req(ext == "xlsx")

    df <- read_excel(input$upload$datapath, sheet = 2, skip = 2) %>% # Read from sheet 2
      select(1:3) %>%
      setNames(c("Loss Year", "Development Year", "Claims Paid")) %>%
      mutate(
        `Loss Year` = as.numeric(`Loss Year`),
        `Development Year` = as.numeric(`Development Year`),
        `Claims Paid` = as.numeric(gsub(",", "", as.character(`Claims Paid`)))
      ) %>%
      filter(!is.na(`Loss Year`), !is.na(`Development Year`), !is.na(`Claims Paid`)) %>%
      filter(`Development Year` %in% c(1, 2, 3))

    claims_data(df)
  })

  # Editable claims table
  output$dynamic <- DT::renderDT({
    req(claims_data())

    DT::datatable(
      claims_data(),
      editable = list(target = "cell"),
      rownames = FALSE,
      selection = "none",
      options = list(pageLength = 10)
    )
  }, server = FALSE)

  # Allow edit in the claims paid column only
  observeEvent(input$dynamic_cell_edit, {
    req(claims_data())

    info <- input$dynamic_cell_edit
    df <- claims_data()

    row <- info$row
    col <- info$col + 1
    value <- suppressWarnings(as.numeric(gsub(",", "", info$value)))

    # Claims Paid column is editable
    if (col == 3 && !is.na(value)) {
      df[row, col] <- value
      claims_data(df)
      DT::replaceData(proxy, df, resetPaging = FALSE, rownames = FALSE)
    } else {
      DT::replaceData(proxy, df, resetPaging = FALSE, rownames = FALSE)
    }
  })

  # Calculate cumulative triangle
  cumulative_table_reactive <- reactive({
    req(claims_data())

    claims_data_df <- claims_data() %>%
      transmute(
        Loss_Year = as.numeric(`Loss Year`),
        Development_Year = as.numeric(`Development Year`),
        Claims_Paid = as.numeric(`Claims Paid`)
      ) %>%
      filter(!is.na(Loss_Year), !is.na(Development_Year), !is.na(Claims_Paid)) %>%
      filter(Development_Year %in% c(1, 2, 3))

    # Calculate cumulative paid claims
    cumulative_data <- claims_data_df %>%
      group_by(Loss_Year, Development_Year) %>%
      summarise(Claims_Paid = sum(Claims_Paid), .groups = "drop") %>%
      arrange(Loss_Year, Development_Year) %>%
      group_by(Loss_Year) %>%
      mutate(Cumulative_Paid_Claims = cumsum(Claims_Paid)) %>%
      ungroup()

    cumulative_table_wide <- cumulative_data %>%
      select(Loss_Year, Development_Year, Cumulative_Paid_Claims) %>%

      pivot_wider(
        names_from = Development_Year,
        values_from = Cumulative_Paid_Claims,
        names_prefix = "Year_"
      ) %>%
      arrange(Loss_Year)

    # Find development factors
    f12 <- sum(cumulative_table_wide$Year_2, na.rm = TRUE) /
      sum(cumulative_table_wide$Year_1[!is.na(cumulative_table_wide$Year_2)], na.rm = TRUE)

    f23 <- sum(cumulative_table_wide$Year_3, na.rm = TRUE) /
      sum(cumulative_table_wide$Year_2[!is.na(cumulative_table_wide$Year_3)], na.rm = TRUE)

    # Fill missing values using dev factor
    if (is.na(cumulative_table_wide$Year_2[3])) {
      cumulative_table_wide$Year_2[3] <- cumulative_table_wide$Year_1[3] * f12
    }

    if (is.na(cumulative_table_wide$Year_3[2])) {
      cumulative_table_wide$Year_3[2] <- cumulative_table_wide$Year_2[2] * f23
    }

    if (is.na(cumulative_table_wide$Year_3[3])) {
      cumulative_table_wide$Year_3[3] <- cumulative_table_wide$Year_2[3] * f23
    }

    # Use tail factor
    cumulative_table_wide$Year_4 <- cumulative_table_wide$Year_3 * input$tail_parameter

    cumulative_table_wide[is.na(cumulative_table_wide)] <- 0

    colnames(cumulative_table_wide) <- c("Loss Year", "1", "2", "3", "4")

    cumulative_table_wide <- cumulative_table_wide %>%
      mutate(across(where(is.numeric), ~ format(round(., 0), big.mark = ",", scientific = FALSE)))

    return(cumulative_table_wide)
  })

  # Display cumulative table
  output$cumulative_table <- renderTable({
    req(cumulative_table_reactive())
    cumulative_table_reactive()
  })

  # Create cumulative plot
  output$cumulative_plot <- renderPlot({
    req(cumulative_table_reactive())

    cumulative_table <- cumulative_table_reactive() %>%
      mutate(across(-`Loss Year`, ~ as.numeric(gsub(",", "", .))))

    cumulative_long <- pivot_longer(
      cumulative_table,
      cols = -`Loss Year`,
      names_to = "Development Year",
      values_to = "Cumulative Paid Claims"
    ) %>%
      mutate(`Development Year` = as.numeric(`Development Year`))

    ggplot(cumulative_long, aes(x = `Development Year`,
                                y = `Cumulative Paid Claims`,
                                color = factor(`Loss Year`),
                                group = `Loss Year`)) +
      geom_line(size = 1.0) +
      geom_point(size = 2) +
      geom_text(aes(label = format(`Cumulative Paid Claims`, big.mark = ",")),
                vjust = -0.5, size = 4, fontface = "bold") +
      labs(title = "Cumulative Paid Claims ($)",
           x = "Development Year", y = "Paid Claims",
           color = "Loss Year") +
      theme_minimal() +
      theme(legend.position = "bottom")
  })

  # Download Section
  output$download_table <- downloadHandler(
    filename = function() {
      paste0("cumulative_table_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      download_data <- cumulative_table_reactive() %>%
        mutate(across(-`Loss Year`, ~ as.numeric(gsub(",", "", .))))

      writexl::write_xlsx(download_data, file)
    }
  )

  output$download_plot <- downloadHandler(
    filename = function() {
      paste0("cumulative_plot_", Sys.Date(), ".png")
    },
    content = function(file) {
      cumulative_table <- cumulative_table_reactive() %>%
        mutate(across(-`Loss Year`, ~ as.numeric(gsub(",", "", .))))

      cumulative_long <- pivot_longer(
        cumulative_table,
        cols = -`Loss Year`,
        names_to = "Development Year",
        values_to = "Cumulative Paid Claims"
      ) %>%
        mutate(`Development Year` = as.numeric(`Development Year`))

      p <- ggplot(cumulative_long, aes(x = `Development Year`,
                                       y = `Cumulative Paid Claims`,
                                       color = factor(`Loss Year`),
                                       group = `Loss Year`)) +
        geom_line(size = 1.0) +
        geom_point(size = 2) +
        geom_text(aes(label = format(`Cumulative Paid Claims`, big.mark = ",")),
                  vjust = -0.5, size = 4, fontface = "bold") +
        labs(title = "Cumulative Paid Claims ($)",
             x = "Development Year",
             y = "Paid Claims",
             color = "Loss Year") +
        theme_minimal() +
        theme(legend.position = "bottom")

      ggsave(file, plot = p, width = 8, height = 5, dpi = 300)
    }
  )

  output$download_all <- downloadHandler(
    filename = function() {
      paste0("claims_analysis_", Sys.Date(), ".zip")
    },
    content = function(file) {
      temp_dir <- tempdir()

      table_file <- file.path(temp_dir, "cumulative_table.xlsx")
      plot_file  <- file.path(temp_dir, "cumulative_plot.png")

      table_data <- cumulative_table_reactive() %>%
        mutate(across(-`Loss Year`, ~ as.numeric(gsub(",", "", .))))

      writexl::write_xlsx(table_data, table_file)

      cumulative_long <- table_data %>%
        pivot_longer(
          cols = -`Loss Year`,
          names_to = "Development Year",
          values_to = "Cumulative Paid Claims"
        ) %>%
        mutate(`Development Year` = as.numeric(`Development Year`))

      p <- ggplot(cumulative_long,
                  aes(x = `Development Year`,
                      y = `Cumulative Paid Claims`,
                      color = factor(`Loss Year`),
                      group = `Loss Year`)) +
        geom_line(size = 1.0) +
        geom_point(size = 2) +
        geom_text(aes(label = format(`Cumulative Paid Claims`, big.mark = ",")),
                  vjust = -0.5, size = 4, fontface = "bold") +
        labs(title = "Cumulative Paid Claims ($)",
             x = "Development Year",
             y = "Paid Claims",
             color = "Loss Year") +
        theme_minimal() +
        theme(legend.position = "bottom")

      ggsave(plot_file, plot = p, width = 8, height = 5, dpi = 300)

      zip::zip(file,
               files = c(table_file, plot_file),
               mode = "cherry-pick")
    }
  )
}

shinyApp(ui, server)
