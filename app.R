library(shiny)
library(pdftools)
library(dplyr)
library(stringr)
library(DT)
library(openxlsx)

# Increase maximum upload size to 150MB (adjust as needed)
options(shiny.maxRequestSize = 150 * 1024^2)

#Helper functions
grab <- function(block, pattern) {
  m <- str_match(block, pattern)[, 2]
  if (is.na(m)) return(NA_character_)
  trimws(m)
}

clean_num <- function(x) as.numeric(str_remove_all(x, "[,%$]"))

junk_pattern <- "SET ENVIRONMENT|AMPscript"

drop_junk_pages <- function(page_text) {
  is_junk <- grepl(junk_pattern, page_text, ignore.case = TRUE)
  page_text[!is_junk]
}

parse_block <- function(block) {
  job_id <- grab(block, "Job ID:\\s*([0-9]+)")
  data_ext <- grab(block, "Data Extensions Sent Count:\\s*([^\\(]+)\\(")
  date_sent <- grab(block, "Date Sent:\\s*([0-9/]+)")
  total_sent <- grab(block, "Total Sent:\\s*([0-9,\\.]+)")
  not_sent <- grab(block, "Total Not Sent:\\s*([0-9,]+)")
  delivery_rt <- grab(block, "Delivery Rate:\\s*\\n?\\s*([0-9.]+%)")  #delivery_rt <- grab(block, "Delivery Rate:\\s*\\n?\\s*([0-9\\.]+)%")
  bounces <- grab(block, "Total Bounces:\\s*\\n?\\s*([0-9,]+)")
  hard_bounce <- grab(block, "Hard Bounce:\\s*\\n?\\s*([0-9,]+)")
  soft_bounce <- grab(block, "Soft Bounce:\\s*\\n?\\s*([0-9,\\.]+)")
  delivered <- grab(block, "Delivered:\\s*\\n?\\s*([0-9,]+)")
  open_rate <- grab(block, "Open Rate:\\s*([0-9.]+)%")
  total_opens <- grab(block, "Total Opens:\\s*([0-9,]+)")
  unique_opens <- grab(block, "Unique Opens:\\s*([0-9,]+)")
  unsubs <- grab(block, "Unsubscribes\\s*-\\n?\\s*([0-9,]+)")
  
  clicks_m <- str_match(block, "Clicks\\s*\\n?\\s*([0-9,]+)\\s*\\n?\\s*([0-9,]+)")
  total_clicks <- clicks_m[, 2]
  unique_clicks <- clicks_m[, 3]
  
  tibble(
    job_id = clean_num(job_id),
    data_extension = str_trim(data_ext),
    date_sent = suppressWarnings(as.Date(date_sent, format = "%m/%d/%Y")),
    total_sent = clean_num(total_sent),
    total_not_sent = clean_num(not_sent),
    total_bounces = clean_num(bounces),
    hard_bounce = clean_num(hard_bounce),
    soft_bounce = clean_num(soft_bounce),
    delivered = clean_num(delivered),
    delivery_rate = clean_num(delivery_rt) / 100,
    unique_opens = clean_num(unique_opens),
    open_rate = clean_num(open_rate) / 100,
    total_opens = clean_num(total_opens),
    total_clicks = clean_num(total_clicks),
    unique_clicks = clean_num(unique_clicks),
    unsubscribes = clean_num(unsubs)
  )
}

extract_pdf <- function(path) {
  page_text <- pdf_text(path)
  page_text <- drop_junk_pages(page_text) # drop AMPscript/junk pages
  full_text <- paste(page_text, collapse = "\n")
  
  blocks <- str_split(full_text, "(?=Job ID:)")[[1]]
  blocks <- blocks[str_detect(blocks, "Job ID:")]
  
  bind_rows(lapply(blocks, parse_block))
}

#-----UI---------
ui <- fluidPage(
  titlePanel("PDF Cleaner & Combiner"),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("pdf_inputs", "Select PDF Files",
                multiple = TRUE,
                accept = c(".pdf", "application/pdf")
                ),
      
      textInput("junk_pattern", "Junk Pattern (Regex)", 
                value = "(?i)SET ENVIRONMENT|AMPscript"),
      
      actionButton("process_btn", "Process & Combine PDFs", class = "btn-primary"),
      hr(),
      uiOutput("download_ui")
    ),
    
    mainPanel(
      h4("Status & Logs"),
      verbatimTextOutput("status_log")
    )
  ),
  
  titlePanel("Send Report PDF \u2192 Excel Extractor"),
  
  sidebarLayout(
    sidebarPanel(
      fileInput(
        "pdf_files", "Upload one or more send-report PDFs",
        multiple = TRUE, accept = ".pdf"
      ),
      helpText("Junk pages containing AMPscript / 'SET ENVIRONMENT' text are automatically excluded."),
      downloadButton("download_xlsx", "Download Consolidated Excel")
    ),
    mainPanel(
      DTOutput("preview_table")
    )
  )
)

#-----SERVER---------
server <- function(input, output, session) {
  
  # Reactive values to hold processed output path and logs
  rv <- reactiveValues(output_path = NULL, log = "Ready to process files.")
  
  observeEvent(input$process_btn, {
    
    req(input$pdf_inputs)
    
    rv$log <- "Processing started...\n"
    
    # Extract file details from Shiny's upload control
    uploaded_files <- input$pdf_inputs$datapath
    file_names <- input$pdf_inputs$name
    junk_pattern <- input$junk_pattern
    
    # Helper function to find non-junk pages
    get_clean_pages <- function(file_path) {
      text <- pdftools::pdf_text(file_path)
      is_junk <- grepl(junk_pattern, text, ignore.case = TRUE)
      valid_pages <- which(!is_junk)
      return(valid_pages)
    }
    
    tryCatch({
      cleaned_temp_files <- sapply(seq_along(uploaded_files), function(i) {
        f <- uploaded_files[i]
        valid_pages <- get_clean_pages(f)
        
        if (length(valid_pages) == 0) {
          rv$log <- paste0(rv$log, "Warning: All pages were marked as junk in ", file_names[i], "\n")
          return(NULL)
        }
        
        temp_out <- tempfile(fileext = ".pdf")
        pdftools::pdf_subset(f, pages = valid_pages, output = temp_out)
        return(temp_out)
      })
      
      # Filter out NULL entries if any files had 0 valid pages
      # cleaned_temp_files <- unlist(compact(cleaned_temp_files))
      cleaned_temp_files <- unlist(cleaned_temp_files[!sapply(cleaned_temp_files, is.null)])
      
      if (length(cleaned_temp_files) == 0) {
        rv$log <- paste0(rv$log, "Error: No valid pages found across all uploaded files.\n")
        return()
      }
      
      # Merge clean PDFs
      final_output <- tempfile(fileext = ".pdf")
      pdf_combine(input = cleaned_temp_files, output = final_output)
      
      # Cleanup individual page temp files
      unlink(cleaned_temp_files)
      
      rv$output_path <- final_output
      rv$log <- paste0(rv$log, "Successfully combined ", length(uploaded_files), " PDFs!\nReady for download.")
      
    }, error = function(e) {
      rv$log <- paste0(rv$log, "Error: ", e$message, "\n")
    })
  })
  
  output$status_log <- renderText({
    rv$log
  })
  
  # Show download button only after processing succeeds
  output$download_ui <- renderUI({
    req(rv$output_path)
    downloadButton("download_pdf", "Download Combined PDF", class = "btn-success")
  })
  
  output$download_pdf <- downloadHandler(
    filename = function() {
      paste0("combined_output_", Sys.Date(), ".pdf")
    },
    content = function(file) {
      file.copy(rv$output_path, file)
    }
  )
  
  #----------------------
  ##Excel extractor
  consolidated <- reactive({
    req(input$pdf_files)
    
    all_rows <- lapply(input$pdf_files$datapath, extract_pdf)
    df <- bind_rows(all_rows) %>%
      mutate(
        click_rate = unique_clicks / delivered,
        unsubscribe_rate = unsubscribes / delivered
      ) %>%
      distinct(job_id, .keep_all = TRUE) %>%
      arrange(date_sent)
    
    df
  })
  
  output$preview_table <- renderDT({
    datatable(consolidated(), options = list(scrollX = TRUE))
  })
  
  output$download_xlsx <- downloadHandler(
    filename = function() {
      paste0("Send_Reports_Consolidated_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      df <- consolidated()
      
      # -------------------------------------------------------------
      # 1. Compute Totals Row
      # -------------------------------------------------------------
      int_cols <- c("total_sent", "total_not_sent", "total_bounces", "hard_bounce",
                    "soft_bounce", "delivered", "total_opens", "unique_opens",
                    "total_clicks", "unique_clicks", "unsubscribes")
      
      # Create a blank 1-row data frame matching df structure
      totals_row <- as.data.frame(matrix(NA, nrow = 1, ncol = ncol(df)))
      colnames(totals_row) <- colnames(df)
      
      # Label the first column as "Total"
      totals_row[[1, 1]] <- "Total"
      
      # Sum integer columns
      for (col in int_cols) {
        if (col %in% names(df)) {
          totals_row[[1, col]] <- sum(df[[col]], na.rm = TRUE)
        }
      }
      # totals_row[[col]] <- prettyNum(totals_row[[col]], big.mark=",", scientific=FALSE)
      
      # format_int_cols <- function(df, cols) {
      #   for (col in cols) {
      #     if (col %in% names(df)) {
      #       df[[col]] <- prettyNum(df[[col]], big.mark = ",", scientific = FALSE)
      #     }
      #   }
      #   return(df)
      # }
      
      # Create a custom percentage function
      as_percent <- function(x, digits = 2) {
        paste0(formatC(x * 100, format = "f", digits = digits), "%")
      }
      
      # Recalculate summary percentages based on total counts
      if ("delivery_rate" %in% names(df) && "total_sent" %in% names(df)) {
        totals_row$delivery_rate <- totals_row$delivered / totals_row$total_sent
      }
      if ("open_rate" %in% names(df) && "delivered" %in% names(df)) {
        totals_row$open_rate <- totals_row$unique_opens / totals_row$delivered
      }
      if ("click_rate" %in% names(df) && "delivered" %in% names(df)) {
        totals_row$click_rate <- totals_row$unique_clicks / totals_row$delivered
      }
      if ("unsubscribe_rate" %in% names(df) && "delivered" %in% names(df)) {
        totals_row$unsubscribe_rate <- totals_row$unsubscribes / totals_row$delivered
      }
      
      totals_row$delivery_rate <- as_percent(totals_row$delivery_rate)
      totals_row$open_rate <- as_percent(totals_row$open_rate)
      totals_row$click_rate <- as_percent(totals_row$click_rate)
      totals_row$unsubscribe_rate <- as_percent(totals_row$unsubscribe_rate)
      
      # totals_row <- format_int_cols(totals_row,int_cols)
      
      
      # Combine data with Totals row
      df_combined <- rbind(df, totals_row)
      
      # df_combined <- format_int_cols(df_combined,int_cols)
      # df_combined <- df_combined %>% mutate(across(all_of(int_cols),as.integer))
      
      
      #-------------------------------------------------------------
        # 2. Build Workbook as Simple Range
      # -------------------------------------------------------------
      
      wb <- createWorkbook()
      addWorksheet(wb, "Data")
      
      # Write simple data range (headerStyle formats the header row)
      header_style <- createStyle(textDecoration = "bold", fgFill = "#4F81BD", fontColour = "#FFFFFF")
      writeData(wb, "Data", df_combined, headerStyle = header_style)
      
      # writeDataTable(wb, "Data", df_combined, tableName = "MyCustomTable", tableStyle = "TableStyleMedium1")
      
      # Format Percentage Columns (including total row)
      pct_cols <- c("delivery_rate", "open_rate", "click_rate", "unsubscribe_rate")
      
      for (col in pct_cols) {
        if (col %in% names(df_combined)) {
          # Strip % signs or character spaces if present as text, then convert to numeric
          clean_vals <- gsub("[% ,]", "", df_combined[[col]])
          num_vals <- as.numeric(clean_vals)
          
          # If values were 99.36 instead of 0.9936, divide by 100
          if (any(num_vals > 1, na.rm = TRUE)) {
            num_vals <- num_vals / 100
          }
          df_combined[[col]] <- num_vals
        }
      }
      
      # pct_idx <- match(pct_cols, names(df_combined))
      pct_idx <- as.integer(na.omit(match(pct_cols, names(df_combined))))
      if (length(pct_idx) > 0) {
      addStyle(wb, "Data", style = createStyle(numFmt = "0.00%"),
               rows = 2:(nrow(df_combined) + 2), cols = pct_idx, gridExpand = TRUE, stack=TRUE)
      }
      
      # Format Integer Columns (including total row)
      int_cols <- c("total_sent", "total_not_sent", "total_bounces", "hard_bounce",
                    "soft_bounce", "delivered", "total_opens", "unique_opens",
                    "total_clicks", "unique_clicks", "unsubscribes")
      int_idx <- match(int_cols, names(df_combined))
      addStyle(wb, "Data", style = createStyle(numFmt = "#,##0"),
               rows = 2:(nrow(df_combined) + 2), cols = int_idx, gridExpand = TRUE, stack=TRUE)
      
      # Format Total Row (Bold text + Accounting top border and double bottom border)
      # total_row_idx <- (nrow(df_combined) + 2)
      total_style <- createStyle(textDecoration = "bold", 
                                 border = "topBottom", 
                                 borderStyle = c("thin", "double"))
      
      addStyle(wb, "Data", style = total_style, 
               rows = (nrow(df_combined) + 1), cols = 1:ncol(df_combined), stack=TRUE)
      
      setColWidths(wb, "Data", cols = 1:ncol(df_combined), widths = "auto")
      freezePane(wb, "Data", firstActiveRow = 2)
      
      # wb$worksheets[[match("Data", names(wb))]]$tables <- NULL
      #Remove the table structure
      
      # tables <- getTables(wb, sheet = "Data")
      # tbl_idx <- match("MyCustomTable", tables)
      
      # if (length(table_names) > 0) {
      #   table_idx <- as.integer(which(table_names == table_names[1]))
      #   removeTable(wb, sheet = "Data", table = table_idx)
      # }
      # removeTable(wb, sheet = "Data", table ="MyCustomTable")
      # 
      # if (!is.na(tbl_idx)) {
      #   removeTable(wb, sheet = "Data", table = tbl_idx)
      # }
      
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui = ui, server = server)


