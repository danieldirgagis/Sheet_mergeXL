library(shiny)
library(openxlsx)

ui <- fluidPage(
  titlePanel("Combine Excel Sheets"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file1", "Select Excel File", accept = ".xlsx"),
      br(),
      selectInput("file_type", label = "Select File Type", 
                  choices = c("Excel", "CSV"), selected = "Excel"),
      br(),
      downloadButton("download_data", "Download Combined Data")
    ),
    mainPanel(
      tableOutput("table")
    )
  )
)

server <- function(input, output) {
  
  # Function to combine the sheets in an Excel file
  combine_sheets <- function(excel_file) {
    # Load the Excel file
    wb <- loadWorkbook(excel_file)
    
    # Get the sheet names
    sheet_names <- names(wb)
    
    # Initialize an empty data frame to store the combined data
    combined_data <- data.frame()
    
    # Loop through each sheet in the Excel file
    for (sheet_name in sheet_names) {
      # Read the sheet data into a data frame
      sheet_data <- read.xlsx(wb, sheet = sheet_name, startRow = 1, colNames = TRUE, detectDates = TRUE)
      
      # Combine the sheet data with the existing data
      combined_data <- rbind(combined_data, sheet_data)
    }
    
    return(combined_data)
  }
  
  # Function to save the combined data as an Excel file
  save_as_excel <- function(data, filename) {
    write.xlsx(data, file = filename, rowNames = FALSE)
  }
  
  # Function to save the combined data as a CSV file
  save_as_csv <- function(data, filename) {
    write.csv(data, file = filename, row.names = FALSE)
  }
  
  # React to file input
  observeEvent(input$file1, {
    # Combine the sheets in the Excel file
    combined_data <- combine_sheets(input$file1$datapath)
    
    # Display the combined data in a table
    output$table <- renderTable(combined_data)
    
    # Save the combined data as either an Excel or CSV file
    if (input$file_type == "Excel") {
      filename <- paste0("combined_data.xlsx")
      save_as_excel(combined_data, filename)
    } else {
      filename <- paste0("combined_data.csv")
      save_as_csv(combined_data, filename)
    }
    
    # Provide a download button to download the combined data
    output$download_data <- downloadHandler(
      filename = function() {
        if (input$file_type == "Excel") {
          paste0("combined_data.xlsx")
        } else {
          paste0("combined_data.csv")
        }
      },
      content = function(file) {
        if (input$file_type == "Excel") {
          save_as_excel(combined_data, file)
        } else {
          save_as_csv(combined_data, file)
        }
      }
    )
  })
}

shinyApp(ui, server)
