# ******************************************************************************
#
# Test pivottabler - shiny app
# Tabellenerstellung
#
# Ulrike Niemann
#
# ******************************************************************************
#
# http://www.pivottabler.org.uk/articles/v14-shiny.html
# erstes/einfachstes Beispiel
# sind auch noch umfangreichere mit multiplen Spalten.. gelistet s.o.
#
library(shiny)
library(htmlwidgets)
library(pivottabler)
#
ui <- fluidPage(
  titlePanel("Pivottabler Minimal Example Shiny App"),
  sidebarLayout(
    sidebarPanel(
      selectInput("selectRows", label = h5("Rows"),
                  choices = list("Train Category" = "TrainCategory",
                                 "TOC" = "TOC",
                                 "Power Type" = "PowerType"), selected = "TOC"),
      selectInput("selectCols", label = h5("Columns"),
                  choices = list("Train Category" = "TrainCategory",
                                 "TOC" = "TOC",
                                 "Power Type" = "PowerType"), selected = "TrainCategory")
    ),
    mainPanel(
      pivottablerOutput('pvt')
    )
  )
)
#
server <- function(input, output) {
  output$pvt <- renderPivottabler({
    pt <- PivotTable$new()
    pt$addData(bhmtrains)
    pt$addColumnDataGroups(input$selectCols)
    pt$addRowDataGroups(input$selectRows)
    pt$defineCalculation(calculationName="TotalTrains", summariseExpression="n()")
    pt$evaluatePivot()
    pivottabler(pt)
  })
}
#
shinyApp(ui = ui, server = server)
#
# ******************************************************************************