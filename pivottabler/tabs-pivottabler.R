# ******************************************************************************
#
# Test pivottabler
# Tabellenerstellung
#
# Ulrike Niemann
#
# ******************************************************************************
#
# http://www.pivottabler.org.uk/
# 
# A simple example of creating a pivot table - summarising the types of trains run by different train companies:
#
library(pivottabler)
# arguments:  qpvt(dataFrame, rows, columns, calculations, ...)
qpvt(bhmtrains, "TOC", "TrainCategory", "n()") # TOC = Train Operating Company 
#
# The equivalent verbose commands to output the same pivot table as above are:
pt <- PivotTable$new()
pt$addData(bhmtrains) # bhmtrains is a data frame with columns TrainCategory, TOC, etc.
pt$addColumnDataGroups("TrainCategory") # e.g. Express Passenger
pt$addRowDataGroups("TOC") # TOC = Train Operating Company e.g. Arriva Trains Wales
pt$defineCalculation(calculationName="TotalTrains", summariseExpression="n()")
pt$evaluatePivot()
pt
#
# Multiple levels can be added to the pivot table row or column headings, e.g. looking at combinations of TOC and PowerType:
#
qpvt(bhmtrains, c("TOC", "PowerType"), "TrainCategory", "n()")
#
pt <- PivotTable$new()
pt$addData(bhmtrains)
pt$addColumnDataGroups("TrainCategory")
pt$addRowDataGroups("TOC")
pt$addRowDataGroups("PowerType") # D/EMU = Diesel/Electric Multiple Unit, HST=High Speed Train
pt$defineCalculation(calculationName="TotalTrains", summariseExpression="n()")
pt$evaluatePivot()
pt
#
# HTML Output
pt$renderPivot()
#
# Outline layout is an alternative way of rendering the row groups, e.g. for the same pivot table as above:
#
pt <- PivotTable$new()
pt$addData(bhmtrains) 
pt$addColumnDataGroups("TrainCategory") 
pt$addRowDataGroups("TOC", 
                    outlineBefore=list(isEmpty=FALSE, groupStyleDeclarations=list(color="blue")), 
                    outlineTotal=list(isEmpty=FALSE, groupStyleDeclarations=list(color="blue"))) 
pt$addRowDataGroups("PowerType", addTotal=FALSE) 
pt$defineCalculation(calculationName="TotalTrains", summariseExpression="n()")
pt$renderPivot()
#
#
# Excel Output
# The same styling/formatting used for the HTML output is also used when outputting to Excel - greatly reducing the amount of script that needs to be written to create Excel output.
#
pt <- PivotTable$new()
pt$addData(bhmtrains) # bhmtrains is a data frame with columns TrainCategory, TOC, etc.
pt$addColumnDataGroups("TrainCategory") # e.g. Express Passenger
pt$addRowDataGroups("TOC") # TOC = Train Operating Company e.g. Arriva Trains Wales
pt$addRowDataGroups("PowerType") # D/EMU = Diesel/Electric Multiple Unit, HST=High Speed Train
pt$defineCalculation(calculationName="TotalTrains", summariseExpression="n()")
pt$evaluatePivot()

library(openxlsx)
wb <- createWorkbook(creator = Sys.getenv("USERNAME"))
addWorksheet(wb, "Data")
pt$writeToExcelWorksheet(wb=wb, wsName="Data", 
                         topRowNumber=2, 
                         leftMostColumnNumber=2, 
                         applyStyles=TRUE)
#
# column width auto size
# https://rdrr.io/cran/openxlsx/man/setColWidths.html
setColWidths(wb, sheet = 1, cols = 1:5, widths = "auto")
#
saveWorkbook(wb, file = "test.xlsx", overwrite = TRUE)
#
# ******************************************************************************