# ******************************************************************************
#
# Test expss
# Tabellenerstellung
#
# Ulrike Niemann
#
# ******************************************************************************
#
# https://cran.r-project.org/web/packages/expss/vignettes/xlsx-export.html
#
library(expss)
library(openxlsx)
data(mtcars)
mtcars = apply_labels(mtcars,
                      mpg = "Miles/(US) gallon",
                      cyl = "Number of cylinders",
                      disp = "Displacement (cu.in.)",
                      hp = "Gross horsepower",
                      drat = "Rear axle ratio",
                      wt = "Weight (lb/1000)",
                      qsec = "1/4 mile time",
                      vs = "Engine",
                      vs = c("V-engine" = 0,
                             "Straight engine" = 1),
                      am = "Transmission",
                      am = c("Automatic" = 0,
                             "Manual"=1),
                      gear = "Number of forward gears",
                      carb = "Number of carburetors"
)
#
mtcars_table = mtcars %>% 
  cross_cpct(
    cell_vars = list(cyl, gear),
    col_vars = list(total(), am, vs)
  ) %>% 
  set_caption("Table 1")
#
mtcars_table
#
# Then we create workbook and add worksheet to it.
#
wb = createWorkbook()
sh = addWorksheet(wb, "Tables")
#
# Export - we should specify workbook and worksheet.
#
xl_write(mtcars_table, wb, sh)
#
# And, finally, we save workbook with table to the xlsx file.
saveWorkbook(wb, "table1.xlsx", overwrite = TRUE)
#
# ******************************************************************************
# Automation of the report generation
# First of all, we create banner which we will use for all our tables.
# 
banner = with(mtcars, list(total(), am, vs))
# Then we generate list with all tables. If variables have small number of discrete values we create column percent table. In other cases we calculate table with means. For both types of tables we mark significant differencies between groups.
#
list_of_tables = lapply(mtcars, function(variable) {
  if(length(unique(variable))<7){
    cro_cpct(variable, banner) %>% significance_cpct()
  } else {
    # if number of unique values greater than seven we calculate mean
    cro_mean_sd_n(variable, banner) %>% significance_means()
  }
})
# Create workbook:
#
wb = createWorkbook()
sh = addWorksheet(wb, "Tables")
#
# Here we export our list with tables with additional formatting. We remove ‘#’ sign from totals and mark total column with bold. You can read about formatting options in the manual fro xl_write (?xl_write in the console).
#
xl_write(list_of_tables, wb, sh, 
         # remove '#' sign from totals 
         col_symbols_to_remove = "#",
         row_symbols_to_remove = "#",
         # format total column as bold
         other_col_labels_formats = list("#" = createStyle(textDecoration = "bold")),
         other_cols_formats = list("#" = createStyle(textDecoration = "bold")),
)
#
# column width auto size
# https://rdrr.io/cran/openxlsx/man/setColWidths.html
setColWidths(wb, sheet = 1, cols = 1:7, widths = "auto")
# Save workbook:
#  
saveWorkbook(wb, "report.xlsx", overwrite = TRUE)
#
# 
# ******************************************************************************