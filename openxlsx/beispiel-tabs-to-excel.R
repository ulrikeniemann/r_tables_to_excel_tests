################################################################################
#
# XXX
# Excel-Tabellen, variable Spalten
#
# Ulrike Niemann, April 2023
# 
################################################################################
rm(list = ls())
# Load packages ----------------------------------------------------------------
if (!require("tidyverse")) install.packages("tidyverse"); library(tidyverse)
options(dplyr.summarise.inform = FALSE)
if (!require("janitor")) install.packages("janitor"); library(janitor)
if (!require("sjlabelled")) install.packages("sjlabelled"); library(sjlabelled)
if (!require("openxlsx")) install.packages("openxlsx"); library(openxlsx)
# data -------------------------------------------------------------------------
load(file = "./Daten/beispiel-daten.RData")
# split ------------------------------------------------------------------------
split <- "Kategorie"
# dat <- dat %>% mutate(Typ = F1_01)
# split <- "Typ"
# functions --------------------------------------------------------------------
crossTab = function(filtereddat, split, var, perc = FALSE) {
  tab <- filtereddat %>%
    group_by({{split}}, {{var}}) %>%
    summarise(n = n()) %>%
    spread({{split}}, n) %>%
    adorn_totals("row", name = "Gesamt") %>%
    adorn_totals("col", name = "Gesamt")
  if (perc == TRUE) {
    tab <- tab %>% 
      adorn_percentages("col") 
  }
  return(tab)
}
meanTab = function(filtereddat, split, var) {
  filtereddat %>%
    group_by({{split}}) %>%
    summarise(Median = median(!!sym(var), na.rm = TRUE),
              Mittelwert = mean(!!sym(var), na.rm = TRUE)) %>% 
    gather(!!var, Wert, 2:3) %>% 
    spread({{split}}, Wert) %>% 
    mutate("Gesamt" = c(median(filtereddat[[var]], na.rm = TRUE),
                        mean(filtereddat[[var]], na.rm = TRUE)))
}
# create styles
pct_st <- createStyle(numFmt = "0%")
titel_st <- createStyle(textDecoration = "Bold", fontSize = 14)
header_st <- createStyle(textDecoration = "Bold", border = "Bottom")
bold_st <- createStyle(textDecoration = "Bold")
total_row_st <- createStyle(textDecoration = "Bold", 
                            border = c("top", "bottom"), 
                            borderStyle = c("thin", "double"))
# ein sheet mit absoluten und relativen Anteilen, evtl. auch Mittelwert --------
addSheetFreqVar = function(wb, dat, var, 
                           strFilter = "", mean = FALSE) {
  sh = addWorksheet(wb, var)
  # Überschrift und ggf. Filter (erst nach Spaltenbreite!)
  writeData(wb, sh, get_label(dat)[var], startRow = 1, headerStyle = header_st)
  addStyle(wb, sh, rows = 1, cols = 1, stack = TRUE,
           style = titel_st)
  writeData(wb, sh, strFilter, startRow = 2)
  # absolute Häufigkeiten
  tab <- crossTab(dat, !!sym(split),!!sym(var))
  writeData(wb, sh, tab, startRow = 4, startCol = 2, headerStyle = header_st)
  # relative Häufigkeiten
  tab <- crossTab(dat, !!sym(split), !!sym(var), perc = TRUE)
  writeData(wb, sh, tab, 
            startRow = nrow(tab) + 9, startCol = 2,
            headerStyle = header_st)
  # bold style (erste Spalte + letzte Spalte Gesamt)
  addStyle(wb, sh, 
           cols = 2, rows = 1:(nrow(tab)*2 + 9), 
           style = bold_st, stack = TRUE)
  addStyle(wb, sh, 
           cols = ncol(tab) + 1, rows = 1:(nrow(tab)*2 + 9), 
           style = bold_st, stack = TRUE)
  # style total / Gesamt-Zeilen
  addStyle(wb, sh, 
           cols = 2:(ncol(tab) + 1), 
           rows = nrow(tab) + 4, 
           style = total_row_st, stack = TRUE)
  addStyle(wb, sh, 
           cols = 2:(ncol(tab) + 1), 
           rows = nrow(tab)*2 + 9, 
           style = total_row_st, stack = TRUE)
  # Prozent-Style
  addStyle(wb, sh, 
           style = pct_st, 
           cols = c(3:(ncol(tab) + 1)), rows = (nrow(tab) + 6):(nrow(tab)*2 + 9), 
           gridExpand = TRUE, stack = TRUE)
  # ggf Mittelwert-Tabelle
  if (mean == TRUE) {
    len <- nrow(tab)
    tab <- meanTab(dat, !!sym(split), #!!sym(var))
                   # falls kategorisierte Var übergeben wird auf Original zurück
                   str_replace(var, "_kat", ""))
    writeData(wb, sh, tab, 
              startRow = len*2 + 13, startCol = 2, headerStyle = header_st)
    addStyle(wb, sh, 
             rows = (len*2 + 15), cols = 2:(ncol(tab) + 1),
             gridExpand = TRUE, stack = TRUE,
             style = createStyle(numFmt = "0.0"))
  }
  # column width 
  setColWidths(wb, sheet = sh, cols = 1, widths = 4)
  setColWidths(wb, sheet = sh, cols = 2, widths = "auto")
}
################################################################################
wb = createWorkbook()
# ------------------------------------------------------------------------------
addSheetFreqVar(wb, dat, 
                var = "F1_01")
# ------------------------------------------------------------------------------
addSheetFreqVar(wb, dat, 
                var = "F1_02_kat", 
                mean = TRUE)
# ------------------------------------------------------------------------------
addSheetFreqVar(wb, dat, 
                var = "F1_03")
# ------------------------------------------------------------------------------
addSheetFreqVar(wb, dat %>% filter(F1_03 == "Ja"), 
                var <- "F2_01", 
                strFilter = '(Filter: 1.c) == "Ja")')
# ------------------------------------------------------------------------------
addSheetFreqVar(wb, dat %>% filter(F2_01 == "Ja"), 
                var <- "F2_02", 
                strFilter = '(Filter: 2.a) == "Ja")')
# ------------------------------------------------------------------------------
#
# hier können noch viele andere Fragen folgen...
#
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Achtung, im Zweifel wird die vorliegende gleichnamige xlsx überschrieben
saveWorkbook(wb, 
             str_c("BeispielTabellen", "_", split, ".xlsx"), 
             overwrite = TRUE)
################################################################################
################################################################################
################################################################################

