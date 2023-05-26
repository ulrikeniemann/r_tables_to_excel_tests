# ******************************************************************************
#
# XXX
# 
# Datenerfassung aus Excel
#
# Ulrike Niemann, April 2023
#
# ******************************************************************************
# libs -------------------------------------------------------------------------
# benötite packages laden (und wenn nötig installieren)
#
if (!require("tidyverse")) install.packages("tidyverse"); library(tidyverse)
if (!require("readxl")) install.packages("readxl"); library(readxl)
if (!require("sjlabelled")) install.packages("sjlabelled"); library(sjlabelled)
#
# ******************************************************************************
# Konstanten -------------------------------------------------------------------
#
# Environment vorsichtshalber leeren
rm(list = ls())
#
# Ordner für Rohdaten-Daten
path <- "./"
#
# Ordnerinhalt: 
#list.files(path = path, pattern = ".xlsx$")
# nur xlsx-Files mit xxx im Namen
excelFiles <- list.files(path = path, pattern = "*beispiel.*[.]xlsx$")
#
# ******************************************************************************
# einlesen ---------------------------------------------------------------------
# es sind mehrere Rohdateien einzulesen
for(i in 1:length(excelFiles)) {
  # einlesen
  sheet <- read_excel(enc2native(str_c(path, excelFiles[i])),  # teils Umlaute 
                      sheet = "Datenerhebung")
  # anfügen
  if (! exists("dat")) {
    dat <- sheet
  } else {
    dat <- bind_rows(dat, sheet)
  }
  rm(sheet)
} 
#
# Spaltennamen immer besser ohne Sonderzeichen
colnames(dat) <- 
  c("Kategorie", "F1_01", "F1_02", "F1_03", "F2_01", "F2_02")
#
# ******************************************************************************
# aufbereiten ------------------------------------------------------------------
# jetzt noch den df hübsch aufbereiten
# 
# Level setzen  / factor
# muss explizit sortiert werden wenn Reihenfolge nicht stimmig ist
# soll ja in Tabellierung richtig dargestellt werden
# (automatisch: Reihenfolge levels alphabetisch?)
# für Ja/Nein also nicht nötig ... hier nur expemplarisch
dat <- dat %>% 
  mutate(
    F1_01 = factor(F1_01) %>% fct_relevel(
      "Typ XXX",
      "Typ YYY"),
    F1_03 = factor(F1_03),
    F2_01 = factor(F2_01),
    F2_02 = factor(F2_02) %>% fct_relevel(
      "Antwort A",
      "Antwort B",
      "Antwort C",
      "Antwort D",
      "Antwort E",
      "Antwort F"
      )
  )
# Labels setzen
dat <- dat %>% 
var_labels(
  Kategorie = "Kategorie",
  F1_01 = "1.1) ... Frage ...",
  F1_02 = "1.2) ... Frage ...",
  F1_03 = "1.3) ... Frage ...",
  F2_01 = "2.1) ... Frage ...",
  F2_02 = "2.2) ... Frage ...")
# explizit NA leveln
dat <- dat %>% 
  mutate_if(is.factor, ~ fct_na_value_to_level(., "keine Angabe"))
#
# F1_02 kategorisieren
dat <- dat %>% 
  mutate(F1_02_kat = case_when(
    F1_02 == 0 ~ "0 Tage",
    F1_02 == 1 ~ "1 Tag",
    F1_02 == 2 ~ "2 Tage",
    F1_02 == 3 ~ "3 Tage",
    F1_02 == 4 ~ "4 Tage",
    F1_02 == 5 ~ "5 Tage",
    F1_02 == 6 ~ "6 Tage",
    F1_02 == 7 ~ "7 Tage",
    F1_02 > 7 ~ "mehr als 7 Tage",
    .default = "keine Angabe"
  ), .after = F1_02) %>% 
  mutate(F1_02_kat = factor(F1_02_kat) %>% fct_relevel(
    "0 Tage",
    "1 Tag",
    "2 Tage",
    "3 Tage",
    "4 Tage",
    "5 Tage",
    "6 Tage",
    "7 Tage",
    "mehr als 7 Tage",
    "keine Angabe"
  ))
# Labels setzen
dat <- dat %>% 
  var_labels(
    F1_02_kat = "1.2) Anzahl ... (kategorisiert)"
  )
table(dat$F1_02_kat, useNA = "ifany")
#
table(dat$Kategorie)
#
# Bereinigung:
# keine Angabe in F1_01 (Typ)
dat <- dat %>% filter(F1_01 != "keine Angabe")
#
# ******************************************************************************
# speichern --------------------------------------------------------------------
save(dat, file = str_c("beispiel-daten.RData"))
#
# ******************************************************************************
# ******************************************************************************
# ******************************************************************************
