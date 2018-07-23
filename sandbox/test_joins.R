# load libraries
library(npstools) # devtools::load_all()
library(tidyverse)
library(here)
library(glue)
library(fs)

# load your own configuration, which could be based off package
nps_config_yaml <- system.file(package="npstools", "nps_config.yaml")
cfg <- get_nps_config(nps_config_yaml)

# specify park and year of interest
park <- "CABR" # "CABR" | "CHIS" | "SABO"
year <- 2015

cfg$dir_R_tables_csv
