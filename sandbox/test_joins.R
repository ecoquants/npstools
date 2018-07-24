# load libraries
library(npstools) # devtools::load_all()
library(tidyverse)
library(here)
library(glue)
library(fs)

# load your own configuration, which could be based off package
nps_config_yaml <- system.file(package="npstools", "nps_config.yaml")
cfg <- get_nps_config(nps_config_yaml)

# bigger function now
for (park in c("CABR", "SAMO")){
  get_pct_cover_tbl(cfg, park, 2015) %>%
    write_csv(here(glue("sandbox/pct_cover_{park}-2015_pre-fix.csv")))
}

get_spp_park_tbl(cfg, "CABR")
get_spp_park_tbl(cfg, "CHIS")

# specify park and year of interest
sz <- function(o){ format(object.size(o), units = "auto") }
get_total_eventpoints_tbl(cfg, "CABR") %>% sz()  #  18.2 Kb
get_total_eventpoints_tbl(cfg, "SAMO") %>% sz()  #  40.5 Kb
get_total_eventpoints_tbl(cfg, "CHIS") %>% sz()  # 300.6 Kb


