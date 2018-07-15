#' Read Park Tables
#'
#' @param nps_config configuration of paths, etc
#' @param park park code (ie CHIS, CABR or SAMO)
#' @param tbls character vector of tables to load (default=NULL, loads all tables)
#' @param append_park whether to append park code to name (eg tbl_Species_CHIS)
#'
#' @return Does not return anything. Loads all tables listed in the nps_config$dir_tables/park folder into the global namespace.
#' @importFrom readr read_csv
#' @importFrom glue glue
#' @importFrom glue collapse
#' @importFrom tidyr gather
#' @importFrom stringr str_sub
#' @importFrom lubridate year
#'
#' @export
#'
#' @examples
#' nps_config <- read_yaml(here("data/nps_config.yaml"))
#' load_park_tables(nps_config, "CINMS", tbls=c("tbl_Phenology_Species", "tlu_Richness"))
#'
load_park_tables <- function(nps_config, park, tbls=NULL, append_park=F){
  #park <- "CABR"; append_park <- F

  dir_tables <- get_dir_tables(nps_config)

  dir_park        <- file.path(dir_tables, park)
  dir_shared      <- file.path(dir_tables, "shared")
  csvs_park       <- list.files(dir_park, ".*\\.csv")
  csvs_shared     <- list.files(dir_shared, ".*\\.csv")
  tbls_park_all   <- path_ext_remove(csvs_park)
  tbls_shared_all <- path_ext_remove(csvs_shared)

  if (is.null(tbls))
    tbls <- tbls_park_all
  tbls_missing <- setdiff(tbls, c(tbls_park_all, tbls_shared_all))
  #browser()
  if (length(tbls_missing) > 0){
    msg <- glue("Table(s) not found in {dir_tables}/[{park}|shared]: {collapse(tbls_missing, sep=', ')}")
    stop(msg)
  }

  for (tbl in tbls){ # csv <- csvs[1]

    dir_csv <- if_else(tbl %in% tbls_park_all, dir_park, dir_shared)
    csv <- glue("{dir_csv}/{tbl}.csv")
    if (append_park)
      tbl <- glue("{tbl}_{park}")

    df  <- read_csv(csv)

    assign(tbl, df, envir=.GlobalEnv)
  }
}

#' Read NPS configuration file
#'
#' @param nps_config_yaml NPS configuration file in YAML format
#'
#' @return list object from reading in the NPS configuration file
#' @importFrom yaml read_yaml
#' @export
#'
#' @examples
#' get_nps_config(here("data/nps_config.yaml"))
get_nps_config <- function(nps_config_yaml){
  read_yaml(nps_config_yaml)
}

#' Get directory of tables with CSV's for R
#'
#' @param nps_config NPS configuration list object
#'
#' @return path to directory of tables with CSV's for R. Evaluates the machine
#'   name that my be inserted into the NPS configuration file, per: \code{Sys.info()[["nodename"]]}
#' @export
#'
#' @examples
#' nps_config <- read_yaml(here("data/nps_config.yaml"))
#' get_dir_tables(nps_config)
get_dir_tables <- function(nps_config){
  machine <- Sys.info()[["nodename"]]
  machine_in_config <- ifelse(machine %in% names(nps_config$dir_R_tables_csv), T, F)
  dir_machine <- ifelse(machine_in_config, nps_config$dir_R_tables_csv[[machine]], "")

  #browser()
  dir_tables <- case_when(
    #machine_in_config ~ nps_config$dir_R_tables_csv[[machine]],
    machine_in_config ~ dir_machine,
    TRUE ~ nps_config$dir_R_tables_csv$default)

  dir_tables
}
