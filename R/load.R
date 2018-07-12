#' Read Park Tables
#'
#' @param nps_config configuration of paths, etc
#' @param park park code (ie CHIS, CABR or SAMO)
#' @param tbls character vector of tables to load (default=NULL, loads all tables)
#' @param append_park whether to append park code to name (eg tbl_Species_CHIS)
#'
#' @return Does not return anything. Loads all tables listed in the nps_config$dir_tables/park folder into the global namespace.
#' @export
#'
#' @examples
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
  if (length(tbls_missing) > 0)
    stop(glue("Table(s) not found in {nps_config$dir_tables}/[{park}|shared]: {paste(tbls_missing, collapse=', ')}"))

  for (tbl in tbls){ # csv <- csvs[1]

    dir_csv <- if_else(tbl %in% tbls_park_all, dir_park, dir_shared)
    csv <- glue("{dir_csv}/{tbl}.csv")
    if (append_park)
      tbl <- glue("{tbl}_{park}")

    df  <- read_csv(csv)

    assign(tbl, df, envir=.GlobalEnv)
  }
}

get_dir_tables <- function(nps_config){
  machine <- Sys.info()[["nodename"]]
  machine_in_config <- ifelse(machine %in% names(nps_config$dir_tables_csv), T, F)

  dir_tables <- case_when(
    machine_in_config ~ nps_config$dir_tables_csv[[machine]],
    TRUE ~ nps_config$dir_tables_csv$default)

  dir_tables
}



#' Get species richness table by park and year
#'
#' @param park not yet implemented
#' @param year not yet implemented
#'
#' @return tibble
#' @export
#'
#' @examples
get_spp_richness_table <- function(park, year){

  read_park_tables(tbls=c("tbl_Phenology_Species", "tlu_Richness"))

  tbl_Phenology_Species %>%
    select(Event_ID, Species_Code, starts_with("Plot")) %>%
    gather(plot_col, plot_val, -Event_ID, -Species_Code) %>%
    filter(plot_col != "Plot_7") %>%
    left_join(
      tlu_Richness,
      by = c("plot_val"="Richness_code")) %>%
    mutate(
      plot_num    = str_sub(plot_col, 6,6),
      plot_length = ifelse(nchar(plot_col) == 6, "5m", "1m")) %>%
    group_by(Event_ID, Species_Code, plot_num) %>%
    summarise(
      present = max(Analysis_value))
}
