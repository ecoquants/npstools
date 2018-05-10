#' Read Park Tables
#'
#' @param nps_config configuration of paths, etc
#' @param park park code (ie CHIS, CABR or SAMO)
#' @param append_park whether to append park code to name (eg tbl_Species_CHIS)
#'
#' @return Does not return anything. Loads all tables listed in the nps_config$dir_tables/park folder into the global namespace.
#' @export
#'
#' @examples
read_park_tables <- function(nps_config, park, append_park=F){
  #park <- "CABR"; append_park <- F

  dir_csv <- file.path(nps_config$dir_tables, park)
  csvs <- list.files(dir_csv, ".*\\.csv")

  for (csv in csvs){ # csv <- csvs[1]
    tbl <- path_ext_remove(csv)
    if (append_park)
      tbl <- glue("{tbl}_{park}")

    df  <- read_csv(file.path(dir_csv, csv))

    assign(tbl, df, envir=.GlobalEnv)
  }
}


#' Get species richness plots
#'
#' @param park not yet implemented
#' @param year not yet implemented
#'
#' @return tibble
#' @export
#'
#' @examples
get_spp_richness_plots <- function(park, year){
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
