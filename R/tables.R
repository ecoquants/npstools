#' Get species richness pivot table by park and year
#'
#' @param park NPS park abbreviation
#' @param year year of data to extract
#' @param xlsx optional Excel spreadsheet path to write out
#'
#' @return pivottabler::PivotTable R6 object. Render with n_spp_tbl$renderPivot()
#' @import dplyr fs pivottabler openxlsx
#' @export
#'
#' @examples
#' nps_config <- get_nps_config(here("data/nps_config.yaml"))
#' park <- "CINMS"
#' year <- 2015
#'
#' # optional path to Excel spreadsheet output
#' n_spp_xlsx <- here(glue("data/spp_richness_pivot_{park}_{year}.xlsx"))
#'
#' n_spp_tbl <- get_n_spp_tbl(park, year, xlsx=n_spp_xlsx)
#'
#' # render pivot table as html widget
#' n_spp_tbl$renderPivot()
get_n_spp_pt_tbl <- function(park, year, xlsx=NULL){

  stopifnot(exists("nps_config"))

  load_park_tables(
    nps_config, park,
    tbls=c("tbl_Phenology_Species", "tlu_Richness", "tbl_Events", "tbl_Locations", "tlu_Project_Taxa", "tlu_Layer"))

  d <- tbl_Phenology_Species %>%
    # convert to 5 m plot values
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
    summarize(
      present = max(Analysis_value)) %>%
    # summarize by transect, ie all plots
    group_by(Event_ID, Species_Code) %>%
    summarize(
      present = max(present)) %>%
    # filter by year
    left_join(
      tbl_Events %>%
        mutate(
          date = as.Date(Start_Date, "%m/%d/%Y %H:%M:%S")),
      by="Event_ID") %>%
    filter(
      year(date) == year) %>%
    # get species nativity, life form, by park
    left_join(
      tlu_Project_Taxa %>%
        select(Species_Code=Species_code, Native, Layer) %>%
        left_join(
          tlu_Layer %>%
            select(Layer=Layer_code, Life_Form=Layer_desc),
          by="Layer"),
      by="Species_Code") %>%
    # get vegetation community
    left_join(
      tbl_Locations %>%
        select(Location_ID, Vegetation_Community),
      by="Location_ID") %>%
    ungroup() %>%
    mutate(
      Nativity = recode(
        Native,
        N = "Non-native",
        Y = "Native",
        U = "Unknown")) %>%
    select(Event_ID, Species_Code, Nativity, Life_Form, Vegetation_Community)
  #table(d$Nativity) # TODO: confirm, eg CABR 2015 has Non-native:70, Unknown:4, Native:651

  calc_n_spp <- function(fxn="mean", pivotCalculator, netFilters, format, baseValues, cell){

    tbl <- pivotCalculator$getFilteredDataFrame(
      pivotCalculator$getDataFrame("d"), netFilters)%>%
      group_by(Event_ID) %>%
      summarise(n_spp = n_distinct(Species_Code)) %>%
      ungroup()
    tbl <- switch(
      fxn,
      mean = summarize(tbl, v = mean(n_spp)),
      sd   = summarize(tbl, v = sd(n_spp)),
      min  = summarize(tbl, v = min(n_spp)),
      max  = summarize(tbl, v = max(n_spp)))
    v <- pull(tbl, v)
    list(
      rawValue = v,
      formattedValue = ifelse(
        is.nan(v) | is.na(v) | is.infinite(v),
        "",
        pivotCalculator$formatValue(v, format=format)))
  }

  calc_n_spp_mean <- function(...) { calc_n_spp("mean", ...) }
  calc_n_spp_sd   <- function(...) { calc_n_spp("sd", ...) }
  calc_n_spp_min  <- function(...) { calc_n_spp("min", ...) }
  calc_n_spp_max  <- function(...) { calc_n_spp("max", ...) }

  # create the pivot table
  pt <- PivotTable$new()
  pt$addData(d, "d")
  pt$addRowDataGroups("Life_Form", totalCaption="All")
  pt$addRowDataGroups("Nativity", totalCaption="All")
  pt$addColumnDataGroups("Vegetation_Community", totalCaption="All")
  pt$defineCalculation(
    calculationName="n_spp_mean", caption="mean", calculationFunction=calc_n_spp_mean,
    format="%.1f", noDataCaption="", type="function") # noDataValue=0,
  pt$defineCalculation(
    calculationName="n_spp_sd", caption="sd", calculationFunction=calc_n_spp_sd,
    format="%.1f", noDataCaption="", type="function")
  pt$defineCalculation(
    calculationName="n_spp_min", caption="min", calculationFunction=calc_n_spp_min,
    format="%.1f", noDataCaption="", type="function")
  pt$defineCalculation(
    calculationName="n_spp_max", caption="max", calculationFunction=calc_n_spp_max,
    format="%.1f", noDataCaption="", type="function")
  pt$evaluatePivot()

  if (!is.null(xlsx)){
    wb <- createWorkbook(creator = Sys.getenv("USERNAME"))
    addWorksheet(wb, "Data")
    pt$writeToExcelWorksheet(
      wb=wb, wsName="Data", "formattedValueAsNumber",
      topRowNumber=1, leftMostColumnNumber=1, applyStyles=TRUE)
    saveWorkbook(wb, file=xlsx, overwrite = TRUE)
  }

  #pt$renderPivot()
  return(pt)
}

#' Get table of species data for given park
#'
#' @param park park abbreviation, eg "CABR", "CHIS" or "SAMO"
#'
#' @return tibble
#' @export
#'
#' @examples
#' nps_config <- get_nps_config(here("data/nps_config.yaml"))
#' get_spp_park_tbl(park = "CABR")
get_spp_park_tbl <- function(park){
  load_park_tables(
    nps_config, park,
    tbls=c("tlu_AnnualPerennial", "tlu_Nativity", "tbl_Events", "tlu_Project_Taxa", "tlu_Layer"))

  d <- tlu_AnnualPerennial %>%
    inner_join(
      tlu_Project_Taxa, by=c("AnnualPerennial_code"="Perennial")) %>%
    inner_join(
      tlu_Nativity, by=c("Native"="Nativity_code")) %>%
    inner_join(
      tlu_Layer, by=c("Layer"="Layer_code")) %>%
    filter(
      !is.null(Species_code), # TODO: filter(!is.na(Species_code)) ?
      Unit_code == park) %>% # Note: because load_park_tables(..., park) should already be limited to park
    select(
      Species_Code=Species_code, Scientific_name, Layer, FxnGroup=Layer_desc, Native, Nativity=Nativity_desc,
      Perennial=AnnualPerennial_code, AnnPer=AnnualPerennial_desc)
  d
}

