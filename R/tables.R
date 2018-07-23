#' Get pivot table of species richness by park and year
#'
#' @param cfg NPS configuration list object; see \code{\link{get_nps_config}}
#' @param park NPS park abbreviation
#' @param year year of data to extract
#' @param xlsx optional Excel spreadsheet path to write out
#'
#' @return pivottabler::PivotTable R6 object. Render with n_spp_tbl$renderPivot()
#' @import dplyr fs pivottabler openxlsx
#' @export
#'
#' @examples
#' cfg  <- get_nps_config(system.file(package="npstools", "nps_config.yaml"))
#' park <- "CABR"
#' year <- 2015
#'
#' n_spp_pivtbl <- get_n_spp_pivtbl(cfg, park, year)
#'
#' # render pivot table as html widget
#' n_spp_pivtbl$renderPivot()
get_n_spp_pivtbl <- function(cfg, park, year, xlsx=NULL){

  load_park_tables(
    cfg, park,
    tbls=c("tbl_Phenology_Species", "tlu_Richness", "tbl_Events", "tbl_Locations", "tlu_Project_Taxa", "tlu_Layer"))

  # change to a character field so the left join will work
  if (park == "CHIS"){
    tbl_Events$Event_ID <- as.character(tbl_Events$Event_ID)
  }

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
#' @param cfg NPS configuration list object; see \code{\link{get_nps_config}}
#' @param park park abbreviation, eg "CABR", "CHIS" or "SAMO"
#'
#' @return tibble with fields: Species_Code, Scientific_name, Layer, FxnGroup,
#'   Native, Nativity, Perennial, AnnPer
#' @export
#'
#' @examples
#' cfg <- get_nps_config(system.file(package="npstools", "nps_config.yaml"))
#' get_spp_park_tbl(cfg, park = "CABR")
get_spp_park_tbl <- function(cfg, park){
  load_park_tables(
    cfg, park,
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

#' Get table of total event points for given park
#'
#' @param cfg NPS configuration list object; see \code{\link{get_nps_config}}
#' @param park park abbreviation, eg "CABR", "CHIS" or "SAMO"
#'
#' @return tibble with fields: Park, IslandCode, Location_ID, SiteCode,
#'   Vegetation_Community, SurveyYear, SurveyDate, NofPoints
#' @export
#' @examples
#' cfg <- get_nps_config(system.file(package="npstools", "nps_config.yaml"))
#' get_total_eventpoints_tbl(cfg, park)
get_total_eventpoints_tbl <- function(cfg, park){
  # VB: mod_ExportQueries.TotalPointsSQL(iPark As Integer) [L202]
  # park <- "CHIS"

  load_park_tables(cfg, park, c("tbl_Sites", "tbl_Locations", "tbl_Events", "tbl_Event_Point"))

  d_ep <- tbl_Sites %>%
    inner_join(
      tbl_Locations %>% select(-Unit_Code), by="Site_ID") %>%
    inner_join(
      tbl_Events %>% select(-Analysis_code), by="Location_ID") %>%
    inner_join(
      tbl_Event_Point, by="Event_ID") %>%
    mutate(
      start_date = lubridate::as_date(
        Start_Date, tz="America/Los_Angeles", format = "%m/%d/%Y %H:%M:%S"),
      SurveyYear = lubridate::year(start_date) %>% as.integer()) %>%
    # VB: ...LocTypeFilter(), HAVING tbl_Sites.Unit_Code = "ParkName(iPark)"
    filter(
      Unit_Code == park,
      Loc_Type == "I&M",
      Monitoring_Status == "Active") %>%
    #names() %>% sort()
    select(
      Park=Unit_Code, IslandCode=Site_Name, Location_ID, SiteCode=Location_Code,
      Vegetation_Community, SurveyYear, SurveyDate=Start_Date, Point_No) %>%
    group_by(
      Park, IslandCode, Location_ID, SiteCode, Vegetation_Community, SurveyYear, SurveyDate) %>%
    summarize(
      NofPoints = n_distinct(Point_No)) # TODO: check is Count(tbl_Event_Point.Point_No) AS NofPoints
  d_ep
}

#' Get table of absolute percent cover for given park and year
#'
#' @param cfg NPS configuration list object; see \code{\link{get_nps_config}}
#' @param park park abbreviation, eg "CABR", "CHIS" or "SAMO"
#' @param year 4-digit year
#'
#' @return Tibble that reproduces from \href{
#'       https://github.com/ecoquants/npstools/blob/3ca70ac9704a4a11d6d5d34f707e3008e35d0a35/inst/accdb_source/mod_ExportQueries.vb}{
#'       mod_ExportQueries}:
#'
#' \itemize{
#'   \item{
#'     \href{
#'       https://github.com/ecoquants/npstools/blob/3ca70ac9704a4a11d6d5d34f707e3008e35d0a35/inst/accdb_source/mod_ExportQueries.vb#L1225-L1289}{
#'       Export_AnnualReport_AbsoluteCover()}}}
#'
#'for "Figure E.2. Absolute foliar cover (\%) of plant growth forms, as observed during 20XX monitoring at CABR. Colored bars show mean values, while error bars extend ±1 s.d. from the means." from MEDN_veg_protocol_NARRATIVE_FINAL_8Sep2016.pdf.
#'
#' @export
#' @examples
#' cfg  <- get_nps_config(system.file(package="npstools", "nps_config.yaml"))
#' park <- "CABR"
#' year <- 2015
#'
#' get_pct_cover_tbl(cfg, park, year)
get_pct_cover_tbl <- function(cfg, park, year){
  # year?
  # VB: mod_ExportQueries.Export_AnnualReport_AbsoluteCover()

  tbl_spp_park <- get_spp_park_tbl(cfg, park) # TODO: CHIS - tbl_Events, tlu_Project_Taxa not found

  load_park_tables(
    cfg, park,
    tbls=c(
      # inner joins
      "tbl_Sites", "tbl_Locations", "tbl_Events", "tbl_Event_Point",
      # left joins
      "tbl_Species_Data", "tlu_Condition"))

  d_ep <- get_total_eventpoints_tbl(cfg, park)

  # VB: ...strRaw =
  d <- tbl_Sites %>%
    inner_join(
      tbl_Locations %>% select(-Unit_Code), by="Site_ID") %>%
    inner_join(
      tbl_Events %>% select(-Analysis_code), by="Location_ID") %>%
    inner_join(
      tbl_Event_Point, by="Event_ID") %>%
    left_join(
      tbl_Species_Data, by="Event_Point_ID") %>%
    left_join(
      tlu_Condition, by="Condition") %>%
    left_join(
      tbl_spp_park, by=c("Species_Code")) %>% # TODO: consider to_lower() or fix column names
    # VB: ...LocTypeFilter()
    filter(
      Unit_Code == park,
      Loc_Type == "I&M",
      Monitoring_Status == "Active") %>%
    # VB: ...strWhere =
    mutate(
      start_date = lubridate::as_date(
        Start_Date, tz="America/Los_Angeles", format = "%m/%d/%Y %H:%M:%S"),
      SurveyYear = lubridate::year(start_date) %>% as.integer()) %>%
    filter(
      SurveyYear == year,
      is.null(Analysis_code) || Analysis_code == "Alive") %>%
    select(
      SurveyYear, Park = Unit_Code, IslandCode = Site_Name, SiteCode = Location_Code, Vegetation_Community,
      Species_Code, Condition = Analysis_code, FxnGroup, Nativity)

  # VB: ...strRawSum =
  d_sum <- d %>%
    group_by(SurveyYear, Park, IslandCode, SiteCode, Vegetation_Community, FxnGroup, Nativity) %>%
    summarize(
      N = n_distinct(Species_Code)) # TODO: confirm same as SQL: Count(qRaw.Species_Code) AS N

  # VB: ...str1 =
  q1 <- tbl_Sites %>%
    inner_join(
      tbl_Locations %>% select(-Unit_Code), by="Site_ID") %>%
    inner_join(
      tbl_Events %>% select(-Analysis_code), by="Location_ID")  %>%
    # VB: ...LocTypeFilter()
    filter(
      Unit_Code == park,
      Loc_Type == "I&M",
      Monitoring_Status == "Active") %>%
    # VB: year
    mutate(
      start_date = lubridate::as_date(
        Start_Date, tz="America/Los_Angeles", format = "%m/%d/%Y %H:%M:%S"),
      SurveyYear = lubridate::year(start_date) %>% as.integer()) %>%
    filter(
      SurveyYear == year) %>%
    # select
    select(SurveyYear, Park=Unit_Code, IslandCode=Site_Name, SiteCode=Location_Code, Vegetation_Community)

  # VB: ...str1 =
  q2 <- tbl_Sites %>%
    inner_join(
      tbl_Locations %>% select(-Unit_Code), by="Site_ID") %>%
    inner_join(
      tbl_Events %>% select(-Analysis_code), by="Location_ID") %>%
    inner_join(
      tbl_Event_Point, by="Event_ID") %>%
    left_join(
      tbl_Species_Data, by="Event_Point_ID") %>%
    left_join(
      tlu_Condition, by="Condition") %>%
    left_join(
      tbl_spp_park, by=c("Species_Code")) %>%
    # VB: ...strWhere =
    mutate(
      start_date = lubridate::as_date(
        Start_Date, tz="America/Los_Angeles", format = "%m/%d/%Y %H:%M:%S")) %>%
    filter(
      lubridate::year(start_date) == year,
      is.null(Analysis_code) || Analysis_code == "Alive")

  # VB: ...str0Data =
  q_0data <- q1 %>%
    full_join(q2, by="Vegetation_Community") %>% # TODO: confirm CROSS JOIN by="Vegetation_Community"
    mutate(
      N = 0) %>%
    select(SurveyYear, Park, IslandCode, SiteCode, Vegetation_Community, FxnGroup, Nativity, N)

  # VB: ...strData = strRawSum + str0Data
  q_data <- q_0data %>%
    bind_rows(
      d_sum) %>%
    group_by(SurveyYear, Park, IslandCode, SiteCode, Vegetation_Community, FxnGroup, Nativity) %>%
    summarize(
      SumOfN = sum(N))

  # VB: ...strAbsCovData = Calculating Absolute Cover (Figure E2)
  q_abscovdata <- q_data %>%
    inner_join(
      d_ep, by = c("SurveyYear", "Park", "IslandCode", "SiteCode", "Vegetation_Community")) %>%
    # TODO: fix +Vegetation_Community in VBA
    mutate(
      AbsCover = SumOfN/NofPoints * 100)

  # VB: ...strAbsCov =
  if (park == "CHIS"){
    q_strAbsCov <- q_abscovdata %>%
      group_by(SurveyYear, Park, IslandCode, Vegetation_Community, FxnGroup, Nativity)

  } else {
    q_strAbsCov <- q_abscovdata %>%
      group_by(SurveyYear, Park, Vegetation_Community, FxnGroup, Nativity)
  }
  q_strAbsCov <- q_strAbsCov %>%
    summarise(
      NofTransects = n_distinct(SiteCode),
      Average      = mean(AbsCover, na.rm=T),
      StdDev       = sd(AbsCover, na.rm=T),
      MinRange     = min(AbsCover, na.rm=T),
      MaxRange     = max(AbsCover, na.rm=T)) %>%
    mutate(
      Query_type = "Annual Report, Absolute Cover (Fig. E2)")

  q_strAbsCov
}
