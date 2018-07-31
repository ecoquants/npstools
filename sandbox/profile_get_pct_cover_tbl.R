om <- function(){
  obj_mem <- function(x){
    #browser()
    get(x) %>% object.size() / 1024^2 #%>% format(units="Mb")
  }

  objects = ls(envir = globalenv())
  tibble(obj = objects) %>%
    mutate(
      mem_mb = map_dbl(obj, obj_mem)) %>%
    arrange(desc(mem_mb))
}

devtools::load_all()
library(tidyverse)
nps_config_yaml <- system.file(package="npstools", "nps_config.yaml")
cfg <- get_nps_config(nps_config_yaml)
park <- "CABR"
#park <- "CHIS"
year <- 2015
om()

load_park_tables(
  cfg, park,
  tbls=c(
    # inner joins
    "tbl_Sites", "tbl_Locations", "tbl_Events", "tbl_Event_Point",
    # left joins
    "tbl_Species_Data", "tlu_Condition"))


d_ep <- get_total_eventpoints_tbl(cfg, park, reload = F)

tbl_spp_park <- get_spp_park_tbl(cfg, park) # TODO: CHIS - tbl_Events, tlu_Project_Taxa not found

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
    N = 0)

#fieldNames <- names(q_0data)
# for (name in fieldNames) {
#   print(name)
# }

# If the join made a .y extention, rename it and continue with the select
# if ("SurveyYear.y" %in% fieldNames){
#   q_0data <- q_0data %>%
#     rename(SurveyYear = SurveyYear.y) %>%
#     select(SurveyYear, Park, IslandCode, SiteCode, Vegetation_Community, FxnGroup, Nativity, N)
# }else{
#   # otherwise, only  run the select
#   q_0data <- q_0data %>%
#   select(SurveyYear, Park, IslandCode, SiteCode, Vegetation_Community, FxnGroup, Nativity, N)
# }

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

om() # q_0data: 17.2, q2: 2.83, tbl_Species_Data: 2.19, tbl_Event_Point: 1.54
#})
