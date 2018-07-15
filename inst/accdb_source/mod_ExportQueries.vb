Option Compare Database
Option Explicit

Dim strSQL As String

Function FilterString(frm As Form, iPark As Integer) As Variant
'SQL filter string for park/island, site, vegetation_community, and year

    Dim ParkFilter As String
    Dim IslandFilter As String
    Dim SiteFilter As String
    Dim VegetationFilter As String
    Dim YearFilter As String

    Dim strFilter As String

    ' Write the SQL filter string-----------------------------------------
    ' Park and Island-----------------------------------------------------
    Select Case iPark
    Case Is = 1 'CABR
        ParkFilter = "((Park)=" & Chr$(34) & "CABR" & Chr$(34) & ")"
        IslandFilter = ""
    Case Is = 2 'CHIS
        ParkFilter = "((Park)=" & Chr$(34) & "CHIS" & Chr$(34) & ")"

        ' Island
        Select Case frm!grpIslandFilter
        Case Is = 1
            IslandFilter = ""
        Case Is = 2
            If frm!lstIsland.ItemsSelected.Count = 0 Then
                IslandFilter = ""
            Else
                IslandFilter = " AND (((IslandCode)=" & ListToString(frm, "lstIsland", "IslandCode", "Text") & "))"
            End If
        End Select
    Case Is = 3 'SAMO
        ParkFilter = "((Park)=" & Chr$(34) & "SAMO" & Chr$(34) & ")"
        IslandFilter = ""
    End Select
    '---------------------------------------------------------------------
    ' Site/Transect-------------------------------------------------------
    Select Case frm!cmbQuery
    Case Is = 12, 13
        SiteFilter = ""
    Case Else
        Select Case frm!grpSiteFilter
        Case Is = 1
            SiteFilter = ""
        Case Is = 2
            If frm!lstSite.ItemsSelected.Count = 0 Then
                SiteFilter = ""
            Else
                SiteFilter = " AND (((Location_ID) IN (" & ListToStringIN(frm, "lstSite", "Numeric") & ")))"
            End If
        End Select
    End Select
    '---------------------------------------------------------------------
    ' Vegetation----------------------------------------------------------
    Select Case frm!grpVegetationFilter
    Case Is = 1
        VegetationFilter = ""
    Case Is = 2
        If frm!lstVegetation.ItemsSelected.Count = 0 Then
            VegetationFilter = ""
        Else
            VegetationFilter = " AND (((Vegetation_Community) IN (" & ListToStringIN(frm, "lstVegetation", "Text") & ")))"
        End If
    End Select
    '---------------------------------------------------------------------
    ' Year----------------------------------------------------------------
    Select Case frm!grpYearFilter
    Case Is = 1
        YearFilter = ""
    Case Is = 2
        If IsNothing(frm!cmbYear) Then
            YearFilter = ""
        Else
            YearFilter = " AND ((SurveyYear)=" & frm!cmbYear & ")"
        End If
    Case Is = 3
        If IsNothing(frm!cmbStartYear) Then
            MsgBox "No start year entered. All years will be used.", vbInformation
            YearFilter = ""
        Else
            If IsNothing(frm!cmbEndYear) Then
                MsgBox "No start year entered. All years will be used.", vbInformation
                YearFilter = ""
            Else
                'evaluate years
                If frm!cmbStartYear > frm!cmbEndYear Then
                    MsgBox "Start year cannot be begin after the end year." & vbNewLine & "All years will be used", vbInformation
                    YearFilter = ""
                ElseIf frm!cmbStartYear <= frm!cmbEndYear Then
                    YearFilter = " AND ((SurveyYear) BETWEEN " & frm!cmbStartYear & " AND " & frm!cmbEndYear & ")"
                End If
            End If
        End If
    End Select
    '---------------------------------------------------------------------
    ' Make the SQL filter string
    strFilter = ParkFilter & IslandFilter & SiteFilter & VegetationFilter & YearFilter

    FilterString = strFilter

End Function

Function ParkName(iPark As Integer)
'SQL value for park name
Select Case iPark
Case Is = 1 ' cabr
    ParkName = "CABR"
Case Is = 2 'chis
    ParkName = "CHIS"
Case Is = 3 'samo
    ParkName = "SAMO"
End Select

End Function

Function ParkSelect(iPark As Integer)
'SQL select statement for park

Select Case iPark
Case Is = 1 'cabr
    ParkSelect = "qry.Park"
Case Is = 2 'chis
    ParkSelect = "qry.Park, qry.IslandCode"
Case Is = 3 'samo
    ParkSelect = "qry.Park"
End Select

End Function

Function ParkSpeciesSQL(iPark As Integer)
'SQL statement for park species list

ParkSpeciesSQL = "SELECT tlu_Project_Taxa.Species_Code, tlu_Project_Taxa.Scientific_name, tlu_Project_Taxa.Layer, tlu_Layer.Layer_desc AS FxnGroup, tlu_Project_Taxa.Native, " & _
    "tlu_Nativity.Nativity_desc AS Nativity, tlu_Project_Taxa.Perennial, tlu_AnnualPerennial.AnnualPerennial_desc AS AnnPer " & _
    "FROM tlu_AnnualPerennial INNER JOIN (tlu_Nativity INNER JOIN (tlu_Project_Taxa INNER JOIN tlu_Layer ON tlu_Project_Taxa.Layer = tlu_Layer.Layer_code) ON tlu_Nativity.Nativity_code = tlu_Project_Taxa.Native) " & _
    "ON tlu_AnnualPerennial.AnnualPerennial_code = tlu_Project_Taxa.Perennial " & _
    "WHERE (((tlu_Project_Taxa.Species_code) Is Not Null) AND ((tlu_Project_Taxa.Unit_code)=" & Chr$(34) & ParkName(iPark) & Chr$(34) & "))"

End Function

Function LocTypeFilter(iPark As Integer)

LocTypeFilter = "((tbl_Sites.Unit_Code)=" & Chr$(34) & ParkName(iPark) & Chr$(34) & ") AND ((tbl_Locations.Loc_Type)=" & Chr$(34) & "I&M" & Chr$(34) & ") " & _
    "AND ((tbl_Locations.Monitoring_Status)=" & Chr$(34) & "Active" & Chr$(34) & ")"

End Function

Function ExpectedVisitsSQL(iPark As Integer)
'SQL statement for expected visits

Dim strYear As String
Dim strDate As String

strYear = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear FROM tbl_Events, tbl_Sites INNER JOIN tbl_Locations ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(iPark) & " AND ((Year([Start_Date])) Is Not Null))"

strDate = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, Year([Start_Date]) AS SurveyYear, " & _
    "tbl_Events.Start_Date AS SurveyDate FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(iPark) & ")"

ExpectedVisitsSQL = "SELECT qryYear.Park, qryYear.IslandCode, qryYear.Location_ID, qryYear.SiteCode, qryYear.Vegetation_Community, qryYear.SurveyYear, qryDate.SurveyDate " & _
    "FROM (" & strYear & ") AS qryYear LEFT JOIN (" & strDate & ") AS qryDate " & _
    "ON (qryYear.SurveyYear = qryDate.SurveyYear) AND (qryYear.SiteCode = qryDate.SiteCode) AND (qryYear.IslandCode = qryDate.IslandCode) AND (qryYear.Park = qryDate.Park)"

End Function

Function TotalHitsSQL(iPark As Integer)
'SQL statement for total hits, includes both live and dead hits

TotalHitsSQL = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year(tbl_Events.Start_Date) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Sum(IIf(IsNull([Park_Spp.Species_code]),0,1)) AS NofHits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(iPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(iPark) & ") " & _
    "GROUP BY tbl_Sites.Unit_Code, tbl_Sites.Site_Name, tbl_Locations.Location_ID, tbl_Locations.Location_Code, tbl_Locations.Vegetation_Community, Year(tbl_Events.Start_Date), tbl_Events.Start_Date"
    '"HAVING (((tbl_Sites.Unit_Code) = " & Chr$(34) & ParkName(iPark) & Chr$(34) & "))"

End Function

Function TotalLiveHitsSQL(iPark As Integer)
'SQL statement for total hits, includes only live hits

TotalLiveHitsSQL = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year(tbl_Events.Start_Date) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Sum(IIf(IsNull([Park_Spp.Species_code]),0,1)) AS NofHits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(iPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(iPark) & " AND ((tbl_Species_Data.Condition) Is Null Or (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & ")) " & _
    "GROUP BY tbl_Sites.Unit_Code, tbl_Sites.Site_Name, tbl_Locations.Location_ID, tbl_Locations.Location_Code, tbl_Locations.Vegetation_Community, Year(tbl_Events.Start_Date), " & _
    "tbl_Events.Start_Date, tbl_Species_Data.Condition"
    '"HAVING (((tbl_Sites.Unit_Code) = " & Chr$(34) & ParkName(iPark) & Chr$(34) & "))"

End Function

Function TotalPointsSQL(iPark As Integer)
'SQL statement for total points

TotalPointsSQL = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year(tbl_Events.Start_Date) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Count(tbl_Event_Point.Point_No) AS NofPoints " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Event_Point ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(iPark) & ") " & _
    "GROUP BY tbl_Sites.Unit_Code, tbl_Sites.Site_Name, tbl_Locations.Location_ID, tbl_Locations.Location_Code, tbl_Locations.Vegetation_Community, tbl_Events.Start_Date, Year(tbl_Events.Start_Date)"
    '"HAVING (((tbl_Sites.Unit_Code) = " & Chr$(34) & ParkName(iPark) & Chr$(34) & "))"

End Function

Function TotalSppObservedSQL(iPark As Integer)
'SQL statement for total number of species observed, includes both live and dead species

Dim strA As String

strA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Park_Spp.Species_Code, IIf(IsNull([Park_Spp.Species_code]),0,1) AS N " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(iPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(iPark) & ")"

TotalSppObservedSQL = "SELECT A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate, Sum(A.N) AS NofSppObserved " & _
    "FROM (" & strA & ") AS A " & _
    "GROUP BY A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate"

End Function

Function TotalLiveSppObservedSQL(iPark As Integer)
'SQL statement for total number of species observed, includes only live species

Dim strA As String

strA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Park_Spp.Species_Code, IIf(IsNull([Park_Spp.Species_code]),0,1) AS N " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data " & _
    "INNER JOIN (" & ParkSpeciesSQL(iPark) & ") AS Park_Spp ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(iPark) & " AND ((tbl_Species_Data.Species_Code) Is Not Null) AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

TotalLiveSppObservedSQL = "SELECT A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate, Sum(A.N) AS NofSppObserved " & _
    "FROM (" & strA & ") AS A " & _
    "GROUP BY A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate"

End Function

Function Export_Native_Cover(xPark As Integer)
'Query #1: Absolute and relative cover of live native species over time by park/island, vegetation community type, and transect

Dim strObsA As String
Dim strNativeObs As String

Dim strPtsA As String
Dim strNativePoints As String

Dim strHitsA As String
Dim strNativeHits As String

Dim strNativeCover As String

strObsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Park_Spp.Species_code, IIf([Native]=" & Chr$(34) & "Y" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data " & _
    "INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) " & _
    "ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Native)=" & Chr$(34) & "Y" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strNativeObs = "SELECT ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate, Nz(Sum(ObsA.Hits),0) AS NofNativesObserved " & _
    "FROM (" & strObsA & ") AS ObsA GROUP BY ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate"

strPtsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, IIf([Native]=" & Chr$(34) & "Y" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data " & _
    "INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Native)=" & Chr$(34) & "Y" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strNativePoints = "SELECT PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate, Nz(Sum(PtsA.Hits),0) AS NofNativePoints " & _
    "FROM (" & strPtsA & ") AS PtsA " & _
    "GROUP BY PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate"

strHitsA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, IIf([Native]=" & Chr$(34) & "Y" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data " & _
    "INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Native)=" & Chr$(34) & "Y" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strNativeHits = "SELECT HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate, Nz(Sum(HitsA.Hits),0) AS NofNativeHits " & _
    "FROM (" & strHitsA & ") AS HitsA " & _
    "GROUP BY HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate"

strNativeCover = "SELECT ExpectedVisits.Park, ExpectedVisits.IslandCode, ExpectedVisits.Location_ID, ExpectedVisits.SiteCode, ExpectedVisits.Vegetation_Community, " & _
    "ExpectedVisits.SurveyYear, ExpectedVisits.SurveyDate, NativeObs.NofNativesObserved, TotalSppObserved.NofSppObserved, NativePoints.NofNativePoints, TotalPoints.NofPoints, " & _
    "NativeHits.NofNativeHits, TotalHits.NofHits, [NofNativePoints]/[NofPoints] AS AbsoluteNativeCover, [NofNativeHits]/[NofHits] AS RelativeNativeCover " & _
    "FROM ((((((" & ExpectedVisitsSQL(xPark) & ") AS ExpectedVisits " & _
    "LEFT JOIN (" & strNativeObs & ") AS NativeObs ON (ExpectedVisits.Park = NativeObs.Park) AND (ExpectedVisits.IslandCode = NativeObs.IslandCode) AND (ExpectedVisits.SiteCode = NativeObs.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = NativeObs.Vegetation_Community) AND (ExpectedVisits.SurveyYear = NativeObs.SurveyYear) AND (ExpectedVisits.SurveyDate = NativeObs.SurveyDate)) " & _
    "LEFT JOIN (" & TotalLiveSppObservedSQL(xPark) & ") AS TotalSppObserved ON (ExpectedVisits.Park = TotalSppObserved.Park) AND (ExpectedVisits.IslandCode = TotalSppObserved.IslandCode) AND (ExpectedVisits.SiteCode = TotalSppObserved.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalSppObserved.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalSppObserved.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalSppObserved.SurveyDate)) " & _
    "LEFT JOIN (" & strNativePoints & ") AS NativePoints ON (ExpectedVisits.Park = NativePoints.Park) AND (ExpectedVisits.IslandCode = NativePoints.IslandCode) AND (ExpectedVisits.SiteCode = NativePoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = NativePoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = NativePoints.SurveyYear) AND (ExpectedVisits.SurveyDate = NativePoints.SurveyDate)) " & _
    "LEFT JOIN (" & TotalPointsSQL(xPark) & ") AS TotalPoints ON (ExpectedVisits.Park = TotalPoints.Park) AND (ExpectedVisits.IslandCode = TotalPoints.IslandCode) AND (ExpectedVisits.SiteCode = TotalPoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalPoints.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalPoints.SurveyDate)) " & _
    "LEFT JOIN (" & strNativeHits & ") AS NativeHits ON (ExpectedVisits.Park = NativeHits.Park) AND (ExpectedVisits.IslandCode = NativeHits.IslandCode) AND (ExpectedVisits.SiteCode = NativeHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = NativeHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = NativeHits.SurveyYear) AND (ExpectedVisits.SurveyDate = NativeHits.SurveyDate)) " & _
    "LEFT JOIN (" & TotalLiveHitsSQL(xPark) & ") AS TotalHits ON (ExpectedVisits.Park = TotalHits.Park) AND (ExpectedVisits.IslandCode = TotalHits.IslandCode) AND (ExpectedVisits.SiteCode = TotalHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalHits.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalHits.SurveyDate)"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, " & _
    "CInt(qry.NofNativesObserved) AS NObserved, qry.NofSppObserved, qry.NofPoints, " & _
    "qry.AbsoluteNativeCover AS AbsoluteCover, qry.RelativeNativeCover AS RelativeCover, " & Chr$(34) & "Native_cover" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strNativeCover & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Native_Cover = strSQL

End Function

Function Export_Nonnative_Cover(xPark As Integer)
'Query #2: Absolute and relative cover of live native species over time by park/island, vegetation community type, and transect

Dim strObsA As String
Dim strNonnativeObs As String

Dim strPtsA As String
Dim strNonnativePoints As String

Dim strHitsA As String
Dim strNonnativeHits As String

Dim strNonnativeCover As String

strObsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Park_Spp.Species_code, IIf([Native]=" & Chr$(34) & "N" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Native)=" & Chr$(34) & "N" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strNonnativeObs = "SELECT ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate, Nz(Sum(ObsA.Hits),0) AS NofNonnativesObserved " & _
    "FROM (" & strObsA & ") AS ObsA " & _
    "GROUP BY ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate"

strPtsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, IIf([Native]=" & Chr$(34) & "N" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Native)=" & Chr$(34) & "N" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strNonnativePoints = "SELECT PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate, Nz(Sum(PtsA.Hits),0) AS NofNonnativePoints " & _
    "FROM (" & strPtsA & ") AS PtsA " & _
    "GROUP BY PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate"

strHitsA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, IIf([Native]=" & Chr$(34) & "N" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Native)=" & Chr$(34) & "N" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strNonnativeHits = "SELECT HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate, Nz(Sum(HitsA.Hits),0) AS NofNonnativeHits " & _
    "FROM (" & strHitsA & ") AS HitsA " & _
    "GROUP BY HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate"

strNonnativeCover = "SELECT ExpectedVisits.Park, ExpectedVisits.IslandCode, ExpectedVisits.Location_ID, ExpectedVisits.SiteCode, ExpectedVisits.Vegetation_Community, " & _
    "ExpectedVisits.SurveyYear, ExpectedVisits.SurveyDate, NonnativeObs.NofNonnativesObserved, TotalSppObserved.NofSppObserved, NonnativePoints.NofNonnativePoints, TotalPoints.NofPoints, " & _
    "NonnativeHits.NofNonnativeHits, TotalHits.NofHits, [NofNonnativePoints]/[NofPoints] AS AbsoluteNonnativeCover, [NofNonnativeHits]/[NofHits] AS RelativeNonnativeCover " & _
    "FROM ((((((" & ExpectedVisitsSQL(xPark) & ") AS ExpectedVisits " & _
    "LEFT JOIN (" & strNonnativeObs & ") AS NonnativeObs ON (ExpectedVisits.Park = NonnativeObs.Park) AND (ExpectedVisits.IslandCode = NonnativeObs.IslandCode) AND (ExpectedVisits.SiteCode = NonnativeObs.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = NonnativeObs.Vegetation_Community) AND (ExpectedVisits.SurveyYear = NonnativeObs.SurveyYear) AND (ExpectedVisits.SurveyDate = NonnativeObs.SurveyDate)) " & _
    "LEFT JOIN (" & TotalLiveSppObservedSQL(xPark) & ") AS TotalSppObserved ON (ExpectedVisits.Park = TotalSppObserved.Park) AND (ExpectedVisits.IslandCode = TotalSppObserved.IslandCode) AND (ExpectedVisits.SiteCode = TotalSppObserved.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalSppObserved.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalSppObserved.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalSppObserved.SurveyDate)) " & _
    "LEFT JOIN (" & strNonnativePoints & ") AS NonnativePoints ON (ExpectedVisits.Park = NonnativePoints.Park) AND (ExpectedVisits.IslandCode = NonnativePoints.IslandCode) AND (ExpectedVisits.SiteCode = NonnativePoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = NonnativePoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = NonnativePoints.SurveyYear) AND (ExpectedVisits.SurveyDate = NonnativePoints.SurveyDate)) " & _
    "LEFT JOIN (" & TotalPointsSQL(xPark) & ") AS TotalPoints ON (ExpectedVisits.Park = TotalPoints.Park) AND (ExpectedVisits.IslandCode = TotalPoints.IslandCode) AND (ExpectedVisits.SiteCode = TotalPoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalPoints.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalPoints.SurveyDate)) " & _
    "LEFT JOIN (" & strNonnativeHits & ") AS NonnativeHits ON (ExpectedVisits.Park = NonnativeHits.Park) AND (ExpectedVisits.IslandCode = NonnativeHits.IslandCode) AND (ExpectedVisits.SiteCode = NonnativeHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = NonnativeHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = NonnativeHits.SurveyYear) AND (ExpectedVisits.SurveyDate = NonnativeHits.SurveyDate)) " & _
    "LEFT JOIN (" & TotalLiveHitsSQL(xPark) & ") AS TotalHits ON (ExpectedVisits.Park = TotalHits.Park) AND (ExpectedVisits.IslandCode = TotalHits.IslandCode) AND (ExpectedVisits.SiteCode = TotalHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalHits.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalHits.SurveyDate)"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, CInt(qry.NofNonnativesObserved) AS NObserved, qry.NofSppObserved, qry.NofPoints, " & _
    "qry.AbsoluteNonnativeCover AS AbsoluteCover, qry.RelativeNonnativeCover AS RelativeCover, " & Chr$(34) & "Nonnative_cover" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strNonnativeCover & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Nonnative_Cover = strSQL

End Function

Function Export_Total_Hits(xPark As Integer)
'Query #3: Number of recorded live "hits" by vegetation community type and transect over time

Dim strTotalHits As String

strTotalHits = "SELECT qPoints.Park, qPoints.IslandCode, qPoints.Location_ID, qPoints.SiteCode, qPoints.Vegetation_Community, qPoints.SurveyYear, qPoints.SurveyDate, qPoints.NofPoints, qHits.NofHits " & _
    "FROM (" & TotalHitsSQL(xPark) & ") AS qHits INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qPoints ON (qHits.SurveyDate = qPoints.SurveyDate) " & _
    "AND (qHits.SurveyYear = qPoints.SurveyYear) AND (qHits.Vegetation_Community = qPoints.Vegetation_Community) AND (qHits.SiteCode = qPoints.SiteCode) " & _
    "AND (qHits.IslandCode = qPoints.IslandCode) AND (qHits.Park = qPoints.Park)"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qHits.NofHits AS TotalNoOfHits, qry.NofPoints AS TotalPoints, " & _
    Chr$(34) & "Total_hits" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strTotalHits & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Total_Hits = strSQL

End Function

Function Export_Species_Frequency(xPark As Integer)
'Query #4: Frequency of a particular species occurring on a transect or within a plant community over time

Dim strA As String
Dim strSppFrequency As String

strA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Park_Spp.Species_code, Park_Spp.Scientific_name, 1 AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Native)<>" & Chr$(34) & "-" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null Or (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strSppFrequency = "SELECT A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate, A.Species_code, A.Scientific_name, Sum(A.Hits) AS Frequency " & _
    "FROM (" & strA & ") AS A " & _
    "GROUP BY A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate, A.Species_code, A.Scientific_name"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Species_code, qry.Scientific_name, qry.Frequency, " & _
    Chr$(34) & "Species_frequency" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strSppFrequency & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Species_Frequency = strSQL

End Function

Function Export_Annual_Cover(xPark As Integer)
'Query #5: Absolute and relative cover of live annual species over time by park/island, vegetation community type, and transect

Dim strObsA As String
Dim strAnnualObs As String

Dim strPtsA As String
Dim strAnnualPoints As String

Dim strHitsA As String
Dim strAnnualHits As String

Dim strAnnualCover As String

strObsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Park_Spp.Species_code, IIf([Perennial]=" & Chr$(34) & "N" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Perennial)=" & Chr$(34) & "N" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strAnnualObs = "SELECT ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate, Nz(Sum(ObsA.Hits),0) AS NofAnnualsObserved " & _
    "FROM (" & strObsA & ") AS ObsA " & _
    "GROUP BY ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate"

strPtsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, IIf([Perennial]=" & Chr$(34) & "N" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Perennial)=" & Chr$(34) & "N" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strAnnualPoints = "SELECT PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate, Nz(Sum(PtsA.Hits),0) AS NofAnnualPoints " & _
    "FROM (" & strPtsA & ") AS PtsA " & _
    "GROUP BY PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate"

strHitsA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, IIf([Perennial]=" & Chr$(34) & "N" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Perennial)=" & Chr$(34) & "N" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strAnnualHits = "SELECT HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate, Nz(Sum(HitsA.Hits),0) AS NofAnnualHits " & _
    "FROM (" & strHitsA & ") AS HitsA " & _
    "GROUP BY HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate"

strAnnualCover = "SELECT ExpectedVisits.Park, ExpectedVisits.IslandCode, ExpectedVisits.Location_ID, ExpectedVisits.SiteCode, ExpectedVisits.Vegetation_Community, " & _
    "ExpectedVisits.SurveyYear, ExpectedVisits.SurveyDate, AnnualObs.NofAnnualsObserved, TotalSppObserved.NofSppObserved, AnnualPoints.NofAnnualPoints, TotalPoints.NofPoints, " & _
    "AnnualHits.NofAnnualHits, TotalHits.NofHits, [NofAnnualPoints]/[NofPoints] AS AbsoluteAnnualCover, [NofAnnualHits]/[NofHits] AS RelativeAnnualCover " & _
    "FROM ((((((" & ExpectedVisitsSQL(xPark) & ") AS ExpectedVisits " & _
    "LEFT JOIN (" & strAnnualObs & ") AS AnnualObs ON (ExpectedVisits.Park = AnnualObs.Park) AND (ExpectedVisits.IslandCode = AnnualObs.IslandCode) AND (ExpectedVisits.SiteCode = AnnualObs.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = AnnualObs.Vegetation_Community) AND (ExpectedVisits.SurveyYear = AnnualObs.SurveyYear) AND (ExpectedVisits.SurveyDate = AnnualObs.SurveyDate)) " & _
    "LEFT JOIN (" & TotalLiveSppObservedSQL(xPark) & ") AS TotalSppObserved ON (ExpectedVisits.Park = TotalSppObserved.Park) AND (ExpectedVisits.IslandCode = TotalSppObserved.IslandCode) AND (ExpectedVisits.SiteCode = TotalSppObserved.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalSppObserved.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalSppObserved.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalSppObserved.SurveyDate)) " & _
    "LEFT JOIN (" & strAnnualPoints & ") AS AnnualPoints ON (ExpectedVisits.Park = AnnualPoints.Park) AND (ExpectedVisits.IslandCode = AnnualPoints.IslandCode) AND (ExpectedVisits.SiteCode = AnnualPoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = AnnualPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = AnnualPoints.SurveyYear) AND (ExpectedVisits.SurveyDate = AnnualPoints.SurveyDate)) " & _
    "LEFT JOIN (" & TotalPointsSQL(xPark) & ") AS TotalPoints ON (ExpectedVisits.Park = TotalPoints.Park) AND (ExpectedVisits.IslandCode = TotalPoints.IslandCode) AND (ExpectedVisits.SiteCode = TotalPoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalPoints.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalPoints.SurveyDate)) " & _
    "LEFT JOIN (" & strAnnualHits & ") AS AnnualHits ON (ExpectedVisits.Park = AnnualHits.Park) AND (ExpectedVisits.IslandCode = AnnualHits.IslandCode) AND (ExpectedVisits.SiteCode = AnnualHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = AnnualHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = AnnualHits.SurveyYear) AND (ExpectedVisits.SurveyDate = AnnualHits.SurveyDate)) " & _
    "LEFT JOIN (" & TotalLiveHitsSQL(xPark) & ") AS TotalHits ON (ExpectedVisits.Park = TotalHits.Park) AND (ExpectedVisits.IslandCode = TotalHits.IslandCode) AND (ExpectedVisits.SiteCode = TotalHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalHits.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalHits.SurveyDate)"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, " & _
    "CInt(qry.NofAnnualsObserved) AS NObserved, qry.NofSppObserved, qry.NofPoints, " & _
    "qry.AbsoluteAnnualCover AS AbsoluteCover, qry.RelativeAnnualCover AS RelativeCover, " & Chr$(34) & "Annual_cover" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strAnnualCover & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Annual_Cover = strSQL

End Function

Function Export_Perennial_Cover(xPark As Integer)
'Query #6: Absolute and relative cover of live perennial species over time by park/island, vegetation community type, and transect

Dim strObsA As String
Dim strPerennialObs As String

Dim strPtsA As String
Dim strPerennialPoints As String

Dim strHitsA As String
Dim strPerennialHits As String

Dim strPerennialCover As String

strObsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Park_Spp.Species_code, IIf([Perennial]=" & Chr$(34) & "Y" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Perennial)=" & Chr$(34) & "Y" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strPerennialObs = "SELECT ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate, Nz(Sum(ObsA.Hits),0) AS NofPerennialsObserved " & _
    "FROM (" & strObsA & ") AS ObsA " & _
    "GROUP BY ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate"

strPtsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, IIf([Perennial]=" & Chr$(34) & "Y" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Perennial)=" & Chr$(34) & "Y" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strPerennialPoints = "SELECT PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate, Nz(Sum(PtsA.Hits),0) AS NofPerennialPoints " & _
    "FROM (" & strPtsA & ") AS PtsA " & _
    "GROUP BY PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate"

strHitsA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, IIf([Perennial]=" & Chr$(34) & "Y" & Chr$(34) & ",1,0) AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Park_Spp.Perennial)=" & Chr$(34) & "Y" & Chr$(34) & ") AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strPerennialHits = "SELECT HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate, Nz(Sum(HitsA.Hits),0) AS NofPerennialHits " & _
    "FROM (" & strHitsA & ") AS HitsA " & _
    "GROUP BY HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate"

strPerennialCover = "SELECT ExpectedVisits.Park, ExpectedVisits.IslandCode, ExpectedVisits.Location_ID, ExpectedVisits.SiteCode, ExpectedVisits.Vegetation_Community, " & _
    "ExpectedVisits.SurveyYear, ExpectedVisits.SurveyDate, PerennialObs.NofPerennialsObserved, TotalSppObserved.NofSppObserved, PerennialPoints.NofPerennialPoints, TotalPoints.NofPoints, " & _
    "PerennialHits.NofPerennialHits, TotalHits.NofHits, [NofPerennialPoints]/[NofPoints] AS AbsolutePerennialCover, [NofPerennialHits]/[NofHits] AS RelativePerennialCover " & _
    "FROM ((((((" & ExpectedVisitsSQL(xPark) & ") AS ExpectedVisits " & _
    "LEFT JOIN (" & strPerennialObs & ") AS PerennialObs ON (ExpectedVisits.Park = PerennialObs.Park) AND (ExpectedVisits.IslandCode = PerennialObs.IslandCode) AND (ExpectedVisits.SiteCode = PerennialObs.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = PerennialObs.Vegetation_Community) AND (ExpectedVisits.SurveyYear = PerennialObs.SurveyYear) AND (ExpectedVisits.SurveyDate = PerennialObs.SurveyDate)) " & _
    "LEFT JOIN (" & TotalLiveSppObservedSQL(xPark) & ") AS TotalSppObserved ON (ExpectedVisits.Park = TotalSppObserved.Park) AND (ExpectedVisits.IslandCode = TotalSppObserved.IslandCode) AND (ExpectedVisits.SiteCode = TotalSppObserved.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalSppObserved.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalSppObserved.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalSppObserved.SurveyDate)) " & _
    "LEFT JOIN (" & strPerennialPoints & ") AS PerennialPoints ON (ExpectedVisits.Park = PerennialPoints.Park) AND (ExpectedVisits.IslandCode = PerennialPoints.IslandCode) AND (ExpectedVisits.SiteCode = PerennialPoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = PerennialPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = PerennialPoints.SurveyYear) AND (ExpectedVisits.SurveyDate = PerennialPoints.SurveyDate)) " & _
    "LEFT JOIN (" & TotalPointsSQL(xPark) & ") AS TotalPoints ON (ExpectedVisits.Park = TotalPoints.Park) AND (ExpectedVisits.IslandCode = TotalPoints.IslandCode) AND (ExpectedVisits.SiteCode = TotalPoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalPoints.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalPoints.SurveyDate)) " & _
    "LEFT JOIN (" & strPerennialHits & ") AS PerennialHits ON (ExpectedVisits.Park = PerennialHits.Park) AND (ExpectedVisits.IslandCode = PerennialHits.IslandCode) AND (ExpectedVisits.SiteCode = PerennialHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = PerennialHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = PerennialHits.SurveyYear) AND (ExpectedVisits.SurveyDate = PerennialHits.SurveyDate)) " & _
    "LEFT JOIN (" & TotalLiveHitsSQL(xPark) & ") AS TotalHits ON (ExpectedVisits.Park = TotalHits.Park) AND (ExpectedVisits.IslandCode = TotalHits.IslandCode) AND (ExpectedVisits.SiteCode = TotalHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalHits.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalHits.SurveyDate)"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, " & _
    "CInt(qry.NofPerennialsObserved) AS NObserved, qry.NofSppObserved, qry.NofPoints, " & _
    "qry.AbsolutePerennialCover AS AbsoluteCover, qry.RelativePerennialCover AS RelativeCover, " & Chr$(34) & "Perennial_cover" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strPerennialCover & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Perennial_Cover = strSQL

End Function

Function Export_Avg_Height(xPark As Integer)
'Query #7: Average recorded heights for species over time

Dim strMaxRank As String
Dim strA As String
Dim strAvgHeights As String

strMaxRank = "SELECT tbl_Event_Point.Event_Point_ID, Max(tbl_Species_Data.Species_Rank) AS MaxOfSpecies_Rank FROM tbl_Event_Point INNER JOIN tbl_Species_Data " & _
    "ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID GROUP BY tbl_Event_Point.Event_Point_ID"

strA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, tbl_Event_Point.MaxHeight_cm, Park_Spp.Species_code, Park_Spp.Scientific_Name " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN ((tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) INNER JOIN (" & strMaxRank & ") AS MaxRank ON (tbl_Species_Data.Event_Point_ID = MaxRank.Event_Point_ID) " & _
    "AND (tbl_Species_Data.Species_Rank = MaxRank.MaxOfSpecies_Rank)) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Condition) Is Null Or (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & ")) " & _
    "GROUP BY tbl_Sites.Unit_Code, tbl_Sites.Site_Name, tbl_Locations.Location_ID, tbl_Locations.Location_Code, tbl_Locations.Vegetation_Community, Year([Start_Date]), tbl_Events.Start_Date, " & _
    "tbl_Event_Point.Point_No, tbl_Event_Point.MaxHeight_cm, Park_Spp.Species_code, Park_Spp.Scientific_Name, tbl_Species_Data.Condition, tbl_Event_Point.MaxHeight_cm " & _
    "HAVING (((Max(tbl_Species_Data.Species_Rank))<>False) AND ((tbl_Event_Point.MaxHeight_cm) Is Not Null))"

strAvgHeights = "SELECT A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate, A.Species_code, A.Scientific_Name, " & _
    "Avg(A.MaxHeight_cm) AS MeanHeight, StDev(A.MaxHeight_cm) AS StdDev, Count(A.MaxHeight_cm) AS N, StDev([MaxHeight_cm])/Count([MaxHeight_cm]) AS StdErr, " & _
    "Min(A.MaxHeight_cm) AS MinRange, Max(A.MaxHeight_cm) AS MaxRange " & _
    "FROM (" & strA & ") AS A " & _
    "GROUP BY A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate, A.Species_code, A.Scientific_Name"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Species_code, qry.Scientific_name, " & _
    "qry.MeanHeight, qry.StdDev, qry.N, qry.StdErr, qry.MinRange, qry.MaxRange, " & Chr$(34) & "Avg_height" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strAvgHeights & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Avg_Height = strSQL

End Function

Function Export_Live_Cover(xPark As Integer)
'Query #8: Absolute and relative cover of live vegetation by species, park/island, vegetation community type, and transect over time

Dim strObsA As String
Dim strLiveObs As String

Dim strPtsA As String
Dim strLivePoints As String

Dim strHitsA As String
Dim strLiveHits As String

Dim strLiveCover As String

strObsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Park_Spp.Species_code, 1 AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strLiveObs = "SELECT ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate, Nz(Sum(ObsA.Hits),0) AS NofLiveObserved " & _
    "FROM (" & strObsA & ") AS ObsA " & _
    "GROUP BY ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate"

strPtsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, 1 AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Condition) Is Null OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strLivePoints = "SELECT PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate, Nz(Sum(PtsA.Hits),0) AS NofLivePoints " & _
    "FROM (" & strPtsA & ") AS PtsA " & _
    "GROUP BY PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate"

strHitsA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, 1 AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Condition) Is Null Or (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & "))"

strLiveHits = "SELECT HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate, Nz(Sum(HitsA.Hits),0) AS NofLiveHits " & _
    "FROM (" & strHitsA & ") AS HitsA " & _
    "GROUP BY HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate"

strLiveCover = "SELECT ExpectedVisits.Park, ExpectedVisits.IslandCode, ExpectedVisits.Location_ID, ExpectedVisits.SiteCode, ExpectedVisits.Vegetation_Community, " & _
    "ExpectedVisits.SurveyYear, ExpectedVisits.SurveyDate, LiveObs.NofLiveObserved, TotalSppObserved.NofSppObserved, LivePoints.NofLivePoints, TotalPoints.NofPoints, " & _
    "LiveHits.NofLiveHits, TotalHits.NofHits, [NofLiveHits]/[NofPoints] AS AbsoluteLiveCover, [NofLiveHits]/[NofHits] AS RelativeLiveCover " & _
    "FROM ((((((" & ExpectedVisitsSQL(xPark) & ") AS ExpectedVisits " & _
    "LEFT JOIN (" & strLiveObs & ") AS LiveObs ON (ExpectedVisits.Park = LiveObs.Park) AND (ExpectedVisits.IslandCode = LiveObs.IslandCode) AND (ExpectedVisits.SiteCode = LiveObs.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = LiveObs.Vegetation_Community) AND (ExpectedVisits.SurveyYear = LiveObs.SurveyYear) AND (ExpectedVisits.SurveyDate = LiveObs.SurveyDate)) " & _
    "LEFT JOIN (" & TotalSppObservedSQL(xPark) & ") AS TotalSppObserved ON (ExpectedVisits.Park = TotalSppObserved.Park) AND (ExpectedVisits.IslandCode = TotalSppObserved.IslandCode) AND (ExpectedVisits.SiteCode = TotalSppObserved.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalSppObserved.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalSppObserved.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalSppObserved.SurveyDate)) " & _
    "LEFT JOIN (" & strLivePoints & ") AS LivePoints ON (ExpectedVisits.Park = LivePoints.Park) AND (ExpectedVisits.IslandCode = LivePoints.IslandCode) AND (ExpectedVisits.SiteCode = LivePoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = LivePoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = LivePoints.SurveyYear) AND (ExpectedVisits.SurveyDate = LivePoints.SurveyDate)) " & _
    "LEFT JOIN (" & TotalPointsSQL(xPark) & ") AS TotalPoints ON (ExpectedVisits.Park = TotalPoints.Park) AND (ExpectedVisits.IslandCode = TotalPoints.IslandCode) AND (ExpectedVisits.SiteCode = TotalPoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalPoints.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalPoints.SurveyDate)) " & _
    "LEFT JOIN (" & strLiveHits & ") AS LiveHits ON (ExpectedVisits.Park = LiveHits.Park) AND (ExpectedVisits.IslandCode = LiveHits.IslandCode) AND (ExpectedVisits.SiteCode = LiveHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = LiveHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = LiveHits.SurveyYear) AND (ExpectedVisits.SurveyDate = LiveHits.SurveyDate)) " & _
    "LEFT JOIN (" & TotalHitsSQL(xPark) & ") AS TotalHits ON (ExpectedVisits.Park = TotalHits.Park) AND (ExpectedVisits.IslandCode = TotalHits.IslandCode) AND (ExpectedVisits.SiteCode = TotalHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalHits.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalHits.SurveyDate)"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, " & _
    "CInt(qry.NofLiveObserved) AS NObserved, qry.NofSppObserved, qry.NofPoints, " & _
    "qry.AbsoluteLiveCover  AS AbsoluteCover, qry.RelativeLiveCover AS RelativeCover, " & Chr$(34) & "Live_cover" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strLiveCover & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Live_Cover = strSQL

End Function

Function Export_Dead_Cover(xPark As Integer)
'Query #9: Absolute and relative cover of dead vegetation by species, park/island, vegetation community type, and transect over time

Dim strObsA As String
Dim strDeadObs As String

Dim strPtsA As String
Dim strDeadPoints As String

Dim strHitsA As String
Dim strDeadHits As String

Dim strDeadCover As String

strObsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Park_Spp.Species_code, 1 AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Condition)=" & Chr$(34) & "Dead" & Chr$(34) & " OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Dead Plant" & Chr$(34) & " Or (tbl_Species_Data.Condition)=" & Chr$(34) & "Dead Branch" & Chr$(34) & "))"

strDeadObs = "SELECT ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate, Nz(Sum(ObsA.Hits),0) AS NofDeadObserved " & _
    "FROM (" & strObsA & ") AS ObsA " & _
    "GROUP BY ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate"

strPtsA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, 1 AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Condition)=" & Chr$(34) & "Dead" & Chr$(34) & " OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Dead Plant" & Chr$(34) & " Or (tbl_Species_Data.Condition)=" & Chr$(34) & "Dead Branch" & Chr$(34) & "))"

strDeadPoints = "SELECT PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate, Nz(Sum(PtsA.Hits),0) AS NofDeadPoints " & _
    "FROM (" & strPtsA & ") AS PtsA " & _
    "GROUP BY PtsA.Park, PtsA.IslandCode, PtsA.Location_ID, PtsA.SiteCode, PtsA.Vegetation_Community, PtsA.SurveyYear, PtsA.SurveyDate"

strHitsA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, 1 AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Condition)=" & Chr$(34) & "Dead" & Chr$(34) & " OR (tbl_Species_Data.Condition)=" & Chr$(34) & "Dead Plant" & Chr$(34) & " Or (tbl_Species_Data.Condition)=" & Chr$(34) & "Dead Branch" & Chr$(34) & "))"

strDeadHits = "SELECT HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate, Nz(Sum(HitsA.Hits),0) AS NofDeadHits " & _
    "FROM (" & strHitsA & ") AS HitsA " & _
    "GROUP BY HitsA.Park, HitsA.IslandCode, HitsA.Location_ID, HitsA.SiteCode, HitsA.Vegetation_Community, HitsA.SurveyYear, HitsA.SurveyDate"

strDeadCover = "SELECT ExpectedVisits.Park, ExpectedVisits.IslandCode, ExpectedVisits.Location_ID, ExpectedVisits.SiteCode, ExpectedVisits.Vegetation_Community, " & _
    "ExpectedVisits.SurveyYear, ExpectedVisits.SurveyDate, DeadObs.NofDeadObserved, TotalSppObserved.NofSppObserved, DeadPoints.NofDeadPoints, TotalPoints.NofPoints, " & _
    "DeadHits.NofDeadHits, TotalHits.NofHits, [NofDeadPoints]/[NofPoints] AS AbsoluteDeadCover, [NofDeadHits]/[NofHits] AS RelativeDeadCover " & _
    "FROM ((((((" & ExpectedVisitsSQL(xPark) & ") AS ExpectedVisits " & _
    "LEFT JOIN (" & strDeadObs & ") AS DeadObs ON (ExpectedVisits.Park = DeadObs.Park) AND (ExpectedVisits.IslandCode = DeadObs.IslandCode) AND (ExpectedVisits.SiteCode = DeadObs.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = DeadObs.Vegetation_Community) AND (ExpectedVisits.SurveyYear = DeadObs.SurveyYear) AND (ExpectedVisits.SurveyDate = DeadObs.SurveyDate)) " & _
    "LEFT JOIN (" & TotalSppObservedSQL(xPark) & ") AS TotalSppObserved ON (ExpectedVisits.Park = TotalSppObserved.Park) AND (ExpectedVisits.IslandCode = TotalSppObserved.IslandCode) AND (ExpectedVisits.SiteCode = TotalSppObserved.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalSppObserved.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalSppObserved.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalSppObserved.SurveyDate)) " & _
    "LEFT JOIN (" & strDeadPoints & ") AS DeadPoints ON (ExpectedVisits.Park = DeadPoints.Park) AND (ExpectedVisits.IslandCode = DeadPoints.IslandCode) AND (ExpectedVisits.SiteCode = DeadPoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = DeadPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = DeadPoints.SurveyYear) AND (ExpectedVisits.SurveyDate = DeadPoints.SurveyDate)) " & _
    "LEFT JOIN (" & TotalPointsSQL(xPark) & ") AS TotalPoints ON (ExpectedVisits.Park = TotalPoints.Park) AND (ExpectedVisits.IslandCode = TotalPoints.IslandCode) AND (ExpectedVisits.SiteCode = TotalPoints.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalPoints.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalPoints.SurveyDate)) " & _
    "LEFT JOIN (" & strDeadHits & ") AS DeadHits ON (ExpectedVisits.Park = DeadHits.Park) AND (ExpectedVisits.IslandCode = DeadHits.IslandCode) AND (ExpectedVisits.SiteCode = DeadHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = DeadHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = DeadHits.SurveyYear) AND (ExpectedVisits.SurveyDate = DeadHits.SurveyDate)) " & _
    "LEFT JOIN (" & TotalHitsSQL(xPark) & ") AS TotalHits ON (ExpectedVisits.Park = TotalHits.Park) AND (ExpectedVisits.IslandCode = TotalHits.IslandCode) AND (ExpectedVisits.SiteCode = TotalHits.SiteCode) " & _
    "AND (ExpectedVisits.Vegetation_Community = TotalHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalHits.SurveyYear) AND (ExpectedVisits.SurveyDate = TotalHits.SurveyDate)"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, " & _
    "CInt(qry.NofDeadObserved) AS NObserved, qry.NofSppObserved, qry.NofPoints, " & _
    "qry.AbsoluteDeadCover  AS AbsoluteCover, qry.RelativeDeadCover AS RelativeCover, " & Chr$(34) & "Dead_cover" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strDeadCover & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Dead_Cover = strSQL

End Function

Function Export_Substrate(xPark As Integer)
'Query #11: Number of recorded substrate "hits" by transect over time

Dim strA As String
Dim strTotalHits As String
Dim strSubstrateHits As String

strA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Iif(IsNull([tbl_Event_Point].[Substrate])," & Chr$(34) & "NA" & Chr$(34) & ",[tbl_Event_Point].[Substrate]) AS Ground_cover, " & _
    "1 AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Event_Point ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & ")"

strTotalHits = "SELECT A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate, A.Ground_cover, Sum(A.Hits) AS NofHits " & _
    "FROM (" & strA & ") AS A GROUP BY A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate, A.Ground_cover"

strSubstrateHits = "SELECT Total_hits.Park, Total_hits.IslandCode, Total_hits.Location_ID, Total_hits.SiteCode, Total_hits.Vegetation_Community, Total_hits.SurveyYear, Total_hits.SurveyDate, " & _
    "Total_hits.Ground_cover, Total_hits.NofHits, Total_points.NofPoints " & _
    "FROM (" & strTotalHits & ") AS Total_hits INNER JOIN (" & TotalPointsSQL(xPark) & ") AS Total_points ON (Total_hits.SurveyDate = Total_points.SurveyDate) " & _
    "AND (Total_hits.SurveyYear = Total_points.SurveyYear) AND (Total_hits.Vegetation_Community = Total_points.Vegetation_Community) AND (Total_hits.SiteCode = Total_points.SiteCode) " & _
    "AND (Total_hits.IslandCode = Total_points.IslandCode) AND (Total_hits.Park = Total_points.Park)"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Ground_cover, qry.NofHits AS TotalNoOfHits, qry.NofPoints AS TotalPoints, " & _
    Chr$(34) & "Substrate" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strSubstrateHits & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Substrate = strSQL

End Function

Function Export_Species_Richness(xPark As Integer)
'Query #12: Species richness within a vegetation community over time

Dim strA As String
Dim strSppRichness As String

strA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Vegetation_Community, Year([Start_Date]) AS SurveyYear, " & _
    "Park_Spp.Species_code, 1 AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
    "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & " Or (tbl_Species_Data.Condition) Is Null))"

strSppRichness = "SELECT A.Park, A.IslandCode, A.Vegetation_Community, A.SurveyYear, Sum(A.Hits) AS Frequency FROM (" & strA & ") AS A " & _
    "GROUP BY A.Park, A.IslandCode, A.Vegetation_Community, A.SurveyYear"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.Vegetation_Community, qry.SurveyYear, qry.Frequency, " & _
    Chr$(34) & "Species_richness" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strSppRichness & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Species_Richness = strSQL

End Function

Function Export_Shannons_Indices(xPark As Integer)
'Query # 13: Shannon Diversity and Evenness index calculations for a veg community type over time

Dim strA As String

Dim strB As String
Dim strTotalHits As String

Dim strC As String
Dim strTotalSppHits As String

Dim strP As String
Dim strPlnP As String

Dim strD As String
Dim strPlnPSum As String

Dim strSppRichness As String
Dim strShannon As String

strB = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Vegetation_Community, Year([Start_Date]) AS SurveyYear, tbl_Species_Data.Species_Code, 1 AS Hit " & _
    "FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Event_Point ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) INNER JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & ")"

strTotalHits = "SELECT B.Park, B.IslandCode, B.Vegetation_Community, B.SurveyYear, Sum(B.Hit) AS NTotalHits FROM (" & strB & ") AS B GROUP BY B.Park, B.IslandCode, B.Vegetation_Community, B.SurveyYear"

strC = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Vegetation_Community, Year([Start_Date]) AS SurveyYear, tbl_Species_Data.Species_Code, 1 AS Hit " & _
    "FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Event_Point ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) INNER JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & " Or (tbl_Species_Data.Condition) Is Null))"

strTotalSppHits = "SELECT C.Park, C.IslandCode, C.Vegetation_Community, C.SurveyYear, C.Species_Code, Sum(C.Hit) AS NSppHits " & _
    "FROM (" & strC & ") AS C GROUP BY C.Park, C.IslandCode, C.Vegetation_Community, C.SurveyYear, C.Species_Code"

strP = "SELECT qTotalSppHits.Park, qTotalSppHits.IslandCode, qTotalSppHits.Vegetation_Community, qTotalSppHits.SurveyYear, qTotalSppHits.Species_Code, qTotalSppHits.NSppHits, " & _
    "qTotalHits.NTotalHits, [NSppHits]/[NTotalHits] AS P FROM (" & strTotalHits & ") AS qTotalHits LEFT JOIN (" & strTotalSppHits & ") AS qTotalSppHits ON (qTotalHits.Park = qTotalSppHits.Park) " & _
    "AND (qTotalHits.IslandCode = qTotalSppHits.IslandCode) AND (qTotalHits.Vegetation_Community = qTotalSppHits.Vegetation_Community) AND (qTotalHits.SurveyYear = qTotalSppHits.SurveyYear) " & _
    "WHERE (((qTotalSppHits.Species_Code) Is Not Null))"

strPlnP = "SELECT qP.Park, qP.IslandCode, qP.Vegetation_Community, qP.SurveyYear, qP.Species_Code, ([P])*(Log([P])) AS PlnP FROM (" & strP & ") AS qP"

strD = "SELECT qPlnP.Park, qPlnP.IslandCode, qPlnP.Vegetation_Community, qPlnP.SurveyYear, Sum(qPlnP.PlnP) AS SumOfPlnP FROM (" & strPlnP & ") AS qPlnP " & _
    "GROUP BY qPlnP.Park, qPlnP.IslandCode, qPlnP.Vegetation_Community, qPlnP.SurveyYear"

strPlnPSum = "SELECT D.Park, D.IslandCode, D.Vegetation_Community, D.SurveyYear, -(D.SumOfPlnP) AS H FROM (" & strD & ") AS D"

strA = "SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Vegetation_Community, Year([Start_Date]) AS SurveyYear, tbl_Species_Data.Species_Code, 1 AS Hits " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & " Or (tbl_Species_Data.Condition) Is Null))"

strSppRichness = "SELECT A.Park, A.IslandCode, A.Vegetation_Community, A.SurveyYear, Sum(A.Hits) AS NSpp FROM (" & strA & ") AS A " & _
    "GROUP BY A.Park, A.IslandCode, A.Vegetation_Community, A.SurveyYear"

strShannon = "SELECT qPlnPSum.Park, qPlnPSum.IslandCode, qPlnPSum.Vegetation_Community, qPlnPSum.SurveyYear, qPlnPSum.H AS Shannons_Diversity_index, [H]/(Log([NSpp])) AS Shannons_Evenness_index " & _
    "FROM (" & strPlnPSum & ") AS qPlnPSum INNER JOIN (" & strSppRichness & ") AS qSppRichness ON (qPlnPSum.SurveyYear = qSppRichness.SurveyYear) " & _
    "AND (qPlnPSum.Vegetation_Community = qSppRichness.Vegetation_Community) AND (qPlnPSum.IslandCode = qSppRichness.IslandCode) AND (qPlnPSum.Park = qSppRichness.Park)"

strSQL = "SELECT " & ParkSelect(xPark) & ", qry.Vegetation_Community, qry.SurveyYear, qry.Shannons_Diversity_index, qry.Shannons_Evenness_index, " & _
    Chr$(34) & "Shannons_indices" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strShannon & ") AS qry " & _
    "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

Export_Shannons_Indices = strSQL

End Function

Function Export_Custom(xPark As Integer)
'Query #10: Custom query based on user parameters

    Dim frm As Form

    Dim strSpeciesFilter As String
    Dim strNativeFilter As String
    Dim strCycleFilter As String
    Dim strStatusFilter As String
    Dim strGrowthFormFilter As String
    Dim strCustomFilter As String

    Dim strParam As String

    Dim strA As String
    Dim strObsA As String
    Dim strObs As String
    Dim strCover As String

    Dim strSQL As String
    ' ------------------------------------------------------------------------

    Set frm = Forms!frm_ExportData

    'Set SQL filter parameters
    ' Species--------------------------------------------------------------
    Select Case frm!grpSpeciesFilter
    Case Is = 1 ' all species
        strSpeciesFilter = ""
    Case Is = 2 ' by species
        strSpeciesFilter = " AND ((Species_Code)=" & Chr$(34) & frm!cmbSpecies & Chr$(34) & ")"
    Case Else
        strSpeciesFilter = ""
    End Select
    ' ------------------------------------------------------------------------
    ' Nativity  -----------------------------------------------------------
    Select Case frm!grpNativeFilter
    Case Is = 1 ' native
        strNativeFilter = " AND ((Nativity)=" & Chr$(34) & "Native" & Chr$(34) & ")"
    Case Is = 2 ' nonnative
        strNativeFilter = " AND ((Nativity)=" & Chr$(34) & "Nonnative" & Chr$(34) & ")"
    Case Else
        strNativeFilter = ""
    End Select
    ' ------------------------------------------------------------------------
    ' Life Cycle Form ---------------------------------------------------------
    Select Case frm!grpCycleFilter
    Case Is = 1 ' native
        strCycleFilter = " AND ((AnnPer)=" & Chr$(34) & "Annual" & Chr$(34) & ")"
    Case Is = 2 ' nonnative
        strCycleFilter = " AND ((AnnPer)=" & Chr$(34) & "Perennial" & Chr$(34) & ")"
    Case Else
        strCycleFilter = ""
    End Select
    ' ------------------------------------------------------------------------
    ' Status Form -------------------------------------------------------------
    Select Case frm!grpDataFilter
        Case Is = 1
            Select Case frm!grpStatusFilter
            Case Is = 1 ' live
                'strStatusFilter = " AND ((tbl_Species_Data.Condition) Is Null Or (tbl_Species_Data.Condition)=" & Chr$(34) & "Alive" & Chr$(34) & ")"
                strStatusFilter = " AND ((Condition)=" & Chr$(34) & "Alive" & Chr$(34) & ")"
            Case Is = 2 'dead
                'strStatusFilter = " AND ((tbl_Species_Data.Condition)=" & Chr$(34) & "Dead" & Chr$(34) & " Or (tbl_Species_Data.Condition)=" & Chr$(34) & "Dead Plant" & Chr$(34) & " " & _
                    "Or (tbl_Species_Data.Condition)=" & Chr$(34) & "Dead Branch" & Chr$(34) & ")"
                strStatusFilter = " AND ((Condition)=" & Chr$(34) & "Dead" & Chr$(34) & ")"
            Case Else
                strStatusFilter = ""
            End Select
        Case 2, 3, 4, 5
            Select Case frm!grpStatusFilter
            Case Is = 1 ' live
                strStatusFilter = " AND ((tlu_Condition.Analysis_code) = " & Chr$(34) & "Alive" & Chr$(34) & ")"
            Case Is = 2 'dead
                strStatusFilter = " AND ((tlu_Condition.Analysis_code) = " & Chr$(34) & "Dead" & Chr$(34) & ")"
            Case Else
                strStatusFilter = ""
            End Select
    End Select
    ' ------------------------------------------------------------------------
    ' Growth Form -------------------------------------------------------------
    If frm!optGrass = -1 Then
        strGrowthFormFilter = "(FxnGroup) = " & Chr$(34) & "Grass" & Chr$(34)
    Else
        strGrowthFormFilter = ""
    End If

    If frm!optHerbaceous = -1 Then
        If strGrowthFormFilter = "" Then
            strGrowthFormFilter = "(FxnGroup) = " & Chr$(34) & "Herbaceous" & Chr$(34)
        Else
            strGrowthFormFilter = strGrowthFormFilter & " OR (FxnGroup) = " & Chr$(34) & "Herbaceous" & Chr$(34)
        End If
    Else
        strGrowthFormFilter = strGrowthFormFilter
    End If

    If frm!optShrub = -1 Then
        If strGrowthFormFilter = "" Then
            strGrowthFormFilter = "(FxnGroup) = " & Chr$(34) & "Shrub" & Chr$(34)
        Else
            strGrowthFormFilter = strGrowthFormFilter & " OR (FxnGroup) = " & Chr$(34) & "S" & Chr$(34)
        End If
    Else
        strGrowthFormFilter = strGrowthFormFilter
    End If

    If frm!optSubshrub = -1 Then
        If strGrowthFormFilter = "" Then
            strGrowthFormFilter = "(FxnGroup) = " & Chr$(34) & "Sub-shrub" & Chr$(34)
        Else
            strGrowthFormFilter = strGrowthFormFilter & " OR (FxnGroup) = " & Chr$(34) & "Sub-shrub" & Chr$(34)
        End If
    Else
        strGrowthFormFilter = strGrowthFormFilter
    End If

    If frm!optTree = -1 Then
        If strGrowthFormFilter = "" Then
            strGrowthFormFilter = "(FxnGroup) = " & Chr$(34) & "Tree" & Chr$(34)
        Else
            strGrowthFormFilter = strGrowthFormFilter & " OR (FxnGroup) = " & Chr$(34) & "Tree" & Chr$(34)
        End If
    Else
        strGrowthFormFilter = strGrowthFormFilter
    End If

    If strGrowthFormFilter = "" Then
        strGrowthFormFilter = ""
    Else
        strGrowthFormFilter = " AND ((" & strGrowthFormFilter & "))"
    End If

    strCustomFilter = strSpeciesFilter & strNativeFilter & strCycleFilter & strStatusFilter & strGrowthFormFilter

    ' ------------------------------------------------------------------------
    ' string value for [Query_parameters] field
    strParam = Chr$(34) & "Park: " & frm!txtParkFilter & "; Island: " & frm!txtIslandFilter & "; Transect: " & frm!txtSiteFilter & "; " & _
        "Vegetation_Community: " & frm!txtVegetationFilter & "; Year: " & frm!txtYearFilter & "; Species: " & frm!txtSpeciesFilter & Chr$(34)

    ' ------------------------------------------------------------------------
    ' SQL strings for Data Output options

    Select Case frm!grpDataFilter
    Case Is = 1 ' Raw Data----------------------------------------------------
        strA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Location_ID, tbl_Locations.Vegetation_Community, " & _
            "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No AS PointNo, tbl_Event_Point.Substrate, tbl_Event_Point.MaxHeight_cm AS MaxHt, " & _
            "tbl_Event_Point.MaxHeight_cm_DEAD AS MaxHt_dead, tbl_Species_Data.Species_Rank, tbl_Species_Data.Species_Code, Park_Spp.Scientific_name, Park_Spp.FxnGroup, Park_Spp.Nativity, " & _
            "Park_Spp.AnnPer, tlu_Condition.Analysis_code AS Condition " & _
            "FROM ((tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point LEFT JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
            "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN tlu_Condition " & _
            "ON tbl_Species_Data.Condition = tlu_Condition.Condition) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON Park_Spp.Species_code = tbl_Species_Data.Species_Code " & _
            "WHERE (" & LocTypeFilter(xPark) & ")"

        strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.PointNo, qry.Substrate, qry.MaxHt, qry.MaxHt_dead, " & _
            "qry.Species_Code, qry.Scientific_name, qry.FxnGroup, qry.Nativity, qry.AnnPer, qry.Condition, " & Chr$(34) & "Custom, Raw Data" & Chr$(34) & " AS Query_type, " & _
            strParam & " AS Query_parameters " & _
            "FROM (" & strA & ") AS qry " & _
            "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & strCustomFilter & ") " & _
            "ORDER BY " & ParkSelect(xPark) & ", qry.SiteCode, qry.SurveyYear, qry.SurveyDate, qry.PointNo, qry.Species_Rank"

        Export_Custom = strSQL

    Case Is = 2 ' Absolute and Relative cover---------------------------------
        strObsA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
            "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, Park_Spp.Species_code, 1 AS Hits " & _
            "FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
            "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
            "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN tlu_Condition ON tbl_Species_Data.Condition = tlu_Condition.Condition " & _
            "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Species_Code) Is Not Null)" & strCustomFilter & ")"

        strObs = "SELECT ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate" & _
            FieldSelect("Species", "ObsA") & ", Sum(ObsA.Hits) AS NofObserved " & _
            "FROM (" & strObsA & ") AS ObsA " & _
            "GROUP BY ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate" & IIf(frm!optSpeciesGroup = 0, ", ObsA.Species_Code", "")

        strCover = "SELECT ExpectedVisits.Park, ExpectedVisits.IslandCode, ExpectedVisits.Location_ID, ExpectedVisits.SiteCode, ExpectedVisits.Vegetation_Community, " & _
            "ExpectedVisits.SurveyYear, ExpectedVisits.SurveyDate" & IIf(frm!optSpeciesGroup = 0, ", Obs.Species_Code", "") & ", Obs.NofObserved, TotalHits.NofHits, TotalPoints.NofPoints, " & _
            "[NofObserved]/[NofPoints] AS AbsoluteCover, [NofObserved]/[NofHits] AS RelativeCover " & _
            "FROM (((" & ExpectedVisitsSQL(xPark) & ") AS ExpectedVisits LEFT JOIN (" & strObs & ") AS Obs ON (ExpectedVisits.Park = Obs.Park) AND (ExpectedVisits.IslandCode = Obs.IslandCode) " & _
            "AND (ExpectedVisits.SiteCode = Obs.SiteCode) AND (ExpectedVisits.Vegetation_Community = Obs.Vegetation_Community) AND (ExpectedVisits.SurveyYear = Obs.SurveyYear)) " & _
            "LEFT JOIN (" & TotalPointsSQL(xPark) & ") AS TotalPoints ON (ExpectedVisits.Park = TotalPoints.Park) AND (ExpectedVisits.IslandCode = TotalPoints.IslandCode) " & _
            "AND (ExpectedVisits.SiteCode = TotalPoints.SiteCode) AND (ExpectedVisits.Vegetation_Community = TotalPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalPoints.SurveyYear)) " & _
            "LEFT JOIN (" & TotalHitsSQL(xPark) & ") AS TotalHits ON (ExpectedVisits.Park = TotalHits.Park) AND (ExpectedVisits.IslandCode = TotalHits.IslandCode) " & _
            "AND (ExpectedVisits.SiteCode = TotalHits.SiteCode) AND (ExpectedVisits.Vegetation_Community = TotalHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalHits.SurveyYear)"

        strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate" & _
            FieldSelect("Species", "qry") & ", qry.NofObserved AS NObserved, qry.NofHits AS TotalNoOfHits, qry.NofPoints AS TotalPoints, qry.AbsoluteCover, qry.RelativeCover, " & _
            Chr$(34) & "Custom, Absolute/Relative Cover" & Chr$(34) & " AS Query_type, " & strParam & " AS Query_parameters " & _
            "FROM (" & strCover & ") AS qry " & _
            "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

        Export_Custom = strSQL

    Case Is = 3 ' Total Hits--------------------------------------------------
        Dim strTotalHits As String

        strObsA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
            "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, Park_Spp.Species_Code, Park_Spp.Nativity, Park_Spp.AnnPer, " & _
            "tlu_Condition.Analysis_code AS Condition, Park_Spp.FxnGroup, 1 AS Hits " & _
            "FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
            "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
            "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN tlu_Condition ON tbl_Species_Data.Condition = tlu_Condition.Condition " & _
            "WHERE (" & LocTypeFilter(xPark) & " AND ((tbl_Species_Data.Species_Code) Is Not Null)" & strCustomFilter & ")"

        strObs = "SELECT ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate, ObsA.Species_Code, ObsA.Nativity, ObsA.AnnPer, " & _
            "ObsA.Condition, ObsA.FxnGroup, Sum(ObsA.Hits) AS NofObserved " & _
            "FROM (" & strObsA & ") AS ObsA " & _
            "GROUP BY ObsA.Park, ObsA.IslandCode, ObsA.Location_ID, ObsA.SiteCode, ObsA.Vegetation_Community, ObsA.SurveyYear, ObsA.SurveyDate, ObsA.Species_Code, ObsA.Nativity, ObsA.AnnPer, " & _
            "ObsA.Condition, ObsA.FxnGroup"

        strTotalHits = "SELECT ExpectedVisits.Park, ExpectedVisits.IslandCode, ExpectedVisits.Location_ID, ExpectedVisits.SiteCode, ExpectedVisits.Vegetation_Community, " & _
            "ExpectedVisits.SurveyYear, ExpectedVisits.SurveyDate, Obs.Species_Code, Obs.Nativity, Obs.AnnPer, Obs.Condition, Obs.FxnGroup, " & _
            "Obs.NofObserved, TotalHits.NofHits, TotalPoints.NofPoints " & _
            "FROM (((" & ExpectedVisitsSQL(xPark) & ") AS ExpectedVisits LEFT JOIN (" & strObs & ") AS Obs ON (ExpectedVisits.Park = Obs.Park) AND (ExpectedVisits.IslandCode = Obs.IslandCode) " & _
            "AND (ExpectedVisits.SiteCode = Obs.SiteCode) AND (ExpectedVisits.Vegetation_Community = Obs.Vegetation_Community) AND (ExpectedVisits.SurveyYear = Obs.SurveyYear)) " & _
            "LEFT JOIN (" & TotalPointsSQL(xPark) & ") AS TotalPoints ON (ExpectedVisits.Park = TotalPoints.Park) AND (ExpectedVisits.IslandCode = TotalPoints.IslandCode) " & _
            "AND (ExpectedVisits.SiteCode = TotalPoints.SiteCode) AND (ExpectedVisits.Vegetation_Community = TotalPoints.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalPoints.SurveyYear)) " & _
            "LEFT JOIN (" & TotalHitsSQL(xPark) & ") AS TotalHits ON (ExpectedVisits.Park = TotalHits.Park) AND (ExpectedVisits.IslandCode = TotalHits.IslandCode) " & _
            "AND (ExpectedVisits.SiteCode = TotalHits.SiteCode) AND (ExpectedVisits.Vegetation_Community = TotalHits.Vegetation_Community) AND (ExpectedVisits.SurveyYear = TotalHits.SurveyYear)"

        strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate" & _
            FieldSelect("Species", "qry") & FieldSelect("Nativity", "qry") & FieldSelect("Cycle", "qry") & FieldSelect("Status", "qry") & FieldSelect("Growth Form", "qry") & ", " & _
            "Sum(qry.NofObserved) AS NObserved, qry.NofHits AS TotalNoOfHits, qry.NofPoints AS TotalPoints, " & _
            Chr$(34) & "Custom, Total Hits" & Chr$(34) & " AS Query_type, " & strParam & " AS Query_parameters " & _
            "FROM (" & strTotalHits & ") AS qry " & _
            "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ") " & _
            "GROUP BY " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate" & _
            FieldSelect("Species", "qry") & FieldSelect("Nativity", "qry") & FieldSelect("Cycle", "qry") & FieldSelect("Status", "qry") & FieldSelect("Growth Form", "qry") & ", " & _
            "qry.NofHits, qry.NofPoints, " & Chr$(34) & "Custom, Total Hits" & Chr$(34) & ", " & strParam

        Export_Custom = strSQL

    Case Is = 4 ' Average Heights---------------------------------------------
        Dim strMaxRank As String
        Dim strAvgHeights As String

        'species are read from bottom to top hence the max species_rank is the max_height
        strMaxRank = "SELECT tbl_Event_Point.Event_Point_ID, Max(tbl_Species_Data.Species_Rank) AS MaxOfSpecies_Rank FROM tbl_Event_Point INNER JOIN tbl_Species_Data " & _
            "ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID GROUP BY tbl_Event_Point.Event_Point_ID"

        strA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
            "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, tbl_Event_Point.Point_No, tbl_Event_Point.MaxHeight_cm, Park_Spp.Species_code, Park_Spp.Scientific_Name, " & _
            "Park_Spp.Nativity, Park_Spp.AnnPer, tlu_Condition.Analysis_code AS Condition, Park_Spp.FxnGroup " & _
            "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (((tbl_Species_Data INNER JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
            "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) INNER JOIN (" & strMaxRank & ") AS MaxRank ON (tbl_Species_Data.Event_Point_ID = MaxRank.Event_Point_ID) " & _
            "AND (tbl_Species_Data.Species_Rank = MaxRank.MaxOfSpecies_Rank)) LEFT JOIN tlu_Condition ON tbl_Species_Data.Condition = tlu_Condition.Condition) " & _
            "ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
            "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
            "WHERE (" & LocTypeFilter(xPark) & strCustomFilter & ") " & _
            "GROUP BY tbl_Sites.Unit_Code, tbl_Sites.Site_Name, tbl_Locations.Location_ID, tbl_Locations.Location_Code, tbl_Locations.Vegetation_Community, Year([Start_Date]), tbl_Events.Start_Date, " & _
            "tbl_Event_Point.Point_No, tbl_Event_Point.MaxHeight_cm, Park_Spp.Species_code, Park_Spp.Scientific_Name, Park_Spp.FxnGroup, Park_Spp.Nativity, Park_Spp.AnnPer, tlu_Condition.Analysis_code  " & _
            "HAVING (((Max(tbl_Species_Data.Species_Rank))<>False) AND ((tbl_Event_Point.MaxHeight_cm) Is Not Null))"

        strAvgHeights = "SELECT A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate" & _
            FieldSelect("Species", "A") & FieldSelect("Nativity", "A") & FieldSelect("Cycle", "A") & FieldSelect("Status", "A") & FieldSelect("Growth Form", "A") & ", " & _
            "Avg(A.MaxHeight_cm) AS MeanHeight, StDev(A.MaxHeight_cm) AS StdDev, " & _
            "Count(A.MaxHeight_cm) AS N, StDev([MaxHeight_cm])/Count([MaxHeight_cm]) AS StdErr, Min(A.MaxHeight_cm) AS MinRange, Max(A.MaxHeight_cm) AS MaxRange " & _
            "FROM (" & strA & ") AS A " & _
            "GROUP BY A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate" & _
            FieldSelect("Species", "A") & FieldSelect("Nativity", "A") & FieldSelect("Cycle", "A") & FieldSelect("Status", "A") & FieldSelect("Growth Form", "A")

        strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate" & _
            FieldSelect("Species", "qry") & FieldSelect("Nativity", "qry") & FieldSelect("Cycle", "qry") & FieldSelect("Status", "qry") & FieldSelect("Growth Form", "qry") & ", " & _
            "qry.MeanHeight, qry.StdDev, qry.N, qry.StdErr, qry.MinRange, qry.MaxRange, " & _
            Chr$(34) & "Avg_height" & Chr$(34) & " AS Query_type, " & strParam & " AS Query_parameters " & _
            "FROM (" & strAvgHeights & ") AS qry " & _
            "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

        Export_Custom = strSQL

    Case Is = 5 ' Species Frequency-------------------------------------------
        Dim strSppFrequency As String

        strA = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
            "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Park_Spp.Species_code, Park_Spp.Scientific_name, Park_Spp.Nativity, Park_Spp.AnnPer, " & _
            "tlu_Condition.Analysis_code AS Condition, Park_Spp.FxnGroup, 1 AS Hits " & _
            "FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp " & _
            "ON tbl_Species_Data.Species_Code = Park_Spp.Species_code) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
            "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN tlu_Condition ON tbl_Species_Data.Condition = tlu_Condition.Condition " & _
            "WHERE (" & LocTypeFilter(xPark) & strCustomFilter & ")"

        strSppFrequency = "SELECT A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate, A.Species_code, A.Scientific_name" & _
            FieldSelect("Species", "A") & FieldSelect("Nativity", "A") & FieldSelect("Cycle", "A") & FieldSelect("Status", "A") & FieldSelect("Growth Form", "A") & ", Sum(A.Hits) AS Frequency " & _
            "FROM (" & strA & ") AS A " & _
            "GROUP BY A.Park, A.IslandCode, A.Location_ID, A.SiteCode, A.Vegetation_Community, A.SurveyYear, A.SurveyDate, A.Species_code, A.Scientific_name" & _
            FieldSelect("Species", "A") & FieldSelect("Nativity", "A") & FieldSelect("Cycle", "A") & FieldSelect("Status", "A") & FieldSelect("Growth Form", "A")

        strSQL = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Species_code, qry.Scientific_name, qry.Frequency" & _
            FieldSelect("Species", "qry") & FieldSelect("Nativity", "qry") & FieldSelect("Cycle", "qry") & FieldSelect("Status", "qry") & FieldSelect("Growth Form", "qry") & ", " & _
            Chr$(34) & "Species_frequency" & Chr$(34) & " AS Query_type, " & strParam & " AS Query_parameters " & _
            "FROM (" & strSppFrequency & ") AS qry " & _
            "WHERE (" & FilterString(Forms!frm_ExportData, xPark) & ")"

        Export_Custom = strSQL

    Case Else
        MsgBox "ERROR!"

    End Select

End Function

Function FieldSelect(strSelect As String, strField As String)
'SQL select statement for species parameters

    Dim frm As Form

    Set frm = Forms!frm_ExportData

    Select Case strSelect
    Case Is = "Species"
        FieldSelect = IIf(frm!optSpeciesGroup = 0, ", " & strField & "." & "Species_Code", "")
    Case Is = "Nativity"
        FieldSelect = IIf(frm!optNativityGroup = 0, ", " & strField & "." & "Nativity", "")
    Case Is = "Cycle"
        FieldSelect = IIf(frm!optCycleGroup = 0, ", " & strField & "." & "AnnPer", "")
    Case Is = "Status"
        FieldSelect = IIf(frm!optStatusGroup = 0, ", " & strField & "." & "Condition", "")
    Case Is = "Growth Form"
        FieldSelect = IIf(frm!optGrowthFormGroup = 0, ", " & strField & "." & "FxnGroup", "")
    End Select

End Function

Function Export_AnnualReport_AbsoluteCover(xPark As Integer, xYear As Integer)

Dim strWhere As String

Dim strRaw As String
Dim strRawSum As String

Dim str1 As String
Dim str2 As String
Dim str0Data As String

Dim strData As String

Dim strAbsCovData As String
Dim strAbsCov As String 'final SQL string

' Create WHERE string --------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ") AND ((tlu_Condition.Analysis_code) Is Null Or (tlu_Condition.Analysis_code)=" & Chr$(34) & "Alive" & Chr$(34) & ")"
' ----------------------------------------------------------------------------

' Create data strings --------------------------------------------------------
strRaw = "SELECT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "tbl_Species_Data.Species_Code, tlu_Condition.Analysis_code AS Condition, Park_Spp.FxnGroup, Park_Spp.Nativity " & _
    "FROM ((tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point LEFT JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN tlu_Condition " & _
    "ON tbl_Species_Data.Condition = tlu_Condition.Condition) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON Park_Spp.Species_code = tbl_Species_Data.Species_Code " & _
    "WHERE (" & strWhere & ")"

strRawSum = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qRaw.Nativity, Count(qRaw.Species_Code) AS N " & _
    "FROM (" & strRaw & ") AS qRaw GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qRaw.Nativity"
'-----
str1 = "SELECT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & "))"

str2 = "SELECT DISTINCT Park_Spp.FxnGroup, Park_Spp.Nativity " & _
    "FROM ((tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point LEFT JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID ) LEFT JOIN tlu_Condition " & _
    "ON tbl_Species_Data.Condition = tlu_Condition.Condition) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON Park_Spp.Species_code = tbl_Species_Data.Species_Code " & _
    "WHERE (" & strWhere & ")"

str0Data = "SELECT qry1.SurveyYear, qry1.Park, qry1.IslandCode, qry1.SiteCode, qry1.Vegetation_Community, qry2.FxnGroup, qry2.Nativity, 0 AS N " & _
    "FROM (" & str1 & ")  AS qry1, (" & str2 & ")  AS qry2"
'-----
'strData = strRawSum + str0Data
strData = "SELECT qryUnion.SurveyYear, qryUnion.Park, qryUnion.IslandCode, qryUnion.SiteCode, qryUnion.Vegetation_Community, qryUnion.FxnGroup, qryUnion.Nativity, Sum(qryUnion.N) AS SumOfN " & _
    "FROM (SELECT * FROM (" & str0Data & ") AS q0Data UNION SELECT * FROM (" & strRawSum & ") AS qryRawSum)  AS qryUnion " & _
    "GROUP BY qryUnion.SurveyYear, qryUnion.Park, qryUnion.IslandCode, qryUnion.SiteCode, qryUnion.Vegetation_Community, qryUnion.FxnGroup, qryUnion.Nativity"
'-----------------------------------------------------------------------------

' Calculating Absolute Cover (Figure E2) -------------------------------------
strAbsCovData = "SELECT qData.SurveyYear, qData.Park, qData.IslandCode, qData.SiteCode, qData.Vegetation_Community, qData.FxnGroup, qData.Nativity, qData.SumOfN, qTotalPoints.NofPoints, " & _
    "([SumOfN]/[NofPoints])*100 AS AbsCover " & _
    "FROM (" & strData & ") AS qData INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qData.SurveyYear = qTotalPoints.SurveyYear) AND (qData.Park = qTotalPoints.Park) " & _
    "AND (qData.IslandCode = qTotalPoints.IslandCode) AND (qData.SiteCode = qTotalPoints.SiteCode)"

strAbsCov = "SELECT qry.SurveyYear, " & ParkSelect(xPark) & ", qry.Vegetation_Community, Count(qry.SiteCode) AS NofTransects, qry.FxnGroup, qry.Nativity, " & _
    "Avg(qry.AbsCover) AS Average, StDev(qry.AbsCover) AS StdDev, Min(qry.AbsCover) AS MinRange, Max(qry.AbsCover) AS MaxRange, " & _
    Chr$(34) & "Annual Report, Absolute Cover (Fig. E2)" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strAbsCovData & ") AS qry " & _
    "GROUP BY qry.SurveyYear, " & ParkSelect(xPark) & ", qry.Vegetation_Community, qry.FxnGroup, qry.Nativity"

Export_AnnualReport_AbsoluteCover = strAbsCov

End Function

Function Export_AnnualReport_RelativeCover(xPark As Integer, xYear As Integer)

Dim strWhere As String

Dim strRaw As String
Dim strRawSum As String

Dim str1 As String
Dim str2 As String
Dim str0Data As String

Dim strData As String

Dim strRelCovData As String
Dim strRelCov As String 'final SQL string

' Create WHERE string --------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ") AND ((tlu_Condition.Analysis_code) Is Null Or (tlu_Condition.Analysis_code)=" & Chr$(34) & "Alive" & Chr$(34) & ")"
' ----------------------------------------------------------------------------

' Create data strings --------------------------------------------------------
strRaw = "SELECT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "tbl_Species_Data.Species_Code, tlu_Condition.Analysis_code AS Condition, Park_Spp.FxnGroup, Park_Spp.Nativity " & _
    "FROM ((tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point LEFT JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN tlu_Condition " & _
    "ON tbl_Species_Data.Condition = tlu_Condition.Condition) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON Park_Spp.Species_code = tbl_Species_Data.Species_Code " & _
    "WHERE (" & strWhere & ")"

strRawSum = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qRaw.Nativity, Count(qRaw.Species_Code) AS N " & _
    "FROM (" & strRaw & ") AS qRaw GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qRaw.Nativity"
'-----
str1 = "SELECT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & "))"

str2 = "SELECT DISTINCT Park_Spp.FxnGroup, Park_Spp.Nativity " & _
    "FROM ((tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point LEFT JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID ) LEFT JOIN tlu_Condition " & _
    "ON tbl_Species_Data.Condition = tlu_Condition.Condition) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON Park_Spp.Species_code = tbl_Species_Data.Species_Code " & _
    "WHERE (" & strWhere & ")"

str0Data = "SELECT qry1.SurveyYear, qry1.Park, qry1.IslandCode, qry1.SiteCode, qry1.Vegetation_Community, qry2.FxnGroup, qry2.Nativity, 0 AS N " & _
    "FROM (" & str1 & ")  AS qry1, (" & str2 & ")  AS qry2"
'-----
' strData = strRawSum + str0Data
strData = "SELECT qryUnion.SurveyYear, qryUnion.Park, qryUnion.IslandCode, qryUnion.SiteCode, qryUnion.Vegetation_Community, qryUnion.FxnGroup, qryUnion.Nativity, Sum(qryUnion.N) AS SumOfN " & _
    "FROM (SELECT * FROM (" & str0Data & ") AS q0Data UNION SELECT * FROM (" & strRawSum & ") AS qryRawSum)  AS qryUnion " & _
    "GROUP BY qryUnion.SurveyYear, qryUnion.Park, qryUnion.IslandCode, qryUnion.SiteCode, qryUnion.Vegetation_Community, qryUnion.FxnGroup, qryUnion.Nativity"
'-----------------------------------------------------------------------------

' Calculating Relative Cover (Figure E3 and E4)-------------------------------
strRelCovData = "SELECT qData.SurveyYear, qData.Park, qData.IslandCode, qData.SiteCode, qData.Vegetation_Community, qData.FxnGroup, qData.Nativity, qData.SumOfN, qTotalLiveHits.NofHits, " & _
    "([SumOfN]/[NofHits])*100 AS RelCover " & _
    "FROM (" & strData & ") AS qData INNER JOIN (" & TotalLiveHitsSQL(xPark) & ") AS qTotalLiveHits ON (qData.SurveyYear = qTotalLiveHits.SurveyYear) AND (qData.Park = qTotalLiveHits.Park) " & _
    "AND (qData.IslandCode = qTotalLiveHits.IslandCode) AND (qData.SiteCode = qTotalLiveHits.SiteCode)"

strRelCov = "SELECT qry.SurveyYear, " & ParkSelect(xPark) & ", qry.Vegetation_Community, Count(qry.SiteCode) AS NofTransects, qry.FxnGroup, qry.Nativity, " & _
    "Avg(qry.RelCover) AS Average, StDev(qry.RelCover) AS StdDev, Min(qry.RelCover) AS MinRange, Max(qry.RelCover) AS MaxRange, " & _
    Chr$(34) & "Annual Report, Relative Cover (Fig. E3 and Fig. E4)" & Chr$(34) & " AS Query_type " & _
    "FROM (" & strRelCovData & ") AS qry " & _
    "GROUP BY qry.SurveyYear, " & ParkSelect(xPark) & ", qry.Vegetation_Community, qry.FxnGroup, qry.Nativity"

Export_AnnualReport_RelativeCover = strRelCov

End Function

Function Export_AnnualReport_AbsoluteCover_byGroup(xPark As Integer, xYear As Integer)
'Table E7

Dim strWhere As String

Dim strRaw As String

Dim strAbs1raw As String
Dim strAbs2raw As String
Dim strAbs3raw As String
Dim strAbs4raw As String
Dim strAbs5raw As String
Dim strAbs6raw As String
Dim strAbs7raw As String
Dim strAbs8raw As String

Dim strAbs1 As String
Dim strAbs2 As String
Dim strAbs3 As String
Dim strAbs4 As String
Dim strAbs5 As String
Dim strAbs6 As String
Dim strAbs7 As String
Dim strAbs8 As String

Dim strAbsUnion As String 'final SQL string

' Create WHERE string --------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ") AND ((tlu_Condition.Analysis_code) Is Null Or (tlu_Condition.Analysis_code)=" & Chr$(34) & "Alive" & Chr$(34) & ")"
' ----------------------------------------------------------------------------

' Create data strings --------------------------------------------------------
strRaw = "SELECT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "tbl_Species_Data.Species_Code, tlu_Condition.Analysis_code AS Condition, Park_Spp.FxnGroup, Park_Spp.Nativity " & _
    "FROM ((tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point LEFT JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN tlu_Condition " & _
    "ON tbl_Species_Data.Condition = tlu_Condition.Condition) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON Park_Spp.Species_code = tbl_Species_Data.Species_Code " & _
    "WHERE (" & strWhere & ")"
' ----------------------------------------------------------------------------

'1 Absolute cover ALL --------------------------------------------------------
strAbs1raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, (Count([Species_Code])/[NofPoints])*100 AS AbsCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRaw.SiteCode = qTotalPoints.SiteCode) " & _
    "AND (qRaw.Park = qTotalPoints.Park) AND (qRaw.IslandCode = qTotalPoints.IslandCode) AND (qRaw.SurveyYear = qTotalPoints.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qTotalPoints.NofPoints"

strAbs1 = "SELECT qAbs1.SurveyYear, qAbs1.Park, qAbs1.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, " & _
    Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(qAbs1.AbsCover) AS Average, StDev(qAbs1.AbsCover) AS StdDev, Min(qAbs1.AbsCover) AS MinRange, Max(qAbs1.AbsCover) AS MaxRange " & _
    "FROM (" & strAbs1raw & ") AS qAbs1 " & _
    "GROUP BY qAbs1.SurveyYear, qAbs1.Park, qAbs1.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34) & ""
'2 Absolute cover ALL NAT ----------------------------------------------------
strAbs2raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Nativity, (Count([Species_Code])/[NofPoints])*100 AS AbsCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRaw.SiteCode = qTotalPoints.SiteCode) " & _
    "AND (qRaw.Park = qTotalPoints.Park) AND (qRaw.IslandCode = qTotalPoints.IslandCode) AND (qRaw.SurveyYear = qTotalPoints.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Nativity, qTotalPoints.NofPoints"

strAbs2 = "SELECT qAbs2.SurveyYear, qAbs2.Park, qAbs2.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, qAbs2.Nativity, " & _
    "Avg(qAbs2.AbsCover) AS Average, StDev(qAbs2.AbsCover) AS StdDev, Min(qAbs2.AbsCover) AS MinRange, Max(qAbs2.AbsCover) AS MaxRange " & _
    "FROM (" & strAbs2raw & ") AS qAbs2 " & _
    "GROUP BY qAbs2.SurveyYear, qAbs2.Park, qAbs2.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34) & ", qAbs2.Nativity"
'3 Absolute cover ALL VEG ----------------------------------------------------
strAbs3raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, (Count([Species_Code])/[NofPoints])*100 AS AbsCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRaw.SiteCode = qTotalPoints.SiteCode) " & _
    "AND (qRaw.Park = qTotalPoints.Park) AND (qRaw.IslandCode = qTotalPoints.IslandCode) AND (qRaw.SurveyYear = qTotalPoints.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qTotalPoints.NofPoints"

strAbs3 = "SELECT qAbs3.SurveyYear, qAbs3.Park, qAbs3.IslandCode, qAbs3.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, " & Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(qAbs3.AbsCover) AS Average, StDev(qAbs3.AbsCover) AS StdDev, Min(qAbs3.AbsCover) AS MinRange, Max(qAbs3.AbsCover) AS MaxRange " & _
    "FROM (" & strAbs3raw & ") AS qAbs3 " & _
    "GROUP BY qAbs3.SurveyYear, qAbs3.Park, qAbs3.IslandCode, qAbs3.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34) & ""
'4 Absolute cover ALL NAT VEG ------------------------------------------------
strAbs4raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.Nativity, (Count([Species_Code])/[NofPoints])*100 AS AbsCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRaw.SiteCode = qTotalPoints.SiteCode) " & _
    "AND (qRaw.Park = qTotalPoints.Park) AND (qRaw.IslandCode = qTotalPoints.IslandCode) AND (qRaw.SurveyYear = qTotalPoints.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.Nativity, qTotalPoints.NofPoints"

strAbs4 = "SELECT qAbs4.SurveyYear, qAbs4.Park, qAbs4.IslandCode, qAbs4.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, qAbs4.Nativity, " & _
    "Avg(qAbs4.AbsCover) AS Average, StDev(qAbs4.AbsCover) AS StdDev, Min(qAbs4.AbsCover) AS MinRange, Max(qAbs4.AbsCover) AS MaxRange " & _
    "FROM (" & strAbs4raw & ") AS qAbs4 " & _
    "GROUP BY qAbs4.SurveyYear, qAbs4.Park, qAbs4.IslandCode, qAbs4.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & ", qAbs4.Nativity"
'5 Absolute cover FXN --------------------------------------------------------
strAbs5raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.FxnGroup, (Count([Species_Code])/[NofPoints])*100 AS AbsCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRaw.SiteCode = qTotalPoints.SiteCode) " & _
    "AND (qRaw.Park = qTotalPoints.Park) AND (qRaw.IslandCode = qTotalPoints.IslandCode) AND (qRaw.SurveyYear = qTotalPoints.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.FxnGroup, qTotalPoints.NofPoints"

strAbs5 = "SELECT qAbs5.SurveyYear, qAbs5.Park, qAbs5.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, qAbs5.FxnGroup, " & Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(qAbs5.AbsCover) AS Average, StDev(qAbs5.AbsCover) AS StdDev, Min(qAbs5.AbsCover) AS MinRange, Max(qAbs5.AbsCover) AS MaxRange " & _
    "FROM (" & strAbs5raw & ") AS qAbs5 " & _
    "GROUP BY qAbs5.SurveyYear, qAbs5.Park, qAbs5.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", qAbs5.FxnGroup, " & Chr$(34) & "All" & Chr$(34)
'6 Absolute cover FXN NAT ----------------------------------------------------
strAbs6raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.FxnGroup, qRaw.Nativity, (Count([Species_Code])/[NofPoints])*100 AS AbsCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRaw.SiteCode = qTotalPoints.SiteCode) " & _
    "AND (qRaw.Park = qTotalPoints.Park) AND (qRaw.IslandCode = qTotalPoints.IslandCode) AND (qRaw.SurveyYear = qTotalPoints.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.FxnGroup, qRaw.Nativity, qTotalPoints.NofPoints"

strAbs6 = "SELECT qAbs6.SurveyYear, qAbs6.Park, qAbs6.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, qAbs6.FxnGroup, qAbs6.Nativity, " & _
    "Avg(qAbs6.AbsCover) AS Average, StDev(qAbs6.AbsCover) AS StdDev, Min(qAbs6.AbsCover) AS MinRange, Max(qAbs6.AbsCover) AS MaxRange " & _
    "FROM (" & strAbs6raw & ") AS qAbs6 " & _
    "GROUP BY qAbs6.SurveyYear, qAbs6.Park, qAbs6.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", qAbs6.FxnGroup, qAbs6.Nativity"
'7 Absolute cover FXN VEG ----------------------------------------------------
strAbs7raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, (Count([Species_Code])/[NofPoints])*100 AS AbsCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRaw.SiteCode = qTotalPoints.SiteCode) " & _
    "AND (qRaw.Park = qTotalPoints.Park) AND (qRaw.IslandCode = qTotalPoints.IslandCode) AND (qRaw.SurveyYear = qTotalPoints.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qTotalPoints.NofPoints"

strAbs7 = "SELECT qAbs7.SurveyYear, qAbs7.Park, qAbs7.IslandCode, qAbs7.Vegetation_Community, qAbs7.FxnGroup, " & Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(qAbs7.AbsCover) AS Average, StDev(qAbs7.AbsCover) AS StdDev, Min(qAbs7.AbsCover) AS MinRange, Max(qAbs7.AbsCover) AS MaxRange " & _
    "FROM (" & strAbs7raw & ") AS qAbs7 " & _
    "GROUP BY qAbs7.SurveyYear, qAbs7.Park, qAbs7.IslandCode, qAbs7.Vegetation_Community, qAbs7.FxnGroup, " & Chr$(34) & "All" & Chr$(34)
'8 Absolute cover FXN NAT VEG ------------------------------------------------
strAbs8raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qRaw.Nativity, (Count([Species_Code])/[NofPoints])*100 AS AbsCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRaw.SiteCode = qTotalPoints.SiteCode) " & _
    "AND (qRaw.Park = qTotalPoints.Park) AND (qRaw.IslandCode = qTotalPoints.IslandCode) AND (qRaw.SurveyYear = qTotalPoints.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qRaw.Nativity, qTotalPoints.NofPoints"

strAbs8 = "SELECT qAbs8.SurveyYear, qAbs8.Park, qAbs8.IslandCode, qAbs8.Vegetation_Community, qAbs8.FxnGroup, qAbs8.Nativity, " & _
    "Avg(qAbs8.AbsCover) AS Average, StDev(qAbs8.AbsCover) AS StdDev, Min(qAbs8.AbsCover) AS MinRange, Max(qAbs8.AbsCover) AS MaxRange " & _
    "FROM (" & strAbs8raw & ") AS qAbs8 " & _
    "GROUP BY qAbs8.SurveyYear, qAbs8.Park, qAbs8.IslandCode, qAbs8.Vegetation_Community, qAbs8.FxnGroup, qAbs8.Nativity"
' ----------------------------------------------------------------------------

' Absolute Cover UNION Table E7 ----------------------------------------------
strAbsUnion = "SELECT qry.SurveyYear, " & ParkSelect(xPark) & ", qry.Vegetation_Community, qry.FxnGroup, qry.Nativity, qry.Average, qry.StdDev, qry.MinRange, qry.MaxRange, " & _
    Chr$(34) & "Annual Report, Absolute Cover by Group (Table E7)" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT * FROM (" & strAbs1 & ") AS q1 UNION SELECT * FROM (" & strAbs2 & ") AS q2 UNION SELECT * FROM (" & strAbs3 & ") AS q3 UNION SELECT * FROM (" & strAbs4 & ") AS q4 " & _
    "UNION SELECT * FROM (" & strAbs5 & ") AS q5 UNION SELECT * FROM (" & strAbs6 & ") AS q6 UNION SELECT * FROM (" & strAbs7 & ") AS q7 UNION SELECT * FROM (" & strAbs8 & ") AS q8) AS qry"

Export_AnnualReport_AbsoluteCover_byGroup = strAbsUnion

End Function

Function Export_AnnualReport_RelativeCover_byGroup(xPark As Integer, xYear As Integer)

Dim qdf As DAO.QueryDef

Dim strWhere As String

Dim strRaw As String

Dim strRel1raw As String
Dim strRel2raw As String
Dim strRel3raw As String
Dim strRel4raw As String
Dim strRel5raw As String
Dim strRel6raw As String
Dim strRel7raw As String
Dim strRel8raw As String

Dim strRel1 As String
Dim strRel2 As String
Dim strRel3 As String
Dim strRel4 As String
Dim strRel5 As String
Dim strRel6 As String
Dim strRel7 As String
Dim strRel8 As String

Dim strRelUnion As String 'final SQL string
' ----------------------------------------------------------------------------

For Each qdf In CurrentDb.QueryDefs
    If qdf.Name = "qdfTemp" Then
        CurrentDb.QueryDefs.Delete "qdfTemp"
    End If
Next qdf

' Create WHERE string --------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ") AND ((tlu_Condition.Analysis_code) Is Null Or (tlu_Condition.Analysis_code)=" & Chr$(34) & "Alive" & Chr$(34) & ")"
' ----------------------------------------------------------------------------

' Create data string ---------------------------------------------------------
strRaw = "SELECT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "tbl_Species_Data.Species_Code, tlu_Condition.Analysis_code AS Condition, Park_Spp.FxnGroup, Park_Spp.Nativity " & _
    "FROM ((tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point LEFT JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN tlu_Condition " & _
    "ON tbl_Species_Data.Condition = tlu_Condition.Condition) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON Park_Spp.Species_code = tbl_Species_Data.Species_Code " & _
    "WHERE (" & strWhere & ")"
'-----------------------------------------------------------------------------

'1 Relative cover ALL --------------------------------------------------------
strRel1raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, (Count([Species_Code])/[NofHits])*100 AS RelCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalLiveHitsSQL(xPark) & ") AS qTotalLiveHits ON (qRaw.SiteCode = qTotalLiveHits.SiteCode) " & _
    "AND (qRaw.Park = qTotalLiveHits.Park) AND (qRaw.IslandCode = qTotalLiveHits.IslandCode) AND (qRaw.SurveyYear = qTotalLiveHits.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qTotalLiveHits.NofHits"

strRel1 = "SELECT qRel1.SurveyYear, qRel1.Park, qRel1.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, " & Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(qRel1.RelCover) AS Average, StDev(qRel1.RelCover) AS StdDev, Min(qRel1.RelCover) AS MinRange, Max(qRel1.RelCover) AS MaxRange " & _
    "FROM (" & strRel1raw & ") AS qRel1 " & _
    "GROUP BY qRel1.SurveyYear, qRel1.Park, qRel1.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34) & ""
'2 Relative cover ALL NAT ----------------------------------------------------
strRel2raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Nativity, (Count([Species_Code])/[NofHits])*100 AS RelCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalLiveHitsSQL(xPark) & ") AS qTotalLiveHits ON (qRaw.SiteCode = qTotalLiveHits.SiteCode) " & _
    "AND (qRaw.Park = qTotalLiveHits.Park) AND (qRaw.IslandCode = qTotalLiveHits.IslandCode) AND (qRaw.SurveyYear = qTotalLiveHits.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Nativity, qTotalLiveHits.NofHits"

strRel2 = "SELECT qRel2.SurveyYear, qRel2.Park, qRel2.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, qRel2.Nativity, " & _
    "Avg(qRel2.RelCover) AS Average, StDev(qRel2.RelCover) AS StdDev, Min(qRel2.RelCover) AS MinRange, Max(qRel2.RelCover) AS MaxRange " & _
    "FROM (" & strRel2raw & ") AS qRel2 " & _
    "GROUP BY qRel2.SurveyYear, qRel2.Park, qRel2.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34) & ", qRel2.Nativity"

'3 Relative cover ALL VEG ----------------------------------------------------
strRel3raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, (Count([Species_Code])/[NofHits])*100 AS RelCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalLiveHitsSQL(xPark) & ") AS qTotalLiveHits ON (qRaw.SiteCode = qTotalLiveHits.SiteCode) " & _
    "AND (qRaw.Park = qTotalLiveHits.Park) AND (qRaw.IslandCode = qTotalLiveHits.IslandCode) AND (qRaw.SurveyYear = qTotalLiveHits.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qTotalLiveHits.NofHits"

strRel3 = "SELECT qRel3.SurveyYear, qRel3.Park, qRel3.IslandCode, qRel3.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, " & Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(qRel3.RelCover) AS Average, StDev(qRel3.RelCover) AS StdDev, Min(qRel3.RelCover) AS MinRange, Max(qRel3.RelCover) AS MaxRange " & _
    "FROM (" & strRel3raw & ") AS qRel3 " & _
    "GROUP BY qRel3.SurveyYear, qRel3.Park, qRel3.IslandCode, qRel3.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34) & ""

'4 Relative cover ALL NAT VEG ------------------------------------------------
strRel4raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.Nativity, (Count([Species_Code])/[NofHits])*100 AS RelCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalLiveHitsSQL(xPark) & ") AS qTotalLiveHits ON (qRaw.SiteCode = qTotalLiveHits.SiteCode) " & _
    "AND (qRaw.Park = qTotalLiveHits.Park) AND (qRaw.IslandCode = qTotalLiveHits.IslandCode) AND (qRaw.SurveyYear = qTotalLiveHits.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.Nativity, qTotalLiveHits.NofHits"

strRel4 = "SELECT qRel4.SurveyYear, qRel4.Park, qRel4.IslandCode, qRel4.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, qRel4.Nativity, " & _
    "Avg(qRel4.RelCover) AS Average, StDev(qRel4.RelCover) AS StdDev, Min(qRel4.RelCover) AS MinRange, Max(qRel4.RelCover) AS MaxRange " & _
    "FROM (" & strRel4raw & ") AS qRel4 " & _
    "GROUP BY qRel4.SurveyYear, qRel4.Park, qRel4.IslandCode, qRel4.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & ", qRel4.Nativity"

'5 Relative cover FXN --------------------------------------------------------
strRel5raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.FxnGroup, (Count([Species_Code])/[NofHits])*100 AS RelCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalLiveHitsSQL(xPark) & ") AS qTotalLiveHits ON (qRaw.SiteCode = qTotalLiveHits.SiteCode) " & _
    "AND (qRaw.Park = qTotalLiveHits.Park) AND (qRaw.IslandCode = qTotalLiveHits.IslandCode) AND (qRaw.SurveyYear = qTotalLiveHits.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.FxnGroup, qTotalLiveHits.NofHits"

strRel5 = "SELECT qRel5.SurveyYear, qRel5.Park, qRel5.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, qRel5.FxnGroup, " & Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(qRel5.RelCover) AS Average, StDev(qRel5.RelCover) AS StdDev, Min(qRel5.RelCover) AS MinRange, Max(qRel5.RelCover) AS MaxRange " & _
    "FROM (" & strRel5raw & ") AS qRel5 " & _
    "GROUP BY qRel5.SurveyYear, qRel5.Park, qRel5.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", qRel5.FxnGroup, " & Chr$(34) & "All" & Chr$(34)

'6 Relative cover FXN NAT ----------------------------------------------------
strRel6raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.FxnGroup, qRaw.Nativity, (Count([Species_Code])/[NofHits])*100 AS RelCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalLiveHitsSQL(xPark) & ") AS qTotalLiveHits ON (qRaw.SiteCode = qTotalLiveHits.SiteCode) " & _
    "AND (qRaw.Park = qTotalLiveHits.Park) AND (qRaw.IslandCode = qTotalLiveHits.IslandCode) AND (qRaw.SurveyYear = qTotalLiveHits.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.FxnGroup, qRaw.Nativity, qTotalLiveHits.NofHits"

strRel6 = "SELECT qRel6.SurveyYear, qRel6.Park, qRel6.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, qRel6.FxnGroup, qRel6.Nativity, " & _
    "Avg(qRel6.RelCover) AS Average, StDev(qRel6.RelCover) AS StdDev, Min(qRel6.RelCover) AS MinRange, Max(qRel6.RelCover) AS MaxRange " & _
    "FROM (" & strRel6raw & ") AS qRel6 " & _
    "GROUP BY qRel6.SurveyYear, qRel6.Park, qRel6.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", qRel6.FxnGroup, qRel6.Nativity"

'7 Relative cover FXN VEG ----------------------------------------------------
strRel7raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, (Count([Species_Code])/[NofHits])*100 AS RelCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalLiveHitsSQL(xPark) & ") AS qTotalLiveHits ON (qRaw.SiteCode = qTotalLiveHits.SiteCode) " & _
    "AND (qRaw.Park = qTotalLiveHits.Park) AND (qRaw.IslandCode = qTotalLiveHits.IslandCode) AND (qRaw.SurveyYear = qTotalLiveHits.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qTotalLiveHits.NofHits"

strRel7 = "SELECT qRel7.SurveyYear, qRel7.Park, qRel7.IslandCode, qRel7.Vegetation_Community, qRel7.FxnGroup, " & Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(qRel7.RelCover) AS Average, StDev(qRel7.RelCover) AS StdDev, Min(qRel7.RelCover) AS MinRange, Max(qRel7.RelCover) AS MaxRange " & _
    "FROM (" & strRel7raw & ") AS qRel7 " & _
    "GROUP BY qRel7.SurveyYear, qRel7.Park, qRel7.IslandCode, qRel7.Vegetation_Community, qRel7.FxnGroup, " & Chr$(34) & "All" & Chr$(34)

'8 Relative cover FXN NAT VEG ------------------------------------------------
strRel8raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qRaw.Nativity, (Count([Species_Code])/[NofHits])*100 AS RelCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalLiveHitsSQL(xPark) & ") AS qTotalLiveHits ON (qRaw.SiteCode = qTotalLiveHits.SiteCode) " & _
    "AND (qRaw.Park = qTotalLiveHits.Park) AND (qRaw.IslandCode = qTotalLiveHits.IslandCode) AND (qRaw.SurveyYear = qTotalLiveHits.SurveyYear) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qRaw.Nativity, qTotalLiveHits.NofHits"

strRel8 = "SELECT qRel8.SurveyYear, qRel8.Park, qRel8.IslandCode, qRel8.Vegetation_Community, qRel8.FxnGroup, qRel8.Nativity, " & _
    "Avg(qRel8.RelCover) AS Average, StDev(qRel8.RelCover) AS StdDev, Min(qRel8.RelCover) AS MinRange, Max(qRel8.RelCover) AS MaxRange " & _
    "FROM (" & strRel8raw & ") AS qRel8 " & _
    "GROUP BY qRel8.SurveyYear, qRel8.Park, qRel8.IslandCode, qRel8.Vegetation_Community, qRel8.FxnGroup, qRel8.Nativity"
' ----------------------------------------------------------------------------

'Relative Cover UNION Table EX -----------------------------------------------
strRelUnion = "SELECT qry.SurveyYear, qry.Park, qRaw.IslandCode, qry.Vegetation_Community, qry.FxnGroup, qry.Nativity, qry.Average, qry.StdDev, qry.MinRange, qry.MaxRange, " & _
    Chr$(34) & "Annual Report, Relative Cover by Group (Table E7)" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT * FROM (" & strRel1 & ") AS q1 UNION SELECT * FROM (" & strRel2 & ") AS q2 UNION SELECT * FROM (" & strRel3 & ") AS q3 UNION SELECT * FROM (" & strRel4 & ") AS q4 " & _
    "UNION SELECT * FROM (" & strRel5 & ") AS q5 UNION SELECT * FROM (" & strRel6 & ") AS q6 UNION SELECT * FROM (" & strRel7 & ") AS q7 UNION SELECT * FROM (" & strRel8 & ") AS q8) AS qry"

Set qdf = CurrentDb.CreateQueryDef("qdfTemp", strRelUnion)

Export_AnnualReport_RelativeCover_byGroup = "SELECT * FROM qdfTemp"

End Function

Function Export_AnnualReport_SpeciesRichness(xPark As Integer, xYear As Integer)
'Table E6

Dim qdf As DAO.QueryDef

Dim strWhere As String

Dim strRaw As String

Dim str1Raw As String 'All
Dim str2Raw As String 'All Nat
Dim str3Raw As String 'All Veg
Dim str4Raw As String 'All Veg Nat
Dim str5Raw As String 'Fxn
Dim str6Raw As String 'Fxn Nat
Dim str7Raw As String 'Fxn Veg
Dim str8Raw As String 'Fxn Veg Nat

Dim str1 As String 'All
Dim str2 As String 'All Nat
Dim str3 As String 'All Veg
Dim str4 As String 'All Veg Nat
Dim str5 As String 'Fxn
Dim str6 As String 'Fxn Nat
Dim str7 As String 'Fxn Veg
Dim str8 As String 'Fxn Veg Nat

Dim strRichness As String 'final SQL string
'-----------------------------------------------------------------------------

' Delete qdfTemp
For Each qdf In CurrentDb.QueryDefs
    If qdf.Name = "qdfTemp" Then
        CurrentDb.QueryDefs.Delete "qdfTemp"
    End If
Next qdf
'-----------------------------------------------------------------------------

' Create WHERE string --------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ")"
' ----------------------------------------------------------------------------

' Create data strings --------------------------------------------------------
strRaw = "SELECT qryUNION.Event_ID, qryUNION.Species_Code, Sum(qryUNION.Analysis_value) AS IsPresent " & _
    "FROM (SELECT tbl.Event_ID, tbl.Species_Code, qry11.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry11 INNER JOIN tbl_Phenology_Species AS tbl ON qry11.Richness_code = tbl.Plot_11 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry1.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry1 INNER JOIN tbl_Phenology_Species AS tbl ON qry1.Richness_code = tbl.Plot_1 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry21.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry21 INNER JOIN tbl_Phenology_Species AS tbl  ON qry21.Richness_code = tbl.Plot_21 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry2.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry2 INNER JOIN tbl_Phenology_Species AS tbl  ON qry2.Richness_code = tbl.Plot_2 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry31.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry31 INNER JOIN tbl_Phenology_Species AS tbl  ON qry31.Richness_code = tbl.Plot_31 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry3.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry3 INNER JOIN tbl_Phenology_Species AS tbl  ON qry3.Richness_code = tbl.Plot_3 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry41.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry41 INNER JOIN tbl_Phenology_Species AS tbl  ON qry41.Richness_code = tbl.Plot_41 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry4.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry4 INNER JOIN tbl_Phenology_Species AS tbl  ON qry4.Richness_code = tbl.Plot_4 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry51.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry51 INNER JOIN tbl_Phenology_Species AS tbl  ON qry51.Richness_code = tbl.Plot_51 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry5.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry5 INNER JOIN tbl_Phenology_Species AS tbl  ON qry5.Richness_code = tbl.Plot_5 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry61.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry61 INNER JOIN tbl_Phenology_Species AS tbl  ON qry61.Richness_code = tbl.Plot_61 " & _
    "UNION SELECT tbl.Event_ID, tbl.Species_Code, qry6.Analysis_value FROM (SELECT * FROM tlu_Richness)  AS qry6 INNER JOIN tbl_Phenology_Species AS tbl  ON qry6.Richness_code = tbl.Plot_6)  AS qryUNION " & _
    "GROUP BY qryUNION.Event_ID, qryUNION.Species_Code HAVING (((qryUNION.Event_ID) Is Not Null) AND ((Sum(qryUNION.Analysis_value))<>0))"
' ----------------------------------------------------------------------------

'All--------------------------------------------------------------------------
str1Raw = "SELECT qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, Count(qry.Species_Code) AS N_Species " & _
    "FROM (SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, tbl_Events.Event_ID, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date, qRichRaw.Species_Code FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events " & _
    "INNER JOIN (" & strRaw & ") AS qRichRaw ON tbl_Events.Event_ID = qRichRaw.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON  qRichRaw.Species_Code = Park_Spp.Species_code WHERE (" & strWhere & ")) AS qry " & _
    "GROUP BY qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode"

str1 = "SELECT q1.SurveyYear, q1.Park, q1.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, " & _
    Chr$(34) & "All" & Chr$(34) & " AS Nativity, Avg(q1.N_Species) AS Average, StDev(q1.N_Species) AS StdDev, Min(q1.N_Species) AS MinRange, Max(q1.N_Species) AS MaxRange " & _
    "FROM (" & str1Raw & ") AS q1 " & _
    "GROUP BY q1.SurveyYear, q1.Park, q1.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34)

' All Nat--------------------------------------------------------------------------
str2Raw = "SELECT qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.Nativity, Count(qry.Species_Code) AS N_Species " & _
    "FROM (SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, tbl_Events.Event_ID, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date, qRichRaw.Species_Code, Park_Spp.Nativity FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events " & _
    "INNER JOIN (" & strRaw & ") AS qRichRaw ON tbl_Events.Event_ID = qRichRaw.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON qRichRaw.Species_Code = Park_Spp.Species_code WHERE (" & strWhere & ")) AS qry " & _
    "GROUP BY qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.Nativity"

str2 = "SELECT q2.SurveyYear, q2.Park, q2.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, " & _
    "q2.Nativity, Avg(q2.N_Species) AS Average, StDev(q2.N_Species) AS StdDev, Min(q2.N_Species) AS MinRange, Max(q2.N_Species) AS MaxRange " & _
    "FROM (" & str2Raw & ") AS q2 " & _
    "GROUP BY q2.SurveyYear, q2.Park, q2.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", q2.Nativity, " & Chr$(34) & "All" & Chr$(34)

' All Veg--------------------------------------------------------------------------
str3Raw = "SELECT qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.Vegetation_Community, Count(qry.Species_Code) AS N_Species " & _
    "FROM (SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, tbl_Events.Event_ID, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date, qRichRaw.Species_Code FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events " & _
    "INNER JOIN (" & strRaw & ") AS qRichRaw ON tbl_Events.Event_ID = qRichRaw.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON qRichRaw.Species_Code = Park_Spp.Species_code WHERE (" & strWhere & ")) AS qry " & _
    "GROUP BY qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.Vegetation_Community"

str3 = "SELECT q3.SurveyYear, q3.Park, q3.IslandCode, q3.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, " & Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(q3.N_Species) AS Average, StDev(q3.N_Species) AS StdDev, Min(q3.N_Species) AS MinRange, Max(q3.N_Species) AS MaxRange " & _
    "FROM (" & str3Raw & ") AS q3 " & _
    "GROUP BY q3.SurveyYear, q3.Park, q3.IslandCode, q3.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & ", " & Chr$(34) & "All" & Chr$(34)

'All Veg Nat--------------------------------------------------------------------------
str4Raw = "SELECT qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.Vegetation_Community, qry.Nativity, Count(qry.Species_Code) AS N_Species " & _
    "FROM (SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, tbl_Events.Event_ID, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date, qRichRaw.Species_Code, Park_Spp.Nativity FROM (tbl_Sites INNER JOIN (tbl_Locations  INNER JOIN (tbl_Events " & _
    "INNER JOIN (" & strRaw & ") AS qRichRaw ON tbl_Events.Event_ID = qRichRaw.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON qRichRaw.Species_Code = Park_Spp.Species_code WHERE (" & strWhere & "))  AS qry " & _
    "GROUP BY qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.Vegetation_Community, qry.Nativity"

str4 = "SELECT q4.SurveyYear, q4.Park, q4.IslandCode, q4.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & " AS FxnGroup, q4.Nativity, " & _
    "Avg(q4.N_Species) AS Average, StDev(q4.N_Species) AS StdDev, Min(q4.N_Species) AS MinRange, Max(q4.N_Species) AS MaxRange " & _
    "FROM (" & str4Raw & ") AS q4 " & _
    "GROUP BY q4.SurveyYear, q4.Park, q4.IslandCode, q4.Vegetation_Community, " & Chr$(34) & "All" & Chr$(34) & ", q4.Nativity"

'Fxn--------------------------------------------------------------------------
str5Raw = "SELECT qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.FxnGroup, Count(qry.Species_Code) AS N_Species " & _
    "FROM (SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, tbl_Events.Event_ID, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date, qRichRaw.Species_Code, Park_Spp.FxnGroup FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events " & _
    "INNER JOIN (" & strRaw & ") AS qRichRaw ON tbl_Events.Event_ID = qRichRaw.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON qRichRaw.Species_code = Park_Spp.Species_Code WHERE (" & strWhere & "))  AS qry " & _
    "GROUP BY qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.FxnGroup"

str5 = "SELECT q5.SurveyYear, q5.Park, q5.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, q5.FxnGroup, " & Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(q5.N_Species) AS Average, StDev(q5.N_Species) AS StdDev, Min(q5.N_Species) AS MinRange, Max(q5.N_Species) AS MaxRange " & _
    "FROM (" & str5Raw & ") AS q5 " & _
    "GROUP BY q5.SurveyYear, q5.Park, q5.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", q5.FxnGroup, " & Chr$(34) & "All" & Chr$(34)

'Fxn Nat--------------------------------------------------------------------------
str6Raw = "SELECT qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.FxnGroup, qry.Nativity, Count(qry.Species_Code) AS N_Species " & _
    "FROM (SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode,tbl_Locations.Vegetation_Community, tbl_Events.Event_ID, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date, qRichRaw.Species_Code, Park_Spp.FxnGroup, Park_Spp.Nativity FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events " & _
    "INNER JOIN (" & strRaw & ") AS qRichRaw ON tbl_Events.Event_ID = qRichRaw.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON qRichRaw.Species_code = Park_Spp.Species_Code WHERE (" & strWhere & ")) AS qry " & _
    "GROUP BY qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.FxnGroup, qry.Nativity"

str6 = "SELECT q6.SurveyYear, q6.Park, q6.IslandCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, q6.FxnGroup, q6.Nativity, " & _
    "Avg(q6.N_Species) AS Average, StDev(q6.N_Species) AS StdDev, Min(q6.N_Species) AS MinRange, Max(q6.N_Species) AS MaxRange " & _
    "FROM (" & str6Raw & ") AS q6 " & _
    "GROUP BY q6.SurveyYear, q6.Park, q6.IslandCode, " & Chr$(34) & "All" & Chr$(34) & ", q6.FxnGroup, q6.Nativity"

'Fxn Veg--------------------------------------------------------------------------
str7Raw = "SELECT qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.Vegetation_Community, qry.FxnGroup, Count(qry.Species_Code) AS N_Species " & _
    "FROM (SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, tbl_Events.Event_ID, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date, qRichRaw.Species_Code, Park_Spp.FxnGroup FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events " & _
    "INNER JOIN (" & strRaw & ") AS qRichRaw ON tbl_Events.Event_ID = qRichRaw.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON qRichRaw.Species_code = Park_Spp.Species_Code WHERE (" & strWhere & "))  AS qry " & _
    "GROUP BY qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.Vegetation_Community, qry.FxnGroup"

str7 = "SELECT q7.SurveyYear, q7.Park, q7.IslandCode, q7.Vegetation_Community, q7.FxnGroup, " & Chr$(34) & "All" & Chr$(34) & " AS Nativity, " & _
    "Avg(q7.N_Species) AS Average, StDev(q7.N_Species) AS StdDev, Min(q7.N_Species) AS MinRange, Max(q7.N_Species) AS MaxRange " & _
    "FROM (" & str7Raw & ") AS q7 " & _
    "GROUP BY q7.SurveyYear, q7.Park, q7.IslandCode, q7.Vegetation_Community, q7.FxnGroup, " & Chr$(34) & "All" & Chr$(34)

'Fxn Veg Nat--------------------------------------------------------------------------
str8Raw = "SELECT qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.Vegetation_Community, qry.FxnGroup, qry.Nativity, Count(qry.Species_Code) AS N_Species " & _
    "FROM (SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, tbl_Events.Event_ID, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date, qRichRaw.Species_Code, Park_Spp.FxnGroup, Park_Spp.Nativity FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events " & _
    "INNER JOIN (" & strRaw & ") AS qRichRaw ON tbl_Events.Event_ID = qRichRaw.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON qRichRaw.Species_code = Park_Spp.Species_Code WHERE (" & strWhere & "))  AS qry " & _
    "GROUP BY qry.SurveyYear, qry.Park, qry.IslandCode, qry.SiteCode, qry.Vegetation_Community, qry.FxnGroup, qry.Nativity"

str8 = "SELECT q8.SurveyYear, q8.Park, q8.IslandCode, q8.Vegetation_Community, q8.FxnGroup, q8.Nativity, " & _
    "Avg(q8.N_Species) AS Average, StDev(q8.N_Species) AS StdDev, Min(q8.N_Species) AS MinRange, Max(q8.N_Species) AS MaxRange " & _
    "FROM (" & str8Raw & ") AS q8 " & _
    "GROUP BY q8.SurveyYear, q8.Park, q8.IslandCode, q8.Vegetation_Community, q8.FxnGroup, q8.Nativity"
' ----------------------------------------------------------------------------

'Richness UNION Table E6--------------------------------------------------------------------------
strRichness = "SELECT qry.SurveyYear, " & ParkSelect(xPark) & ", qry.Vegetation_Community, qry.FxnGroup, qry.Nativity, qry.Average, qry.StdDev, qry.MinRange, qry.MaxRange, " & _
    Chr$(34) & "Annual Report, Species Richness (Table E6)" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT * FROM (" & str1 & ") AS qry1 UNION SELECT * FROM (" & str2 & ") AS qry2 UNION SELECT * FROM (" & str3 & ") AS qry3 " & _
    "UNION SELECT * FROM (" & str4 & ") AS qry4 UNION SELECT * FROM (" & str5 & ") AS qry5 UNION SELECT * FROM (" & str6 & ") AS qry6 " & _
    "UNION SELECT * FROM (" & str7 & ") AS qry7 UNION SELECT * FROM (" & str8 & ") AS qry8)  AS qry"

Set qdf = CurrentDb.CreateQueryDef("qdfTemp", strRichness)

Export_AnnualReport_SpeciesRichness = "SELECT * FROM qdfTemp"

End Function

Function Export_AnnualReport_ShrubData(xPark As Integer, xYear As Integer, stType As String)

Dim strWhere As String

Dim strShrubInd As String
Dim strShrubStems As String
Dim strShrubDeadStems As String

Dim strShrubData As String
Dim strShrubDensity As String
'-----------------------------------------------------------------------------

'Create WHERE string ---------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ")"
'-----------------------------------------------------------------------------

' Create data strings --------------------------------------------------------
strShrubInd = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, Year([Start_Date]) AS SurveyYear, " & _
    "tbl_Events.Start_Date AS SurveyDate, tbl_Shrub_Data.Species_Code, tbl_Shrub_Count.Age_Category, tlu_Shrub_Plot_Number.Plot, tbl_Shrub_Count.Number_of_Individuals AS N " & _
    "FROM tbl_Sites INNER JOIN ((tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Shrub_Data ON tbl_Events.Event_ID = tbl_Shrub_Data.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "INNER JOIN (tlu_Shrub_Plot_Number INNER JOIN tbl_Shrub_Count ON tlu_Shrub_Plot_Number.[Shrub_Plot_ID] = tbl_Shrub_Count.Plot_Number) " & _
    "ON tbl_Shrub_Data.Shrub_Data_ID = tbl_Shrub_Count.Shrub_Data_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & strWhere & ")"

strShrubStems = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, Year([Start_Date]) AS SurveyYear, " & _
    "tbl_Events.Start_Date AS SurveyDate, tbl_Shrub_Data.Species_Code, " & Chr$(34) & "NoOfStems" & Chr$(34) & " AS Age_Category, tlu_Shrub_Plot_Number.Plot, tbl_Shrub_Count.Number_of_Stems AS N " & _
    "FROM tbl_Sites INNER JOIN ((tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Shrub_Data ON tbl_Events.Event_ID = tbl_Shrub_Data.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "INNER JOIN (tlu_Shrub_Plot_Number INNER JOIN tbl_Shrub_Count ON tlu_Shrub_Plot_Number.[Shrub_Plot_ID] = tbl_Shrub_Count.Plot_Number) " & _
    "ON tbl_Shrub_Data.Shrub_Data_ID = tbl_Shrub_Count.Shrub_Data_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (((tbl_Shrub_Count.Number_of_Stems) Is Not Null) AND " & strWhere & ")"

strShrubDeadStems = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, Year([Start_Date]) AS SurveyYear, " & _
    "tbl_Events.Start_Date AS SurveyDate, tbl_Shrub_Data.Species_Code, " & Chr$(34) & "NoOfDeadStems" & Chr$(34) & " AS Age_Category, tlu_Shrub_Plot_Number.Plot, tbl_Shrub_Count.Number_of_Dead_Stems AS N " & _
    "FROM tbl_Sites INNER JOIN ((tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Shrub_Data ON tbl_Events.Event_ID = tbl_Shrub_Data.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "INNER JOIN (tlu_Shrub_Plot_Number INNER JOIN tbl_Shrub_Count ON tlu_Shrub_Plot_Number.[Shrub_Plot_ID] = tbl_Shrub_Count.Plot_Number) " & _
    "ON tbl_Shrub_Data.Shrub_Data_ID = tbl_Shrub_Count.Shrub_Data_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (((tbl_Shrub_Count.Number_of_Dead_Stems) Is Not Null) AND " & strWhere & ")"
'-----------------------------------------------------------------------------

' Count data string ----------
strShrubData = "TRANSFORM Sum(qry.N) AS SumN " & _
    "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Species_Code, qry.Age_Category, " & _
    Chr$(34) & "Annual Report, Shrub Count " & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT * FROM (" & strShrubInd & ") as qInd UNION SELECT * FROM (" & strShrubStems & ") AS qStems UNION SELECT * FROM (" & strShrubDeadStems & ") AS qDeadSteams) AS qry " & _
    "GROUP BY " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Species_Code, qry.Age_Category " & _
    "PIVOT qry.Plot IN (" & Chr$(34) & "0-5m" & Chr$(34) & ", " & Chr$(34) & "5-10m" & Chr$(34) & ", " & Chr$(34) & "10-15m" & Chr$(34) & ", " & Chr$(34) & "15-20m" & Chr$(34) & ", " & _
    Chr$(34) & "20-25m" & Chr$(34) & ", " & Chr$(34) & "25-30m" & Chr$(34) & ")"
' Density data string --------
strShrubDensity = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Species_Code, (Sum(qry.N)/30) AS Shrub_Density, " & _
    Chr$(34) & "Annual Report, Shrub Density" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT * FROM (" & strShrubInd & ") as qInd UNION SELECT * FROM (" & strShrubStems & ") AS qStems UNION SELECT * FROM (" & strShrubDeadStems & ") AS qDeadSteams) AS qry " & _
    "GROUP BY " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Species_Code"
'-----------------------------------------------------------------------------

' Export dataset based on input parameters -----------------------------------
Select Case stType
    Case Is = "Count"
        Export_AnnualReport_ShrubData = strShrubData
    Case Is = "Density"
        Export_AnnualReport_ShrubData = strShrubDensity
    Case Else
        MsgBox "Error!"
        Export_AnnualReport_ShrubData = ""
End Select

End Function

Function Export_AnnualReport_TreeData(xPark As Integer, xYear As Integer, stType As String)

Dim strWhere As String

Dim strTreeDBH As String
Dim strTreeData As String
'-----------------------------------------------------------------------------

' Create WHERE string ---------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ")"
'-----------------------------------------------------------------------------

' Create data strings --------------------------------------------------------
' DBH ------------------------------------------------------------------------
strTreeDBH = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Species_Code, qry.Age_Category, qry.Species_Type, qry.DBH_cm, " & _
    Chr$(34) & "Annual Report, Tree DBH" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, Year([Start_Date]) AS SurveyYear, " & _
    "tbl_Events.Start_Date AS SurveyDate, tbl_Tree_Data.Species_Code, tbl_Tree_Count.Age_Category, tbl_Tree_Data.Species_Type, tbl_Tree_DBH.DBH_cm " & _
    "FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Tree_Data ON tbl_Events.Event_ID = tbl_Tree_Data.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) INNER JOIN (tbl_Tree_Count INNER JOIN tbl_Tree_DBH ON tbl_Tree_Count.Tree_Count_ID = tbl_Tree_DBH.Tree_Count_ID) " & _
    "ON tbl_Tree_Data.Tree_Data_ID = tbl_Tree_Count.Tree_Data_ID WHERE (" & strWhere & ")) AS qry"
' Count ----------------------------------------------------------------------
strTreeData = "TRANSFORM Sum(qry.SumCount) AS SumOfSumCount " & _
    "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Species_Code, qry.Age_Category, " & _
    Chr$(34) & "Annual Report, Tree Count" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, Year([Start_Date]) AS SurveyYear, " & _
    "tbl_Events.Start_Date AS SurveyDate, tbl_Tree_Data.Species_Code, tbl_Tree_Count.Age_Category, tbl_Tree_Data.Species_Type, tbl_Tree_Count.Count_of_Units AS SumCount " & _
    "FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Tree_Data ON tbl_Events.Event_ID = tbl_Tree_Data.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID) INNER JOIN tbl_Tree_Count ON tbl_Tree_Data.Tree_Data_ID = tbl_Tree_Count.Tree_Data_ID " & _
    "WHERE (" & strWhere & ")) AS qry " & _
    "GROUP BY " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.Species_Code, qry.Age_Category " & _
    "PIVOT qry.Species_Type IN (" & Chr$(34) & "Tree" & Chr$(34) & ", " & Chr$(34) & "Shrub" & Chr$(34) & ")"
'-----------------------------------------------------------------------------

' Export dataset based on input parameters -----------------------------------
Select Case stType
    Case Is = "DBH"
        Export_AnnualReport_TreeData = strTreeDBH
    Case Is = "Count"
        Export_AnnualReport_TreeData = strTreeData
    Case Else
        MsgBox "Error!"
        Export_AnnualReport_TreeData = ""
End Select

End Function

Function Export_AnnualReport_SubstrateCover(xPark As Integer, xYear As Integer)

Dim strWhere As String

Dim strRaw As String

Dim strSubstrate1raw As String
Dim strSubstrate2raw As String

Dim strSubstrate1 As String
Dim strSubstrate2 As String

Dim strSubstrateUnion As String 'final SQL string
'-----------------------------------------------------------------------------

'Create WHERE string ---------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ")"
'-----------------------------------------------------------------------------

' Create data strings --------------------------------------------------------
strRaw = "SELECT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "tlu_Substrate.Substrate_description AS Soil_surface, Count(tbl_Event_Point.Substrate) AS N " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point LEFT JOIN tlu_Substrate ON tbl_Event_Point.Substrate = tlu_Substrate.Substrate_code) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "GROUP BY Year([Start_Date]), tbl_Sites.Unit_Code, tbl_Sites.Site_Name, tbl_Locations.Location_Code, tbl_Locations.Vegetation_Community, tlu_Substrate.Substrate_description, " & _
    "tbl_Locations.Loc_Type, tbl_Locations.Monitoring_Status " & _
    "HAVING (" & strWhere & ")"
'-----------------------------------------------------------------------------

' ALL ------------------------------------------------------------------------
strSubstrate1raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, " & Chr$(34) & "All" & Chr$(34) & " AS Vegetation_Community, qRaw.Soil_surface, " & _
    "(Sum([N])/[NofPoints])*100 AS AbsSubCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRaw.SurveyYear = qTotalPoints.SurveyYear) AND (qRaw.Park = qTotalPoints.Park) " & _
    "AND (qRaw.IslandCode = qTotalPoints.IslandCode) AND (qRaw.SiteCode = qTotalPoints.SiteCode) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, " & Chr$(34) & "All" & Chr$(34) & ", qRaw.Soil_surface, qRaw.N, qTotalPoints.NofPoints"

strSubstrate1 = "SELECT qSub1.SurveyYear, qSub1.Park, qSub1.IslandCode, qSub1.Vegetation_Community, qSub1.Soil_surface, " & _
    "Avg(qSub1.AbsSubCover) AS Average, StDev(qSub1.AbsSubCover) AS StdDev, Min(qSub1.AbsSubCover) AS MinRange, Max(qSub1.AbsSubCover) AS MaxRange " & _
    "FROM (" & strSubstrate1raw & ") AS qSub1 " & _
    "GROUP BY qSub1.SurveyYear, qSub1.Park, qSub1.IslandCode, qSub1.Vegetation_Community, qSub1.Soil_surface"

' Veg ------------------------------------------------------------------------
strSubstrate2raw = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.Soil_surface, (Sum([N])/[NofPoints])*100 AS AbsSubCover " & _
    "FROM (" & strRaw & ") AS qRaw INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRaw.SurveyYear = qTotalPoints.SurveyYear) AND (qRaw.Park = qTotalPoints.Park) " & _
    "AND (qRaw.IslandCode = qTotalPoints.IslandCode) AND (qRaw.SiteCode = qTotalPoints.SiteCode) " & _
    "GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.Soil_surface, qRaw.N, qTotalPoints.NofPoints"

strSubstrate2 = "SELECT qSub2.SurveyYear, qSub2.Park, qRaw.IslandCode, qSub2.Vegetation_Community, qSub2.Soil_surface, " & _
    "Avg(qSub2.AbsSubCover) AS Average, StDev(qSub2.AbsSubCover) AS StdDev, Min(qSub2.AbsSubCover) AS MinRange, Max(qSub2.AbsSubCover) AS MaxRange " & _
    "FROM (" & strSubstrate2raw & ") AS qSub2 " & _
    "GROUP BY qSub2.SurveyYear, qSub2.Park, qSub2.IslandCode, qSub2.Vegetation_Community, qSub2.Soil_surface"
'-----------------------------------------------------------------------------

strSubstrateUnion = "SELECT qry.SurveyYear, " & ParkSelect(xPark) & ", qry.Vegetation_Community, qry.Soil_surface, qry.Average, qry.StdDev, qry.MinRange, qry.MaxRange, " & _
    Chr$(34) & "Annual Report, Substrate" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT * FROM (" & strSubstrate1 & ") AS q1 UNION SELECT * FROM (" & strSubstrate2 & ") AS q2) AS qry"

Export_AnnualReport_SubstrateCover = strSubstrateUnion

End Function

Function Export_AnnualReport_AbsoluteOpenness(xPark As Integer, xYear As Integer)

Dim strWhere As String

Dim strRaw As String
Dim strRawSum As String

Dim strOpenness As String 'final SQL string
'-----------------------------------------------------------------------------

'Create WHERE string ---------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ")"
'-----------------------------------------------------------------------------

' Create data strings --------------------------------------------------------
strRaw = "SELECT DISTINCT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "tbl_Event_Point.Point_No, IIf(IsNull([tbl_Species_Data.ID]), " & Chr$(34) & "Substrate" & Chr$(34) & ", " & Chr$(34) & "Plant" & Chr$(34) & ") AS Openness " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point LEFT JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & strWhere & ")"

strRawSum = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.Openness, Count(qRaw.Openness) AS N " & _
    "FROM (" & strRaw & ") AS qRaw GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.Openness"
'-----------------------------------------------------------------------------

strOpenness = "TRANSFORM Sum(qry.N) AS SumN " & _
    "SELECT qry.SurveyYear, " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.NofPoints, " & _
    Chr$(34) & "Annual Report, Absolute Openness" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT qRawSum.SurveyYear, qRawSum.Park, qRawSum.IslandCode, qRawSum.SiteCode, qRawSum.Vegetation_Community, qRawSum.Openness, qRawSum.N, qTotalPoints.NofPoints " & _
    "FROM (" & strRawSum & ") AS qRawSum INNER JOIN (" & TotalPointsSQL(xPark) & ") AS qTotalPoints ON (qRawSum.SurveyYear = qTotalPoints.SurveyYear) AND (qRawSum.Park = qTotalPoints.Park) " & _
    "AND (qRawSum.IslandCode = qTotalPoints.IslandCode) AND (qRawSum.SiteCode = qTotalPoints.SiteCode)) AS qry " & _
    "GROUP BY qry.SurveyYear, " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.NofPoints " & _
    "PIVOT qry.Openness IN (" & Chr$(34) & "Plant" & Chr$(34) & ", " & Chr$(34) & "Substrate" & Chr$(34) & ")"

Export_AnnualReport_AbsoluteOpenness = strOpenness

End Function

Function Export_AnnualReport_RelativeOpenness(xPark As Integer, xYear As Integer)

Dim strWhere As String

Dim strRaw As String
Dim strRawSum As String

Dim strOpenness As String 'final SQL string
'-----------------------------------------------------------------------------

'Create WHERE string ---------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ")"
'-----------------------------------------------------------------------------

' Create data strings --------------------------------------------------------
strRaw = "SELECT DISTINCT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "tbl_Event_Point.Point_No, IIf(IsNull([tbl_Species_Data.ID]), " & Chr$(34) & "Substrate" & Chr$(34) & ", " & Chr$(34) & "Plant" & Chr$(34) & ") AS Openness " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point LEFT JOIN tbl_Species_Data ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) " & _
    "ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & strWhere & ")"

strRawSum = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.Openness, Count(qRaw.Openness) AS N " & _
    "FROM (" & strRaw & ") AS qRaw GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.Openness"
'-----------------------------------------------------------------------------

strOpenness = "TRANSFORM Sum(qry.N) AS SumN " & _
    "SELECT qry.SurveyYear, " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.TotalNoOfHits, " & _
    Chr$(34) & "Annual Report, Relative Openness" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT qRawSum.SurveyYear, qRawSum.Park, qRawSum.IslandCode, qRawSum.SiteCode, qRawSum.Vegetation_Community, qRawSum.Openness, qRawSum.N, qTotalHits.NofHits AS TotalNoOfHits " & _
    "FROM (" & strRawSum & ") AS qRawSum INNER JOIN (" & TotalHitsSQL(xPark) & ") AS qTotalHits ON (qRawSum.SurveyYear = qTotalHits.SurveyYear) AND (qRawSum.Park = qTotalHits.Park) " & _
    "AND (qRawSum.IslandCode = qTotalHits.IslandCode) AND (qRawSum.SiteCode = qTotalHits.SiteCode)) AS qry " & _
    "GROUP BY qry.SurveyYear, " & ParkSelect(xPark) & ", qry.SiteCode, qry.Vegetation_Community, qry.TotalNoOfHits " & _
    "PIVOT qry.Openness IN (" & Chr$(34) & "Plant" & Chr$(34) & ", " & Chr$(34) & "Substrate" & Chr$(34) & ")"

Export_AnnualReport_RelativeOpenness = strOpenness

End Function

Function Export_AnnualReport_Proportion(xPark As Integer, xYear As Integer)

Dim strWhere As String

Dim strRaw As String
Dim strRawSum As String

Dim strLiveVsDead As String 'final SQL string
'-----------------------------------------------------------------------------

'Create WHERE string ---------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ")"
'-----------------------------------------------------------------------------

' Create data strings --------------------------------------------------------
strRaw = "SELECT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Loc_Type, tbl_Locations.Panel, " & _
    "tbl_Locations.Monitoring_Status, tbl_Locations.Vegetation_Community, tlu_Condition.Analysis_code AS Condition_category " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN (tbl_Event_Point INNER JOIN (tbl_Species_Data LEFT JOIN tlu_Condition " & _
    "ON tbl_Species_Data.Condition = tlu_Condition.Condition) ON tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & strWhere & ")"

strRawSum = "SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Monitoring_Status, qRaw.Loc_Type, qRaw.Panel, qRaw.Vegetation_Community, qRaw.Condition_category, " & _
    "Count(qRaw.Condition_category) AS N " & _
    "FROM (" & strRaw & ") AS qRaw GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Monitoring_Status, qRaw.Loc_Type, qRaw.Panel, qRaw.Vegetation_Community, " & _
    "qRaw.Condition_category"
'----------------------------------------------------------------------------_

strLiveVsDead = "TRANSFORM Sum(qry.N) AS SumN " & _
    "SELECT qry.SurveyYear, " & ParkSelect(xPark) & ", qry.SiteCode, qry.Monitoring_Status, qry.Loc_Type, qry.Panel, qry.Vegetation_Community, qry.TotalNoOfHits, " & _
    Chr$(34) & "Annual Report, Proportion" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT qRawSum.SurveyYear, qRawSum.Park, qRawSum.IslandCode, qRawSum.SiteCode, qRaw.Monitoring_Status, qRaw.Loc_Type, qRaw.Panel, qRawSum.Vegetation_Community, " & _
    "qRawSum.Condition_category, qRawSum.N, qTotalHits.NofHits AS TotalNoOfHits " & _
    "FROM (" & strRawSum & ") AS qRawSum INNER JOIN (" & TotalHitsSQL(xPark) & ") AS qTotalHits ON (qRawSum.SurveyYear = qTotalHits.SurveyYear) AND (qRawSum.Park = qTotalHits.Park) " & _
    "AND (qRawSum.IslandCode = qTotalHits.IslandCode) AND (qRawSum.SiteCode = qTotalHits.SiteCode)) AS qry " & _
    "GROUP BY qry.SurveyYear, " & ParkSelect(xPark) & ", qry.SiteCode, qry.Monitoring_Status, qry.Loc_Type, qry.Panel, qry.Vegetation_Community, qry.TotalNoOfHits " & _
    "PIVOT qry.Condition_category IN (" & Chr$(34) & "Alive" & Chr$(34) & ", " & Chr$(34) & "Dead" & Chr$(34) & ")"

Export_AnnualReport_Proportion = strLiveVsDead

End Function

Function Export_AnnualReport_SurveyDates(xPark As Integer, xYear As Integer, stType As String)

Dim strWhere As String

Dim strSites As String
Dim strDates As String
Dim strDateSurveyed As String 'final SQL string

Dim strNPoints As String
Dim strPointSurveyed As String 'final SQL string
'-----------------------------------------------------------------------------

' Create WHERE string --------------------------------------------------------
strWhere = LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ")"
'--------------------

' Create data strings --------------------------------------------------------
strSites = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Loc_Type, tbl_Locations.Panel, " & _
    "tbl_Locations.Monitoring_Status, tbl_Locations.Vegetation_Community " & _
    "FROM tbl_Sites INNER JOIN tbl_Locations ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & LocTypeFilter(xPark) & ")"
'-----------------------------------------------------------------------------

strDates = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " & _
    "ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & strWhere & ")"

strDateSurveyed = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Monitoring_Status, qry.Loc_Type, qry.Panel, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, " & _
    Chr$(34) & "Annual Report, Survey Dates" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT qrySites.Park, qrySites.IslandCode, qrySites.SiteCode, qrySites.Monitoring_Status, qrySites.Loc_Type, qrySites.Panel, qrySites.Vegetation_Community, " & _
    "qryDates.SurveyYear, qryDates.SurveyDate FROM (" & strSites & ") AS qrySites LEFT JOIN (" & strDates & ") AS qryDates ON (qrySites.Park = qryDates.Park) " & _
    "AND (qrySites.IslandCode = qryDates.IslandCode) AND (qrySites.SiteCode = qryDates.SiteCode) AND (qrySites.Vegetation_Community = qryDates.Vegetation_Community)) AS qry " & _
    "ORDER BY " & ParkSelect(xPark) & ", qrySites.SiteCode"
'-----------------------------------------------------------------------------

strNPoints = "SELECT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, " & _
    "Year([tbl_Events].[Start_Date]) AS SurveyYear, tbl_Events.Start_Date AS SurveyDate, Count(tbl_Event_Point.Point_No) AS NofPoints " & _
    "FROM tbl_Sites INNER JOIN (tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Event_Point ON tbl_Events.Event_ID = tbl_Event_Point.Event_ID) " & _
    "ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID " & _
    "WHERE (" & strWhere & ") " & _
    "GROUP BY tbl_Sites.Unit_Code, tbl_Sites.Site_Name, tbl_Locations.Location_Code, tbl_Locations.Vegetation_Community, Year([tbl_Events].[Start_Date]), tbl_Events.Start_Date"

strPointSurveyed = "SELECT " & ParkSelect(xPark) & ", qry.SiteCode, qry.Monitoring_Status, qry.Loc_Type, qry.Panel, qry.Vegetation_Community, qry.SurveyYear, qry.SurveyDate, qry.NofPoints, " & _
    Chr$(34) & "Annual Report, Points Surveyed" & Chr$(34) & " AS Query_type " & _
    "FROM (SELECT qrySites.Park, qrySites.IslandCode, qrySites.SiteCode, qrySites.Monitoring_Status, qrySites.Loc_Type, qrySites.Panel, qrySites.Vegetation_Community, " & _
    "qryPts.SurveyYear, qryPts.SurveyDate, qryPts.NofPoints FROM (" & strSites & ") AS qrySites LEFT JOIN (" & strNPoints & ") AS qryPts ON (qrySites.Park = qryPts.Park) " & _
    "AND (qrySites.IslandCode = qryPts.IslandCode) AND (qrySites.SiteCode = qryPts.SiteCode) AND (qrySites.Vegetation_Community = qryPts.Vegetation_Community)) AS qry " & _
    "ORDER BY " & ParkSelect(xPark) & ", qry.SiteCode"
'--------------------

' Export dataset based on input parameters -----------------------------------
Select Case stType
    Case Is = "Dates"
        Export_AnnualReport_SurveyDates = strDateSurveyed
    Case Is = "Points"
        Export_AnnualReport_SurveyDates = strPointSurveyed
    Case Else
        MsgBox "Error"
        Export_AnnualReport_SurveyDates = ""
End Select

End Function
