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