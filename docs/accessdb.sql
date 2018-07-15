--- TotalPointsSQL = Function TotalPointsSQL(iPark As Integer)
SELECT DISTINCT tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_ID, 
  tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community, Year(tbl_Events.Start_Date) AS SurveyYear,
  tbl_Events.Start_Date AS SurveyDate, Count(tbl_Event_Point.Point_No) AS NofPoints
FROM 
  tbl_Sites INNER JOIN (
    tbl_Locations INNER JOIN (
      tbl_Events INNER JOIN 
        tbl_Event_Point ON 
        tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON 
      tbl_Locations.Location_ID = tbl_Events.Location_ID) ON 
    tbl_Sites.Site_ID = tbl_Locations.Site_ID
WHERE (" & LocTypeFilter(iPark) & ") 
GROUP BY tbl_Sites.Unit_Code, tbl_Sites.Site_Name, tbl_Locations.Location_ID, tbl_Locations.Location_Code,
  tbl_Locations.Vegetation_Community, tbl_Events.Start_Date, Year(tbl_Events.Start_Date)
HAVING tbl_Sites.Unit_Code = "ParkName(iPark)"

--- strAbsCovData = Calculating Absolute Cover (Figure E2)
SELECT qData.SurveyYear, qData.Park, qData.IslandCode, qData.SiteCode, qData.Vegetation_Community, qData.FxnGroup, 
  qData.Nativity, qData.SumOfN, qTotalPoints.NofPoints, ([SumOfN]/[NofPoints])*100 AS AbsCover 
FROM ("strData") AS qData INNER JOIN ("TotalPointsSQL(xPark)") AS qTotalPoints ON (qData.SurveyYear = qTotalPoints.SurveyYear) AND (qData.Park = qTotalPoints.Park) AND (qData.IslandCode = qTotalPoints.IslandCode) AND (qData.SiteCode = qTotalPoints.SiteCode)

--- strData = strRawSum + str0Data
SELECT qryUnion.SurveyYear, qryUnion.Park, qryUnion.IslandCode, qryUnion.SiteCode, qryUnion.Vegetation_Community, 
  qryUnion.FxnGroup, qryUnion.Nativity, Sum(qryUnion.N) AS SumOfN 
FROM (SELECT * FROM ("str0Data") AS q0Data UNION 
  SELECT * FROM ("strRawSum") AS qryRawSum)  AS qryUnion
  GROUP BY qryUnion.SurveyYear, qryUnion.Park, qryUnion.IslandCode, qryUnion.SiteCode, 
  qryUnion.Vegetation_Community, qryUnion.FxnGroup, qryUnion.Nativity"

--- str0Data = 
SELECT qry1.SurveyYear, qry1.Park, qry1.IslandCode, qry1.SiteCode, qry1.Vegetation_Community, qry2.FxnGroup, qry2.Nativity, 0 AS N 
FROM (" & str1 & ")  AS qry1, (" & str2 & ")  AS qry2

--- str2 =
SELECT DISTINCT Park_Spp.FxnGroup, Park_Spp.Nativity 
FROM ((
  tbl_Sites INNER JOIN (
    tbl_Locations INNER JOIN (
      tbl_Events INNER JOIN (
        tbl_Event_Point LEFT JOIN 
          tbl_Species_Data ON 
          tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON 
        tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON 
      tbl_Locations.Location_ID = tbl_Events.Location_ID) ON 
    tbl_Sites.Site_ID = tbl_Locations.Site_ID ) LEFT JOIN 
    tlu_Condition ON 
    tbl_Species_Data.Condition = tlu_Condition.Condition) LEFT JOIN 
    (" & ParkSpeciesSQL(xPark) & ") AS Park_Spp ON 
    Park_Spp.Species_code = tbl_Species_Data.Species_Code " & _
    "WHERE (" & strWhere & ")"

--- str1 = 
SELECT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community 
FROM 
  tbl_Sites INNER JOIN (
    tbl_Locations INNER JOIN 
      tbl_Events ON 
      tbl_Locations.Location_ID = tbl_Events.Location_ID) ON 
    tbl_Sites.Site_ID = tbl_Locations.Site_ID
WHERE (" & LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & "))"
    

-- strWhere = 
LocTypeFilter(xPark) & " AND ((Year([Start_Date]))=" & xYear & ") AND ((tlu_Condition.Analysis_code) Is Null Or (tlu_Condition.Analysis_code)=" & Chr$(34) & "Alive" & Chr$(34) & ")"

-- LocTypeFilter = 
LocTypeFilter = "((tbl_Sites.Unit_Code)=" & Chr$(34) & ParkName(iPark) & Chr$(34) & ") AND ((tbl_Locations.Loc_Type)=" & Chr$(34) & "I&M" & Chr$(34) & ") " & _
    "AND ((tbl_Locations.Monitoring_Status)=" & Chr$(34) & "Active" & Chr$(34) & ")"

-- VBA: ...strRaw = 
SELECT Year([Start_Date]) AS SurveyYear, tbl_Sites.Unit_Code AS Park, tbl_Sites.Site_Name AS
  IslandCode, tbl_Locations.Location_Code AS SiteCode, tbl_Locations.Vegetation_Community,
  tbl_Species_Data.Species_Code, tlu_Condition.Analysis_code AS Condition, Park_Spp.FxnGroup,
  Park_Spp.Nativity
FROM ((
  tbl_Sites INNER JOIN (
    tbl_Locations INNER JOIN (
      tbl_Events INNER JOIN (
        tbl_Event_Point LEFT JOIN 
          tbl_Species_Data ON 
          tbl_Event_Point.Event_Point_ID = tbl_Species_Data.Event_Point_ID) ON
        tbl_Events.Event_ID = tbl_Event_Point.Event_ID) ON 
      tbl_Locations.Location_ID = tbl_Events.Location_ID) ON 
    tbl_Sites.Site_ID = tbl_Locations.Site_ID) LEFT JOIN
  tlu_Condition ON 
  tbl_Species_Data.Condition = tlu_Condition.Condition) LEFT JOIN 
  (ParkSpeciesSQL(xPark)) AS Park_Spp ON 
  Park_Spp.Species_code = tbl_Species_Data.Species_Code 
  WHERE (" & strWhere & ")"

--- strRawSum =
SELECT qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode, qRaw.Vegetation_Community, qRaw.FxnGroup, qRaw.Nativity, Count(qRaw.Species_Code) AS N
FROM [strRaw] AS qRaw GROUP BY qRaw.SurveyYear, qRaw.Park, qRaw.IslandCode, qRaw.SiteCode,
  qRaw.Vegetation_Community, qRaw.FxnGroup, qRaw.Nativity


-- ParkSpeciesSQL()
SELECT tlu_Project_Taxa.Species_Code, tlu_Project_Taxa.Scientific_name, tlu_Project_Taxa.Layer,
  tlu_Layer.Layer_desc AS FxnGroup, tlu_Project_Taxa.Native, tlu_Nativity.Nativity_desc AS Nativity,
  tlu_Project_Taxa.Perennial, tlu_AnnualPerennial.AnnualPerennial_desc AS AnnPer 
FROM 
  tlu_AnnualPerennial INNER JOIN (
    tlu_Nativity INNER JOIN (
      tlu_Project_Taxa INNER JOIN 
        tlu_Layer ON 
        tlu_Project_Taxa.Layer = tlu_Layer.Layer_code) ON 
      tlu_Nativity.Nativity_code = tlu_Project_Taxa.Native) ON
    tlu_AnnualPerennial.AnnualPerennial_code = tlu_Project_Taxa.Perennial 
WHERE ((tlu_Project_Taxa.Species_code Is Not Null) AND (tlu_Project_Taxa.Unit_code=ParkName(iPark)))