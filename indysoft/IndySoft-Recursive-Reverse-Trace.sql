-- -------------------------------------------------------------------------
-- To run this, you cannot use the client's ISSQL.exe; instead, you will require access to the DB directly — e.g., SSMS and/or SQL Clients (or IDEs).
-- -------------------------------------------------------------------------

WITH TraceabilityHierarchy AS (
    -- -------------------------------------------------------------------------
    -- 1. THE ANCHOR MEMBER: Find the initial calibration events for the Master Standard
    -- -------------------------------------------------------------------------
    SELECT 
        1 AS TraceLevel, -- Starts the counter at Level 1
        t.TRACE_COMPANY AS Standard_Company,
        t.TRACE_GAGE_SN AS Standard_ID,
        t.SCHED_LAST AS Standard_Cal_Date,
        e.EVENT_NUM,
        e.EVENT_DATE,
        e.EVENT_TYPE,
        e.CUSTOM_FIELD5 AS AS_FOUND,
        e.CUSTOM_FIELD6 AS AS_LEFT,
        e.EVENT_RESULT,
        e.COMPANY AS UUT_Company,     -- UUT = Unit Under Test
        e.GAGE_SN AS UUT_ID
    FROM TRACE2 t
    INNER JOIN EVENTS e ON e.EVENT_NUM = t.EVENT_NUM
    WHERE t.TRACE_COMPANY = 'INSERT STANDARD COMPANY' 
      AND t.TRACE_GAGE_SN = 'INSERT STANDARD ID'
      AND e.EVENT_TYPE LIKE '%CAL%'
      AND CONVERT(DATE, e.EVENT_DATE) >= CONVERT(DATE, 'INSERT START DATE MM/DD/YYYY') 
      AND CONVERT(DATE, e.EVENT_DATE) <= CONVERT(DATE, 'INSERT STOP DATE MM/DD/YYYY')

    UNION ALL

    -- -------------------------------------------------------------------------
    -- 2. THE RECURSIVE MEMBER: Find events where the previous UUT is now the Standard
    -- -------------------------------------------------------------------------
    SELECT 
        th.TraceLevel + 1, -- Increments the level for each loop
        t.TRACE_COMPANY, 
        t.TRACE_GAGE_SN, 
        t.SCHED_LAST,
        e.EVENT_NUM,
        e.EVENT_DATE,
        e.EVENT_TYPE,
        e.CUSTOM_FIELD5,
        e.CUSTOM_FIELD6,
        e.EVENT_RESULT,
        e.COMPANY,
        e.GAGE_SN
    FROM TraceabilityHierarchy th
    -- Join the TRACE2 table to the output of the previous level (th)
    INNER JOIN TRACE2 t ON t.TRACE_COMPANY = th.UUT_Company 
                       AND t.TRACE_GAGE_SN = th.UUT_ID
    INNER JOIN EVENTS e ON e.EVENT_NUM = t.EVENT_NUM
    WHERE e.EVENT_TYPE LIKE '%CAL%'
      -- Ensure chronological flow: The standard must have been calibrated before it was used
      AND CONVERT(DATE, t.SCHED_LAST) > CONVERT(DATE, th.Standard_Cal_Date)
      -- Failsafe: Prevent infinite loops in case of bad data (e.g., Gage A calibrates B, and B calibrates A)
      AND th.TraceLevel < 10 
)

-- -------------------------------------------------------------------------
-- 3. THE FINAL SELECT: Join equipment descriptions to the recursive results
-- -------------------------------------------------------------------------
SELECT 
    th.TraceLevel,
    th.Standard_Company,
    th.Standard_ID,
    g_std.MANUFACTURER AS Standard_MFR,
    g_std.MODEL_NUM AS Standard_MOD_NO,
    th.Standard_Cal_Date,
    
    th.EVENT_NUM,
    th.EVENT_DATE,
    th.EVENT_TYPE,
    th.AS_FOUND,
    th.AS_LEFT,
    th.EVENT_RESULT,
    
    th.UUT_Company,
    th.UUT_ID,
    g_uut.MANUFACTURER AS UUT_MFR,
    g_uut.MODEL_NUM AS UUT_MOD_NO,
    g_uut.GAGE_DESCR AS UUT_DESC
FROM TraceabilityHierarchy th
-- We join to GAGES down here so we don't carry heavy text strings through the recursive loop
LEFT JOIN GAGES g_std ON g_std.COMPANY = th.Standard_Company AND g_std.GAGE_SN = th.Standard_ID
LEFT JOIN GAGES g_uut ON g_uut.COMPANY = th.UUT_Company AND g_uut.GAGE_SN = th.UUT_ID

ORDER BY th.TraceLevel ASC, th.EVENT_DATE DESC;
