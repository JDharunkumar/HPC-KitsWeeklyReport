MODELS_QUERY = """
SELECT 
    pl.Name            AS ProductLineName,
    tpm.Name,
    tpm.Description,
    tpm.CurrentListPrice,
    tpm.stdCost,
    FORMAT(tpm.DateCreated, 'MM/dd/yyyy')  AS DateCreated,
    FORMAT(tpm.DateModified, 'MM/dd/yyyy') AS DateModified,
    tpm.SiteName
FROM tProductModel tpm WITH (NOLOCK)
INNER JOIN tProductLine pl WITH (NOLOCK)
    ON tpm.DBID_ProductLine = pl.DBID
WHERE tpm.Remove <> 1
ORDER BY pl.Name
"""

OPTIONS_QUERY = """
SELECT 
    pl.Name            AS ProductLineName,
    p.Name             AS ProductName,
    o.Name             AS OptionName,
    o.Description,
    o.CurrentListPrice,
    o.stdCost,
    FORMAT(o.DateCreated, 'MM/dd/yyyy')  AS DateCreated,
    FORMAT(o.DateModified, 'MM/dd/yyyy') AS DateModified,
    o.Notes,
    f.CategoryTag
FROM tPLPFO pf WITH (NOLOCK)
INNER JOIN tOption o WITH (NOLOCK)
    ON pf.DBID_ProductLine = o.DBID_ProductLine
   AND pf.DBID_Option      = o.DBID
   AND pf.DBID_PLRev       = o.DBID_PLRev
INNER JOIN tFeature f WITH (NOLOCK)
    ON pf.DBID_Feature     = f.DBID
   AND pf.DBID_ProductLine = f.DBID_ProductLine
   AND pf.DBID_PLRev       = f.DBID_PLRev
INNER JOIN tProductLine pl WITH (NOLOCK)
    ON pf.DBID_ProductLine = pl.DBID
INNER JOIN tProduct p WITH (NOLOCK)
    ON pf.DBID_ProductLine = p.DBID_ProductLine
   AND pf.DBID_Product     = p.DBID
WHERE 
    pf.[Remove] <> 1
    AND (o.KitExpiryDate >= CONVERT(char, GETDATE(),101) OR o.KitExpiryDate IS NULL)
ORDER BY pl.Name, o.Name
"""

SPECIALS_QUERY = """
SELECT 
    pl.Name            AS ProductLineName,
    p.Name             AS ProductName,
    o.Name             AS OptionName,
    o.Description,
    o.CurrentListPrice,
    o.stdCost,
    FORMAT(o.DateCreated, 'MM/dd/yyyy')  AS DateCreated,
    FORMAT(o.DateModified, 'MM/dd/yyyy') AS DateModified,
    o.Notes,
    s.CategoryTag
FROM tPLPSO s WITH (NOLOCK)
INNER JOIN tOption o WITH (NOLOCK)
    ON s.DBID_ProductLine = o.DBID_ProductLine
   AND s.DBID_Option      = o.DBID
   AND s.DBID_PLRev       = o.DBID_PLRev
INNER JOIN tProductLine pl WITH (NOLOCK)
    ON s.DBID_ProductLine = pl.DBID
INNER JOIN tProduct p WITH (NOLOCK)
    ON s.DBID_ProductLine = p.DBID_ProductLine
   AND s.DBID_Product     = p.DBID
WHERE 
    s.[Remove] <> 1
    AND (o.KitExpiryDate >= CONVERT(char, GETDATE(),101) OR o.KitExpiryDate IS NULL)
ORDER BY pl.Name, o.Name
"""
