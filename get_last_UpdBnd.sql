select DB_Name() AS Klient, MAX(ChDt) AS HBTransDato
from UpdBnd
Select DB_Name() AS Klient, MAX(ChDt) AS OrdreDato
from Ord
select DB_Name() AS Klient, MAX(ChDt) AS ProdTrDato
from ProdTr