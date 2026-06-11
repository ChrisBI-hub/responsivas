SELECT 
[Host],
[No_Serie],
[Empresa],
[Area] AS [Sub_area],
[Estado],
[Marca],
[Modelo],
[Procesador],
[RAM],
CONCAT([Capacidad_Discos], ' ', [Tipo_Disco]) AS [Capacidad_Disco],
[Observaciones],
[Codigo_QR]
FROM [BI].[Inventario].[CPU]
