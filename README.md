# Access

Hi, this is my place to save Access VB stuff, so I can use it in future projects!

RD2WGS84: converts RD (Rijksdriehoek-) coördinates to WGS84

How to:
* Add Module to Accesss (for example "basCoordinate") with function fRD2WGS84 to convert RD coördinates to Latitude and Longitude.
* In a query add a field (for example "LatLon") in which coordinates are converted to WGS 84 with function fRD2WGS84.
* Formula in query for LatLon. [X] and [Y] are the field in the query which have the X an Y values:
LatLon: IIf(Len([X])<>5;"X no 5 positions";IIf(Len([Y])<>6;"Y no 6 positions";IIf(IsNull([X]);"";IIf(IsNull([Y]);"";fRD2WGS84([X];[Y])))))
