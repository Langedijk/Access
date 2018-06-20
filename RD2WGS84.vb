Option Compare Database
Option Explicit

'Ariën Langedijk: functie fRD2WGS84 rekend RD coördinaten (Rijksdriehoekmeten) om naar DD.dddddd° voor gebruik in Google Maps)
'Met dank aan Octafish (via www.helpmij.nl) voor grotendeels vertalen van de code naar VBA

Function fRD2WGS84(ByVal X As Long, ByVal Y As Long) As String

Dim dX As Double
Dim dY As Double
Dim SomN As Double
Dim SomE As Double
Dim Latitude As String
Dim LatitudeGraden As Integer
Dim LatitudeMinuten As String
Dim Longitude As Double
Dim LongitudeGraden As Integer
Dim LongitudeMinuten As String
Dim Lat As String
Dim Lon As String

dX = (X - 155000) * 10 ^ -5
dY = (Y - 463000) * 10 ^ -5

SomN = (3235.65389 * dY) + (-32.58297 * dX ^ 2) + (-0.2475 * dY ^ 2) + (-0.84978 * dX ^ 2 * dY) + (-0.0655 * dY ^ 3) + (-0.01709 * dX ^ 2 * dY ^ 2) _
+ (-0.00738 * dX) + (0.0053 * dX ^ 4) + (-0.00039 * dX ^ 2 * dY ^ 3) + (0.00033 * dX ^ 4 * dY) + (-0.00012 * dX * dY)
SomE = (5260.52916 * dX) + (105.94684 * dX * dY) + (2.45656 * dX * dY ^ 2) + (-0.81885 * dX ^ 3) + (0.05594 * dX * dY ^ 3) + (-0.05607 * dX ^ 3 * dY) _
+ (0.01199 * dY) + (-0.00256 * dX ^ 3 * dY ^ 2) + (0.00128 * dX * dY ^ 4) + (0.00022 * dY ^ 2) + (-0.00022 * dX ^ 2) + (0.00026 * dX ^ 5)

Latitude = 52.15517 + (SomN / 3600)
Longitude = 5.387206 + (SomE / 3600)

Lat = Replace(Latitude, ",", ".")  'vervang komma door punt (voor gebruik in Google Maps)
Lon = Replace(Longitude, ",", ".") 'vervang komma door punt (voor gebruik in Google Maps)
fRD2WGS84 = Lat & ", " & Lon       'Coördinaten (DD.dddddd° notatie)

'Gebruik onderstaande regel voor coördinaten met een komma
'fRD2WGS84 = Latitude & ", " & Longitude 'graden (DD.dddddd° notatie)'
    
'Gebruik onderstaande ipv bovenstaande regel om het coördinaat weer te geven in graden en minuten (DD°MM.mmm’ notatie)
'LatitudeGraden = Int(Latitude)
'LongitudeGraden = Int(Longitude)
'LatitudeMinuten = (Latitude - LatitudeGraden) * 60
'LongitudeMinuten = (Longitude - LongitudeGraden) * 60
'fRD2WGS84 = LatitudeGraden & "° " & LatitudeMinuten & ", " & LongitudeGraden & "° " & LongitudeMinuten
    
End Function
