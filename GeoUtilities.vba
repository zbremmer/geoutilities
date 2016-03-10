Option Explicit

Function Geocode(address As String, Optional apiKey) As String
  
  'This function takes an address and optional apiKey as parameters and returns lat/lon in the format: 32.325623,-143.234323
  'https://developers.google.com/maps/documentation/geocoding/intro
  
  Dim query As String
  Dim lat As String
  Dim lon As String
  Dim result As New MSXML2.DOMDocument
  Dim service As New MSXML2.XMLHTTP
  Dim nodes As MSXML2.IXMLDOMNodeList
  Dim node As MSXML2.IXMLDOMNode
  Dim status As String
    
  'Build the URL query
  query = "https://maps.googleapis.com/maps/api/geocode/xml?"
  query = query & "address=" & URLEncode(address)
  query = query & "&sensor=false"
  If Not IsMissing(apiKey) Then query = query & "&key=" & Trim(apiKey)
  
  'Create and send synchronous HTTP request
  service.Open "GET", query, False
  service.send
  
  'Load XML, get lat/long, and handle any errors
  result.LoadXML (service.responseText)
  
  If StrComp(result.SelectSingleNode("GeocodeResponse/status").Text, "OK", vbTextCompare) = 0 Then
    Set nodes = result.getElementsByTagName("result")
    If nodes.Length = 1 Then 'One result returned. Get coordinates.
        Set nodes = result.getElementsByTagName("geometry")
        For Each node In nodes
            lat = node.ChildNodes(0).ChildNodes(0).Text
            lon = node.ChildNodes(0).ChildNodes(1).Text
            Geocode = lat & "," & lon
        Next node
    ElseIf nodes.Length > 1 Then 'multiple results returned
            Geocode = "Multiple results returned. Try using a more specific address."
    End If
  Else 'attribute level error (ZERO_RESULTS, INVALID_REQUEST, OVER_QUERY_LIMIT)
    status = result.SelectSingleNode("GeocodeResponse/status").Text
    If Not result.SelectSingleNode("GeocodeResponse/error_message") Is Nothing Then
        status = status + " " + result.SelectSingleNode("GeocodeResponse/error_message").Text
    End If
    Geocode = status
  End If
  

 
End Function

Public Function Elevation(latitude As String, longitude As String, Optional apiKey, Optional units As String = "metric") As String

' This function takes latitude and longitude along with optional apiKey and units and returns elevation for that location.
' Default units are meters. Results can be in feet if units parameter is 'imperial'.
' https://developers.google.com/maps/documentation/elevation/intro

  Dim query As String
  Dim result As New MSXML2.DOMDocument
  Dim service As New MSXML2.XMLHTTP
  Dim nodes As MSXML2.IXMLDOMNodeList
  Dim node As MSXML2.IXMLDOMNode
  Dim status As String
  Dim elev As String
  
  units = Trim(units)
  
  'Test units value and handle errors
  If StrComp(units, "metric", vbTextCompare) <> 0 And StrComp(units, "imperial", vbTextCompare) <> 0 Then
        Elevation = "Invalid unit value. Enter 'imperial', 'metric', or leave blank."
        GoTo Error1
  End If
  
  'Build the URL query
  query = "https://maps.googleapis.com/maps/api/elevation/xml?"
  query = query & "locations=" & latitude & "," & longitude
  If Not IsMissing(apiKey) Then query = query & "&key=" & Trim(apiKey)
  
  'Create and send synchronous HTTP request
  service.Open "GET", query, False
  service.send
  
  'Load XML, get elevation, and handle errors
  result.LoadXML (service.responseText)
 
  If StrComp(result.SelectSingleNode("ElevationResponse/status").Text, "OK", vbTextCompare) = 0 Then
    Set nodes = result.getElementsByTagName("result")
    If nodes.Length = 1 Then 'One result returned. Get elevation.
        If units = "metric" Then
            Elevation = result.SelectSingleNode("ElevationResponse/result/elevation").Text
        ElseIf units = "imperial" Then
            Elevation = CDbl(result.SelectSingleNode("ElevationResponse/result/elevation").Text) * 3.28084
        End If
    ElseIf nodes.Length > 1 Then 'multiple results returned
            Elevation = "Multiple results returned. Try using one pair of latitude / longitude values."
    End If
  Else 'attribute level error (INVALID_REQUEST, OVER_QUERY_LIMIT, REQUEST_DENIED, UNKNOWN_ERROR)
    status = result.SelectSingleNode("ElevationResponse/status").Text
    If Not result.SelectSingleNode("ElevationResponse/error_message") Is Nothing Then
        status = status + " " + result.SelectSingleNode("ElevationResponse/error_message").Text
    End If
    Elevation = status
  End If
  
Error1:

End Function

Public Function TransitDistAddr(origin As String, destination As String, Optional apiKey As String, Optional mode As String = "driving", Optional units As String = "metric") As String

' This function takes an origin and destination address and returns the distance between the locations.
' The route followed is based on the value for the mode parameter. Default mode is driving. Default units returned are meters.
' https://developers.google.com/maps/documentation/distance-matrix/intro#travel_modes

  Dim query As String
  Dim result As New MSXML2.DOMDocument
  Dim service As New MSXML2.XMLHTTP
  Dim nodes As MSXML2.IXMLDOMNodeList
  Dim node As MSXML2.IXMLDOMNode
  Dim status As String
       
  mode = Trim(mode)
  units = Trim(units)
  
       
  'Check to ensure values are correct
  If StrComp(units, "metric", vbTextCompare) <> 0 And StrComp(units, "imperial", vbTextCompare) <> 0 Then
        TransitDistAddr = "Invalid unit value. Enter 'imperial', 'metric', or leave blank."
        GoTo Error1
  End If
  
  If StrComp(mode, "driving", vbTextCompare) <> 0 And StrComp(mode, "walking", vbTextCompare) <> 0 And StrComp(mode, "bicycling", vbTextCompare) <> 0 And StrComp(mode, "transit", vbTextCompare) <> 0 Then
        TransitDistAddr = "Invalid mode value. Enter 'walking', 'driving', 'bicycling', 'transit', or leave blank."
        GoTo Error1
  End If
    
  'Build the URL query
  query = "https://maps.googleapis.com/maps/api/distancematrix/xml?"
  query = query & "origins=" & URLEncode(origin)
  query = query & "&destinations=" & URLEncode(destination)
  query = query & "&mode=" & mode
  If Not IsMissing(apiKey) Then query = query & "&key=" & Trim(apiKey)

  'Create and send synchronous HTTP request
  service.Open "GET", query, False
  service.send
  
  'Load XML, get results, and handle any errors
  result.LoadXML (service.responseText)
     
  If StrComp(result.SelectSingleNode("DistanceMatrixResponse/status").Text, "OK", vbTextCompare) = 0 Then
    If StrComp(result.SelectSingleNode("DistanceMatrixResponse/row/element/status").Text, "OK", vbTextCompare) = 0 Then
        Set nodes = result.getElementsByTagName("row")
        If nodes.Length = 1 Then 'One result returned. Get distance.
            If units = "metric" Then
                TransitDistAddr = result.SelectSingleNode("DistanceMatrixResponse/row/element/distance/value").Text
            ElseIf units = "imperial" Then
                TransitDistAddr = CDbl(result.SelectSingleNode("DistanceMatrixResponse/row/element/distance/value").Text) * 3.28084
            End If
        ElseIf nodes.Length > 1 Then 'multiple results returned
            TransitDistAddr = "Multiple results returned. Try using one address for the start and one for the destination."
        End If
    Else 'element level error (NOT_FOUND, ZERO_RESULTS)
        TransitDistAddr = result.SelectSingleNode("DistanceMatrixResponse/row/element/status").Text
    End If
  Else 'attribute level error
    status = result.SelectSingleNode("DistanceMatrixResponse/status").Text
    If Not result.SelectSingleNode("DistanceMatrixResponse/status/error_message") Is Nothing Then
        status = status & ": " & result.SelectSingleNode("DistanceMatrixResponse/status/error_message").Text
    End If
    TransitDistAddr = status
  End If
 
Error1:
    
End Function

Public Function TransitDistCoord(startLat As String, startLon As String, endLat As String, endLon As String, Optional apiKey As String, Optional mode As String = "driving", Optional units As String = "metric") As String

' This function takes latitude / longitude values for the origin and destination and returns the distance between the locations.
' The route followed is based on the value for the mode parameter. Default mode is driving. Default units returned are meters.
' https://developers.google.com/maps/documentation/distance-matrix/intro#travel_modes

  Dim query As String
  Dim result As New MSXML2.DOMDocument
  Dim service As New MSXML2.XMLHTTP
  Dim nodes As MSXML2.IXMLDOMNodeList
  Dim node As MSXML2.IXMLDOMNode
  Dim status As String
       
  mode = Trim(mode)
  units = Trim(units)
  
       
  'Check to ensure values are correct
  If StrComp(units, "metric", vbTextCompare) <> 0 And StrComp(units, "imperial", vbTextCompare) <> 0 Then
        TransitDistCoord = "Invalid unit value. Enter 'imperial', 'metric', or leave blank."
        GoTo Error1
  End If
  
  If StrComp(mode, "driving", vbTextCompare) <> 0 And StrComp(mode, "walking", vbTextCompare) <> 0 And StrComp(mode, "bicycling", vbTextCompare) <> 0 And StrComp(mode, "transit", vbTextCompare) <> 0 Then
        TransitDistCoord = "Invalid mode value. Enter 'walking', 'driving', 'bicycling', 'transit', or leave blank."
        GoTo Error1
  End If

  'Build the URL query
  query = "https://maps.googleapis.com/maps/api/distancematrix/xml?"
  query = query & "origins=" & URLEncode(startLat) & "," & URLEncode(startLon)
  query = query & "&destinations=" & URLEncode(endLat) & "," & URLEncode(endLon)
  query = query & "&mode=" & mode
  If Not IsMissing(apiKey) Then query = query & "&key=" & Trim(apiKey)

  'Create and send synchronous HTTP request
  service.Open "GET", query, False
  service.send

  'Load XML, get results, and handle any errors
  result.LoadXML (service.responseText)
     
  If StrComp(result.SelectSingleNode("DistanceMatrixResponse/status").Text, "OK", vbTextCompare) = 0 Then
    If StrComp(result.SelectSingleNode("DistanceMatrixResponse/row/element/status").Text, "OK", vbTextCompare) = 0 Then
        Set nodes = result.getElementsByTagName("row")
        If nodes.Length = 1 Then 'One result returned. Get distance.
            If units = "metric" Then
                TransitDistCoord = result.SelectSingleNode("DistanceMatrixResponse/row/element/distance/value").Text
            ElseIf units = "imperial" Then
                TransitDistCoord = CDbl(result.SelectSingleNode("DistanceMatrixResponse/row/element/distance/value").Text) * 3.28084
            End If
        ElseIf nodes.Length > 1 Then 'multiple results returned
            TransitDistCoord = "Multiple results returned. Try using one address for the start and one for the destination."
        End If
    Else 'element level error (NOT_FOUND, ZERO_RESULTS)
        TransitDistCoord = result.SelectSingleNode("DistanceMatrixResponse/row/element/status").Text
    End If
  Else 'attribute level error
    status = result.SelectSingleNode("DistanceMatrixResponse/status").Text
    If Not result.SelectSingleNode("DistanceMatrixResponse/status/error_message") Is Nothing Then
        status = status & ": " & result.SelectSingleNode("DistanceMatrixResponse/status/error_message").Text
    End If
    TransitDistCoord = status
  End If
 
Error1:
    
End Function
Public Function GeoDistAddr(startAddress As String, endAddress As String, Optional units As String = "metric", Optional apiKey) As String

' This function takes an address for the origin and destination and returns the stright line distance between the locations.
' Default units returned are meters.

units = Trim(units)
       
Dim startCoords As String
Dim endCoords As String
Dim startLat As Double
Dim startLon As Double
Dim endLat As Double
Dim endLon As Double
Dim base As Double

'Check to ensure unit values are correct
If StrComp(units, "metric", vbTextCompare) <> 0 And StrComp(units, "imperial", vbTextCompare) <> 0 Then
    GeoDistAddr = "Invalid unit value. Enter 'imperial', 'metric', or leave blank."
    GoTo Error1
End If

' Get lat/lon coords
If IsMissing(apiKey) Then
    startCoords = Geocode(startAddress)
    endCoords = Geocode(endAddress)
Else
    startCoords = Geocode(startAddress, apiKey)
    endCoords = Geocode(endAddress, apiKey)
End If

'Check to see if lat/lon values were returned and handle any errors
If InStr(startCoords, ",") = 0 Then
    If InStr(endCoords, ",") = 0 Then
        GeoDistAddr = "There was a problem geocoding the starting and ending address."
        GoTo Error1
    Else
        GeoDistAddr = "There was a problem geocoding the starting address."
        GoTo Error1
    End If
ElseIf InStr(endCoords, ",") = 0 Then
        GeoDistAddr = "There was a problem geocoding the ending address."
        GoTo Error1
End If

'Split coords and convert to radians - Haversine requires radians
startLat = (CDbl(Left(startCoords, InStr(startCoords, ",") - 1)) / 180) * 3.14159265359
startLon = (CDbl(Mid(startCoords, InStr(startCoords, ",") + 1, Len(startCoords) - InStr(startCoords, ","))) / 180) * 3.14159265359
endLat = (CDbl(Left(endCoords, InStr(endCoords, ",") - 1)) / 180) * 3.14159265359
endLon = (CDbl(Mid(endCoords, InStr(endCoords, ",") + 1, Len(endCoords) - InStr(endCoords, ","))) / 180) * 3.14159265359

'Calculate distance given unit designation (meters for metric, feet for imperial)
base = Sin(Sqr(Sin((startLat - endLat) / 2) ^ 2 + Cos(startLat) * Cos(endLat) * Sin((startLon - endLon) / 2) ^ 2))
If units = "metric" Then
    GeoDistAddr = (base / Sqr(-base * base + 1)) * 2 * 6371000
ElseIf units = "imperial" Then
    GeoDistAddr = (base / Sqr(-base * base + 1)) * 2 * 20902224.50448
End If

Error1:

End Function

Public Function GeoDistCoord(startLat As String, startLon As String, endLat As String, endLon As String, Optional units As String = "metric") As String

' This function takes the latitude and longitude of the origin and destination and returns the stright line distance between the locations.
' Default units returned are meters.

Dim base As Double

units = Trim(units)
       
'Check to ensure unit values are correct
If StrComp(units, "metric", vbTextCompare) <> 0 And StrComp(units, "imperial", vbTextCompare) <> 0 Then
    GeoDistCoord = "Invalid unit value. Enter 'imperial', 'metric', or leave blank."
    GoTo Error1
End If

'Convert to radians - Haversine requires radians
startLat = (startLat / 180) * 3.14159265359
startLon = (startLon / 180) * 3.14159265359
endLat = (endLat / 180) * 3.14159265359
endLon = (endLon / 180) * 3.14159265359

'Calculate distance given unit designation (meters for metric, feet for imperial)
base = Sin(Sqr(Sin((startLat - endLat) / 2) ^ 2 + Cos(startLat) * Cos(endLat) * Sin((startLon - endLon) / 2) ^ 2))
If units = "metric" Then
    GeoDistCoord = (base / Sqr(-base * base + 1)) * 2 * 6371000
ElseIf units = "imperial" Then
    GeoDistCoord = (base / Sqr(-base * base + 1)) * 2 * 20902224.50448
End If


Error1:

End Function


Public Function CountyByCoord(latitude As String, longitude As String) As String

' This function takes the latitude and longitude value of a location and returns the county name (US only)
' https://www.fcc.gov/general/census-block-conversions-api

Dim query As String
Dim result As New MSXML2.DOMDocument
Dim service As New MSXML2.XMLHTTP
Dim nodes As MSXML2.IXMLDOMNodeList
Dim node As MSXML2.IXMLDOMNode
Dim status As String

latitude = URLEncode(latitude)
longitude = URLEncode(longitude)

'Build the URL query
query = "http://data.fcc.gov/api/block/find?"
query = query & "latitude=" & latitude
query = query & "&longitude=" & longitude
query = query & "&showall=true"

'Create and send synchronous HTTP request
service.Open "GET", query, False
service.send
result.LoadXML (service.responseText)

'Load XML, get results, and handle any errors
If StrComp(result.SelectSingleNode("Response").Attributes.getNamedItem("status").Text, "OK", vbTextCompare) = 0 Then
    CountyByCoord = result.SelectSingleNode("Response/County").Attributes.getNamedItem("name").Text
Else
    CountyByCoord = "There was an error processing this request. Recheck lat/long values."
End If
 
End Function
 
Public Function FIPSByCoord(latitude As String, longitude As String)

' This function takes the latitude and longitude value of a location and returns the county FIPS (US only)
' https://www.fcc.gov/general/census-block-conversions-api

Dim query As String
Dim result As New MSXML2.DOMDocument
Dim service As New MSXML2.XMLHTTP
Dim nodes As MSXML2.IXMLDOMNodeList
Dim node As MSXML2.IXMLDOMNode
Dim status As String

latitude = URLEncode(latitude)
longitude = URLEncode(longitude)

'Build the URL query
query = "http://data.fcc.gov/api/block/find?"
query = query & "latitude=" & latitude
query = query & "&longitude=" & longitude
query = query & "&showall=true"

'Create and send synchronous HTTP request
service.Open "GET", query, False
service.send
result.LoadXML (service.responseText)

'Load XML, get results, and handle any errors
If StrComp(result.SelectSingleNode("Response").Attributes.getNamedItem("status").Text, "OK", vbTextCompare) = 0 Then
    FIPSByCoord = result.SelectSingleNode("Response/County").Attributes.getNamedItem("FIPS").Text
Else
    FIPSByCoord = "There was an error processing this request. Recheck lat/lon values."
End If

End Function

Public Function StateByCoord(latitude As String, longitude As String, Optional abbreviation As Boolean = False)

' This function takes the latitude and longitude value of a location and returns the state name (US only)
' https://www.fcc.gov/general/census-block-conversions-api

Dim query As String
Dim result As New MSXML2.DOMDocument
Dim service As New MSXML2.XMLHTTP
Dim nodes As MSXML2.IXMLDOMNodeList
Dim node As MSXML2.IXMLDOMNode
Dim status As String

latitude = URLEncode(latitude)
longitude = URLEncode(longitude)

'Build the URL query
query = "http://data.fcc.gov/api/block/find?"
query = query & "latitude=" & latitude
query = query & "&longitude=" & longitude
query = query & "&showall=true"

'Create and send synchronous HTTP request
service.Open "GET", query, False
service.send
result.LoadXML (service.responseText)

'Load XML, get results, and handle any errors
If StrComp(result.SelectSingleNode("Response").Attributes.getNamedItem("status").Text, "OK", vbTextCompare) = 0 Then
    If abbreviation Then
        StateByCoord = result.SelectSingleNode("Response/State").Attributes.getNamedItem("code").Text
    Else
        StateByCoord = result.SelectSingleNode("Response/State").Attributes.getNamedItem("name").Text
    End If
Else
    StateByCoord = "There was an error processing this request. Recheck lat/lon values."
End If

End Function

Public Function CountyByAddr(address As String) As String

' Checked and verified. All errors are handled.
' https://www.fcc.gov/general/census-block-conversions-api

Dim query As String
Dim result As New MSXML2.DOMDocument
Dim service As New MSXML2.XMLHTTP
Dim nodes As MSXML2.IXMLDOMNodeList
Dim node As MSXML2.IXMLDOMNode
Dim status As String
Dim latitude As String
Dim longitude As String
Dim geoAddr As String

geoAddr = Geocode(address)

'Check to see if lat/lon values were returned and handle any errors
If InStr(geoAddr, ",") = 0 Then
    CountyByAddr = "There was a problem geocoding the address."
    GoTo Error1
End If

latitude = URLEncode(Left(geoAddr, InStr(geoAddr, ",") - 1))
longitude = URLEncode(Mid(geoAddr, InStr(geoAddr, ",") + 1, Len(geoAddr) - InStr(geoAddr, ",")))

'Build the URL query
query = "http://data.fcc.gov/api/block/find?"
query = query & "latitude=" & latitude
query = query & "&longitude=" & longitude
query = query & "&showall=true"

'Create and send synchronous HTTP request
service.Open "GET", query, False
service.send
result.LoadXML (service.responseText)

'Load XML, get results, and handle any errors
If StrComp(result.SelectSingleNode("Response").Attributes.getNamedItem("status").Text, "OK", vbTextCompare) = 0 Then
    CountyByAddr = result.SelectSingleNode("Response/County").Attributes.getNamedItem("name").Text
Else
    CountyByAddr = "There was an error processing this request. Recheck address."
End If
 
Error1:
 
End Function
 
Public Function FIPSByAddr(address As String)

' Checked and verified. All errors are handled.
' https://www.fcc.gov/general/census-block-conversions-api

Dim query As String
Dim result As New MSXML2.DOMDocument
Dim service As New MSXML2.XMLHTTP
Dim nodes As MSXML2.IXMLDOMNodeList
Dim node As MSXML2.IXMLDOMNode
Dim status As String
Dim latitude As String
Dim longitude As String
Dim geoAddr As String

geoAddr = Geocode(address)

'Check to see if lat/lon values were returned and handle any errors
If InStr(geoAddr, ",") = 0 Then
    FIPSByAddr = "There was a problem geocoding the address."
    GoTo Error1
End If

latitude = URLEncode(Left(geoAddr, InStr(geoAddr, ",") - 1))
longitude = URLEncode(Mid(geoAddr, InStr(geoAddr, ",") + 1, Len(geoAddr) - InStr(geoAddr, ",")))

'Build the URL query
query = "http://data.fcc.gov/api/block/find?"
query = query & "latitude=" & latitude
query = query & "&longitude=" & longitude
query = query & "&showall=true"

'Create and send synchronous HTTP request
service.Open "GET", query, False
service.send
result.LoadXML (service.responseText)

'Load XML, get results, and handle any errors
If StrComp(result.SelectSingleNode("Response").Attributes.getNamedItem("status").Text, "OK", vbTextCompare) = 0 Then
    FIPSByAddr = result.SelectSingleNode("Response/County").Attributes.getNamedItem("FIPS").Text
Else
    FIPSByAddr = "There was an error processing this request. Recheck address."
End If

Error1:

End Function

Public Function StateByAddr(address As String, Optional abbreviation As Boolean = False)
' Checked and verified. All errors are handled.
' https://www.fcc.gov/general/census-block-conversions-api

Dim query As String
Dim result As New MSXML2.DOMDocument
Dim service As New MSXML2.XMLHTTP
Dim nodes As MSXML2.IXMLDOMNodeList
Dim node As MSXML2.IXMLDOMNode
Dim status As String
Dim latitude As String
Dim longitude As String
Dim geoAddr As String

geoAddr = Geocode(address)

'Check to see if lat/lon values were returned and handle any errors
If InStr(geoAddr, ",") = 0 Then
    StateByAddr = "There was a problem geocoding the address."
    GoTo Error1
End If

latitude = URLEncode(Left(geoAddr, InStr(geoAddr, ",") - 1))
longitude = URLEncode(Mid(geoAddr, InStr(geoAddr, ",") + 1, Len(geoAddr) - InStr(geoAddr, ",")))

'Build the URL query
query = "http://data.fcc.gov/api/block/find?"
query = query & "latitude=" & latitude
query = query & "&longitude=" & longitude
query = query & "&showall=true"

'Create and send synchronous HTTP request
service.Open "GET", query, False
service.send
result.LoadXML (service.responseText)

'Load XML, get results, and handle any errors
If StrComp(result.SelectSingleNode("Response").Attributes.getNamedItem("status").Text, "OK", vbTextCompare) = 0 Then
    If abbreviation Then
        StateByAddr = result.SelectSingleNode("Response/State").Attributes.getNamedItem("code").Text
    Else
        StateByAddr = result.SelectSingleNode("Response/State").Attributes.getNamedItem("name").Text
    End If
Else
    StateByAddr = "There was an error processing this request. Recheck address."
End If

Error1:

End Function

Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
  Dim StringLen As Long: StringLen = Len(StringVal)
  StringVal = Trim(StringVal)
  
  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String
    
    If SpaceAsPlus Then Space = "+" Else Space = "%20"
    
    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
        
      Select Case CharCode
      Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
        result(i) = Char
      Case 32
        result(i) = Space
      Case 0 To 15
        result(i) = "%0" & Hex(CharCode)
      Case Else
        result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function
