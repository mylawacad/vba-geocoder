Attribute VB_Name = "AddressDecoder"
Private Const CP_UTF8 = 65001
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwflags As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    
Public Function UTF16To8(ByVal UTF16 As String) As String
Dim sBuffer As String
Dim lLength As Long
If UTF16 <> "" Then
    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, 0, 0, 0, 0)
    sBuffer = Space$(lLength)
    lLength = WideCharToMultiByte( _
        CP_UTF8, 0, StrPtr(UTF16), -1, StrPtr(sBuffer), Len(sBuffer), 0, 0)
    sBuffer = StrConv(sBuffer, vbUnicode)
    UTF16To8 = Left$(sBuffer, lLength - 1)
Else
    UTF16To8 = ""
End If
End Function

Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False, _
   Optional UTF8Encode As Boolean = True _
) As String

Dim StringValCopy As String: StringValCopy = _
    IIf(UTF8Encode, UTF16To8(StringVal), StringVal)
Dim StringLen As Long: StringLen = Len(StringValCopy)

If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

  If SpaceAsPlus Then Space = "+" Else Space = "%20"

  For i = 1 To StringLen
    Char = Mid$(StringValCopy, i, 1)
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


Function Get_Formatted_Address(address_string As String) As String
  'Requests link to Microsoft XML, v6.0
  
  Dim sXMLURL As String
  Dim objXMLHTTP As MSXML2.ServerXMLHTTP60
  Dim domResponse As DOMDocument60
  Dim ixnStatus
  Dim arrAddr(7) As String

  'Get the XML From Google
  sXMLURL = "http://maps.googleapis.com/maps/api/geocode/xml?region=us&address=" & URLEncode(address_string) & "&sensor=false"
  Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  With objXMLHTTP
    .Open "GET", sXMLURL, False
    .setRequestHeader "Content-Type", "application/x-www-form-URLEncoded"
    .send
  End With
  'Debug.Print objXMLHTTP.responseText

  'Load XML
  Set domResponse = New DOMDocument60
  domResponse.loadXML objXMLHTTP.responseText
  Set ixnStatus = domResponse.selectSingleNode("//status")
  If ixnStatus Is Nothing Then
    GoTo ExitFunc:
  End If
  
  'MsgBox ixnStatus.Text
  If ixnStatus.Text = "OK" Then
    'Get_Town_From_Adress.Lat = domResponse.selectSingleNode("//result/geometry/location/lat").Text
    'Get_Town_From_Adress.Lng = domResponse.selectSingleNode("//result/geometry/location/lng").Text
    
    For Each attachNode In domResponse.selectNodes("//result//address_component")
        'street_number
        If attachNode.selectSingleNode("type").Text = "street_number" Then
            StreetNum = attachNode.selectSingleNode("long_name").Text
        End If
        
        'street
        If attachNode.selectSingleNode("type").Text = "route" Then
            StreetName = attachNode.selectSingleNode("long_name").Text
        End If
        
        'Town
        If attachNode.selectSingleNode("type").Text = "postal_town" Then
            TownName = attachNode.selectSingleNode("short_name").Text
        End If
        
        'Address 2
        If attachNode.selectSingleNode("type").Text = "sublocality_level_1" Or attachNode.selectSingleNode("type").Text = "sublocality" Then
            AddressB = attachNode.selectSingleNode("long_name").Text
        End If
        
        'Postal code
        If attachNode.selectSingleNode("type").Text = "administrative_area_level_1" Then
            StateName = attachNode.selectSingleNode("long_name").Text
        End If
        
        'Country
        If attachNode.selectSingleNode("type").Text = "country" Then
            CountryName = attachNode.selectSingleNode("long_name").Text
        End If
    Next
    
    Get_Formatted_Address = domResponse.selectSingleNode("//result/formatted_address").Text & " || " & StateName & " || " & CountryName
  Else
    GoTo ExitFunc:
  End If
  
Exit Function
ExitFunc:
Get_Formatted_Address = "Goggle Address Decoding Error. PLease, try again later or contact code developer."
Set domResponse = Nothing
Set objXMLHTTP = Nothing
End Function
