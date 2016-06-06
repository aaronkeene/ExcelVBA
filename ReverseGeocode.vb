Option Explicit

'Requires Microsoft XML, v6.0 Reference

Private Const googleBaseApiUrl As String = "https://maps.googleapis.com/maps/api"

Enum ResponseProperty
  Result = 1
  Address_Type = 2
  Country = 3
  MainDivision = 4
  SubDivision = 5
  City = 6
  PostalCode = 7
  StreetAddress = 8
  FullAddress = 9
End Enum

Private p_ApiKey        As String
Private p_Lattitude     As Double
Private p_Longitude     As Double
Private p_Result        As DOMDocument60


Public Property Let ApiKey(value As String)
  p_ApiKey = value
End Property

Public Property Get ApiKey() As String
  ApiKey = p_ApiKey
End Property

Public Property Let Lattitude(value As Double)
  p_Lattitude = value
End Property

Public Property Get Lattitude() As Double
  Lattitude = p_Lattitude
End Property

Public Property Let Longitude(value As Double)
  p_Longitude = value
End Property

Public Property Get Longitude() As Double
  Longitude = p_Longitude
End Property


Public Function ResponseProperty(respProperty As ResponseProperty) As String

  Dim xNode         As IXMLDOMNode
  Dim xNodeChild    As IXMLDOMNode
  Dim xAddrNodeList As IXMLDOMNodeList        'Address nodes in first result
  Dim xAddrNode     As IXMLDOMNode
  Dim tmpString     As String

  'Check if the result object is populated; If not get it from Google
  If p_Result Is Nothing Then: GetResponse
  
  'Return for the result property
  If respProperty = Result Then
    Set xNode = p_Result.SelectSingleNode("//status")
    ResponseProperty = xNode.Text
  
  End If
  
  'If status is not OK, no other values are present so exit.
  If Not p_Result.SelectSingleNode("//status").Text = "OK" Then: Exit Function
  
  Select Case respProperty
      
    Case Is = Address_Type
      Set xNode = p_Result.SelectSingleNode("//result/type")
      ResponseProperty = xNode.Text
        
    Case Is = Country
      Set xNode = p_Result.SelectSingleNode("//result")
      Set xAddrNodeList = xNode.SelectNodes("address_component")
      
      For Each xAddrNode In xAddrNodeList
        
        If xAddrNode.SelectSingleNode("type").Text = "country" Then
          ResponseProperty = xAddrNode.SelectSingleNode("short_name").Text
        
        End If
      Next xAddrNode
    
    Case Is = MainDivision
      Set xNode = p_Result.SelectSingleNode("//result")
      Set xAddrNodeList = xNode.SelectNodes("address_component")
      
      For Each xAddrNode In xAddrNodeList
        
        If xAddrNode.SelectSingleNode("type").Text = "administrative_area_level_1" Then
          ResponseProperty = xAddrNode.SelectSingleNode("short_name").Text
        
        End If
      Next xAddrNode
    
    Case Is = SubDivision
      Set xNode = p_Result.SelectSingleNode("//result")
      Set xAddrNodeList = xNode.SelectNodes("address_component")
      
      For Each xAddrNode In xAddrNodeList
        
        If xAddrNode.SelectSingleNode("type").Text = "administrative_area_level_2" Then
          ResponseProperty = xAddrNode.SelectSingleNode("short_name").Text
        
        End If
      Next xAddrNode
      
    Case Is = City
      Set xNode = p_Result.SelectSingleNode("//result")
      Set xAddrNodeList = xNode.SelectNodes("address_component")
      
      For Each xAddrNode In xAddrNodeList
        
        If xAddrNode.SelectSingleNode("type").Text = "sublocality_level_1" Then
          ResponseProperty = xAddrNode.SelectSingleNode("short_name").Text
        
        End If
      Next xAddrNode
      
    Case Is = PostalCode
      Set xNode = p_Result.SelectSingleNode("//result")
      Set xAddrNodeList = xNode.SelectNodes("address_component")
      
      For Each xAddrNode In xAddrNodeList
        
        If xAddrNode.SelectSingleNode("type").Text = "postal_code" Then
          ResponseProperty = xAddrNode.SelectSingleNode("short_name").Text
        
        End If
      Next xAddrNode
      
    Case Is = StreetAddress
      Set xNode = p_Result.SelectSingleNode("//result")
      Set xAddrNodeList = xNode.SelectNodes("address_component")
      
      For Each xAddrNode In xAddrNodeList
        
        If xAddrNode.SelectSingleNode("type").Text = "street_number" Then
          tmpString = xAddrNode.SelectSingleNode("short_name").Text
        
        End If
        
        If xAddrNode.SelectSingleNode("type").Text = "route" Then
          If Not tmpString = vbNullString Then: tmpString = tmpString & " "
          tmpString = tmpString & xAddrNode.SelectSingleNode("short_name").Text
          
        End If
      
      Next xAddrNode
      
      ResponseProperty = tmpString
      
    Case Is = FullAddress
      Set xNode = p_Result.SelectSingleNode("//result/formatted_address")
      ResponseProperty = xNode.Text
      
  End Select

  Set xNode = Nothing

End Function


Private Sub GetResponse()

  Dim xRequest   As XMLHTTP60
  
  Set xRequest = New XMLHTTP60

  With xRequest
    .Open "GET", LatLongApi(Lattitude, Longitude, ApiKey), False
    .send
    
    If Not .responseText = vbNullString Then
      Set p_Result = New DOMDocument60
      p_Result.LoadXML .responseText
  
    End If
  End With
End Sub


Private Function LatLongApi(Lattitude _
                          , Longitude _
                          , Optional ApiKey As String) As String

  Dim url As String
  
  url = googleBaseApiUrl & "/geocode/xml?"
  url = url & "latlng=" & Lattitude & "," & Longitude

  If Not ApiKey = vbNullString Then: url = url & "&key=" & ApiKey
  
  LatLongApi = url
  
End Function
