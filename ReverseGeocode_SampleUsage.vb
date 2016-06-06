Sub ReverseGeocode_Sample

  Dim gc As GoogleReverseGeoCode
  Dim key As String
  
  Set gc = New GoogleReverseGeoCode
  key = ""                                'Add your Google Map API key
  
  With gc
    .Lattitude = 25
    .Longitude = -45
    .ApiKey = key
        
    MsgBox .ResponseProperty(Result)
    MsgBox .ResponseProperty(Address_Type)
    MsgBox .ResponseProperty(StreetAddress)
    MsgBox .ResponseProperty(City)
    MsgBox .ResponseProperty(SubDivision)
    MsgBox .ResponseProperty(MainDivision)
    MsgBox .ResponseProperty(PostalCode)
    MsgBox .ResponseProperty(Country)
      
  End With
      
  Set gc = Nothing
  
End Sub
