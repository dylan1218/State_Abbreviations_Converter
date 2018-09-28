Private Function WhereInArray(arr1 As Variant, vFind As Variant) As Variant
'WhereInArray Function DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check where a value is in an array
Dim i As Long
For i = LBound(arr1) To UBound(arr1)
    If arr1(i) = vFind Then
        WhereInArray = i
        Exit Function
    End If
Next i
'if you get here, vFind was not in the array. Set to null
WhereInArray = Null
End Function


Public Function State_Conversion(state_Identifier As String, abbreviation_IsTrue As Binary) As String

state_Abbreviations = Array("AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY")
state_Full = Array("Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming")

If abbreviation_IsTrue Then
    For Each state_Abbreviation In state_Abbreviations
        If state_Identifier = state_Abreviation Then
            arrayIndex = WhereInArray(state_Abbreviations, state_Identifier)
            State_Conversion = state_Full(arrayIndex)
            Exit For
        Else
            State_Conversion = "#N/A"
        End If
    
    Next
End Function

