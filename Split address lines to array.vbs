' Address Split single multiline zone to single lines to map to specific fields i.e. Address1, Address2 etc

fullAdd=UserField_Site_Address.value
addArray=Split(fullAdd,VbCrLf)
' Get number of elements/address lines from array
addLines=UBound(addArray)
'msgbox addLines

Select Case(addLines)

Case 6

UserField_Address1.value=addArray(0)
UserField_Address2.value=addArray(1)
UserField_Address3.value=addArray(2)
UserField_Address4.value=addArray(3)
UserField_Address5.value=addArray(4)
UserField_Address6.value=addArray(5)

Case 7

UserField_Address1.value=addArray(1)
UserField_Address2.value=addArray(2)
UserField_Address3.value=addArray(3)
UserField_Address4.value=addArray(4)
UserField_Address5.value=addArray(5)
UserField_Address6.value=addArray(6)

Case 8

UserField_Address1.value=addArray(2)
UserField_Address2.value=addArray(3)
UserField_Address3.value=addArray(4)
UserField_Address4.value=addArray(5)
UserField_Address5.value=addArray(6)
UserField_Address6.value=addArray(7)

End Select
