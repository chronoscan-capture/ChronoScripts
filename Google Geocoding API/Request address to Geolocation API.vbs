' ChronoScan HubGEOAPI
' since version 1.0.2.73
' It requires an active Google Geolocation service configured:
'   Either a ChronoScan service account with credits or an external google cloud application with a valid API_KEY 
Dim response
Dim prop
Dim lat
Dim lng
' 1 => Ok, 0 => Fail
Dim requestSuccess

requestSuccess = HubGEOAPI.ProcessAddressInfo("Calle Alcalá, 28080", "")

If requestSuccess <> 1 Then
    ' getting the error
    response = HubGEOAPI.GetLastErrorText()
    msgbox response    
else 
    ' response will hold a full google Geocoding API response json
    ' response =  HubGEOAPI.GetResponseText()
    
    ' Retrieving an especific property from response json
    ' available properties: 
    '   "formatted_address", "country_name", "country_code", "street", "street_number", "city", "postal_code", 
    '   "latitude" or "lat", "longitude" or "lng", "place_id"

    ' example 1: get full address
    prop = HubGEOAPI.GetAddressComponent("formatted_address")
    msgbox "Full address: " & prop

    ' example 2: get coordinates
    lat = HubGEOAPI.GetAddressComponent("lat")
    lng = HubGEOAPI.GetAddressComponent("lng")
    msgbox "Coordinates, latitude " & lat & " - longitude: " & lng

End If

