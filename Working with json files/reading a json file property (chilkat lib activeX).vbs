
' To use Chilkat_9_5_0.JsonObject
' Download and follow the instructions from https://www.chilkatsoft.com/downloads_ActiveX.asp 

' example with a simple json file that contains: {"details": { "class": "example" }}

Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")

jsonFilePath = "path_to_json_file"

'read content of JSON
Dim fileContent
Set file = fso.OpenTextFile(jsonFilePath)
fileContent = file.ReadAll
file.Close

' Convert the content of the JSON file to a JSON object
Dim jsonObj
Set json = CreateObject("Chilkat_9_5_0.JsonObject")
json.Load(fileContent)

' Get the value of the "class" property inside "details" node
Dim propValue
propValue = json.StringOf("details.class")

' Print value
msgbox  "Value of property 'class': " & propValue