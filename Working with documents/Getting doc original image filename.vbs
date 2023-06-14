' getting the original image file name of a document

Dim filePath
filePath = ChronoDocument.get_page_field_value(1, "SrcFile") 

Dim fso, file, fileName
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.GetFile(filePath)
fileName = fso.GetBaseName(file)

' assign to field
UserField_file_name_orig.value = fileName
