' it Ouputs a MessageBox vbs variable as a php print_r function for debugging purposes.
' note: it does not support classes
' credit: http://tappetyclick.com/blog/2014/04/17/printr-function-classic-asp-vbscript#.XcGbDFX0myo

' functions

function print_r(data, dumpRef)
  dim dumpData
  if isArray(data) or cbool(instr(TypeName(data),"Dictionary")) or TypeName(data) = "ISessionObject" then
    dumpData = dump(data, 0)
  else
    if TypeName(data) = "Recordset" then
      dumpData = dumpQuery(data)
    else
      dumpData = TypeName(data) & ": " & data
    end if
  end if 
  dumpData = "----------- DUMP CALLED " & dumpRef &" -------------" & vbcrlf & dumpData
  globalDumpData = globalDumpData & vbcrlf & dumpData
  if left(dumpRef,5) = "<pre:" then
    response.write "<pre>"&dumpData&"</pre>"
    if dumpRef = "<pre:stop>" then
      response.end
    end if
  end if
  print_r = dumpData
end function
 
function dumpQuery(recordset)
  dim col, header, data, wrapper, q
  wrapper = "<table border=""1"">"&vbcrlf
  set q = recordset
  q.movefirst
  if q.absoluteposition = 1 then
    header = "  <tr>" & vbcrlf
    for each col in q.fields
      header = header & "    <th align=""left"" valign=""top"">"&col.name&"</th>" & vbcrlf
    next
    header = header & "  </tr>" & vbcrlf
    q.movefirst
  end if
 
  data =   rs.GetString(2, q.recordcount+1, "</td>" & vbcrlf & "    <td valign=""top"">", "[#]", "")
  data = left(data, len(data) - 3)
  data = replace(data, "[#]", "</td>" & vbcrlf & "  </tr>" & vbcrlf & "  <tr>" & vbcrlf & "    <td valign=""top"">")
  data = "  <tr>" & vbcrlf & "    <td valign=""top"">" & data & "</td>" & vbcrlf & "  </tr>"
  wrapper = wrapper & header & data & vbcrlf & "</table>"
  q.movefirst
  dumpQuery = wrapper
end function
 
function dump(data, depth)
  dim output, x
  if isArray(data) then
    output = "Array <br />"
    output = output & Tab(depth) & "(<br />"
    for x=0 to uBound(data)
      output = output & Tab(depth+1) & "["&x&"] => "
      output = output & dump(data(x), depth+2) 
      output = output & "<br />"
    next
    output = output & Tab(depth) & ")"
  elseif cbool(instr(TypeName(data),"Dictionary")) then
    output = TypeName(data) & " <br />"
    output = output & Tab(depth) & "(<br />"
    for each x in data
      output = output & Tab(depth+1) & "["&x&"] => "
      output = output & dump(data(x), depth+2) 
      output = output & "<br />"
    next
    output = output & Tab(depth) & ")"
  elseif TypeName(data) = "ISessionObject" then
    output = TypeName(data) & "<br />(<br/>"& Tab(depth+1) & "Contents<br />"
    output = output & Tab(depth+1) & "(<br />"
    for each x in data.contents
      output = output & Tab(depth+2) & "["&x&"] => "
      output = output & dump(data(x), depth+2) 
      output = output & "<br />"
    next
    output = output & Tab(depth+1) & ")<br/><br/>"
    output = output & Tab(depth+1) & "StaticObjects<br />"
    output = output & Tab(depth+1) & "(<br />"
    for each x in data.StaticObjects
      output = output & Tab(depth+2) & "["&x&"] => "
      output = output & dump(data(x), depth+2) 
      output = output & "<br />"
    next
    output = output & Tab(depth+1) & ")<br/>"
    output = output & Tab(depth) & ")"
  elseif TypeName(data) = "Recordset" then
    output = output & dumpQuery(data)
    output = output & "<br />"
  else
    output = output & data
  end if
  dump = output
end function
 
public function Tab(spaces)
  dim val, x
  val = ""
  for x=1 to spaces
    val = val & "    "
  next
  Tab = val
end function

' Usage example
dim output
output = print_r(array("test", "me", "please"), "[pre:]array")

' Display message.
Response = MsgBox(output)
