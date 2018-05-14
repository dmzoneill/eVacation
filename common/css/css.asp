<% 

response.ContentType="text/css"

Dim cssfiles(5) 
cssfiles(0) = "/eVacation/common/css/jquery-ui.min.css"
cssfiles(1) = "/eVacation/common/css/jquery-ui.structure.min.css"
cssfiles(2) = "/eVacation/common/css/jquery-ui.theme.min.css"
cssfiles(3) = "/eVacation/common/css/bootstrap.min.css"
cssfiles(4) = "/eVacation/common/css/evacation.css"
cssfiles(5) = "/eVacation/common/css/colorbox.css"

For Each cssfile In cssfiles

  Dim FSO
  Dim cssfile
  Dim TextStream
  Dim Filepath
  Dim Contents
  
  set FSO = server.createObject("Scripting.FileSystemObject")  
  Filepath = Server.MapPath(cssfile)

  if FSO.FileExists(Filepath) Then      
      Set TextStream = FSO.OpenTextFile(Filepath, 1, False, -2)      
      Contents = TextStream.ReadAll
      Response.write Contents & vbCrLf  & vbCrLf
      TextStream.Close    
  End If

  Set FSO = nothing
  Set TextStream = nothing
  Set Contents = nothing 
  
Next


%>