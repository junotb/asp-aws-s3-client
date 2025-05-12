<!--#include virtual="/cAwsS3Client.asp"-->
<%
Dim client : Set client = New cAwsS3Client
client.Class_Initialize()

' 컨텐츠 획득
Dim key : key = "/sample/example.txt"
Dim content : content = client.GetS3ObjectContent(key)

' 컨텐츠 출력
Response.ContentType = client.GetContentTypeFromKey(key)
Response.Write content

client.Class_Terminate()
Set client = Nothing
%>