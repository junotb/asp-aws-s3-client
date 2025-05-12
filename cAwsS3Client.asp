<%
' S3 클라이언트 클래스
Class cAwsS3Client
    ' S3 상수 선언
    Private S3 ' S3
    Private S3_BUCKET ' 버킷 이름
    Private S3_REGION ' 리전
    Private S3_DOMAIN ' 도메인
    Private S3_ROOT_DIRECTORY ' 루트 디렉토리
    Private S3_ACL ' 접근 제어 목록

    Private S3_USER_NAME ' 사용자 이름
    Private S3_ACCESS_KEY ' 액세스 키
    Private S3_SECRET_KEY ' 비밀 키

    Private S3_HOST ' S3 호스트
    Private CURRENT_UTC_DATETIME ' 현재 시간

    Sub Class_Initialize()
        S3 = "" ' 예시: "s3"
        S3_BUCKET = "" ' 예시: "my-bucket"
        S3_REGION = "" ' 예시: "us-east-1"
        S3_DOMAIN = "" ' 예시: "amazonaws.com"
        S3_ROOT_DIRECTORY = "" ' 예시: "my-directory"
        S3_ACL = "" ' 예시: "public-read"
        
        S3_USER_NAME = "" ' 예시: "my-user"
        S3_ACCESS_KEY = "" ' 예시: "my-access-key"
        S3_SECRET_KEY = "" ' 예시: "my-secret-key"

        S3_HOST = S3_BUCKET & "." & S3 & "." & S3_REGION & "." & S3_DOMAIN
        CURRENT_UTC_DATETIME = GetCustomUTCDateTime()
    End Sub

    Sub Class_Terminate()
        S3 = ""
        S3_BUCKET = ""
        S3_REGION = ""
        S3_DOMAIN = ""
        S3_ROOT_DIRECTORY = ""
        S3_ACL = ""

        S3_USER_NAME = ""
        S3_ACCESS_KEY = ""
        S3_SECRET_KEY = ""

        S3_HOST = ""
        CURRENT_UTC_DATETIME = ""
    End Sub

    '**
    ' S3 객체 컨텐츠 획득
    '**
    Function GetS3ObjectContent(ByVal s3ObjectKey)
        Dim Authorization : Authorization = GetAuthorization(s3ObjectKey)

        With Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
            .Open "GET", "https://" & S3_HOST & s3ObjectKey, False
            .SetOption 2, .GetOption(2)
            .SetRequestHeader "Authorization", Authorization
            .SetRequestHeader "Content-Type", GetContentTypeFromKey(s3ObjectKey)
            .SetRequestHeader "Host", S3_HOST
            .SetRequestHeader "x-amz-content-sha256", "UNSIGNED-PAYLOAD"
            .SetRequestHeader "x-amz-date", CURRENT_UTC_DATETIME
            .Send
            GetS3ObjectContent = .ResponseText
        End With
    End Function

    '**
    ' 컨텐츠 타입 획득
    '**
    Function GetContentTypeFromKey(ByVal key)
        Dim ext : ext = LCase(Right(key, Len(key) - InStrRev(key, ".")))
        
        Select Case ext
            Case "vtt"
                GetContentTypeFromKey = "text/vtt"
            Case "txt"
                GetContentTypeFromKey = "text/plain"
            Case "json"
                GetContentTypeFromKey = "application/json"
            Case "xml"
                GetContentTypeFromKey = "application/xml"
            Case "html", "htm"
                GetContentTypeFromKey = "text/html"
            Case "jpg", "jpeg"
                GetContentTypeFromKey = "image/jpeg"
            Case "png"
                GetContentTypeFromKey = "image/png"
            Case "mp4"
                GetContentTypeFromKey = "video/mp4"
            Case Else
                GetContentTypeFromKey = "application/octet-stream"
        End Select
    End Function

    ' S3 인증 헤더 획득
    Private Function GetAuthorization(ByVal s3ObjectKey)
        Dim Authorization : Authorization = _
            "AWS4-HMAC-SHA256 " & _
            "Credential=" & GetCredential() & ", " & _
            "SignedHeaders=" & GetSignedHeaders() & ", " & _
            "Signature=" & GetSignature(s3ObjectKey)

        GetAuthorization = Authorization  
    End Function

    ' S3 자격 증명 획득
    Private Function GetCredential()
        Dim Credential : Credential = S3_ACCESS_KEY & "/" & Left(CURRENT_UTC_DATETIME, 8) & "/" & S3_REGION & "/" & S3 & "/aws4_request"
        GetCredential = Credential
    End Function

    ' S3 서명 헤더 획득
    Private Function GetSignedHeaders()
        Dim SignedHeaders : SignedHeaders = "host;x-amz-content-sha256;x-amz-date"
        GetSignedHeaders = SignedHeaders
    End Function

    ' S3 서명 획득
    Private Function GetSignature(ByVal s3ObjectKey)
        Dim StringToSign : StringToSign = _
            "AWS4-HMAC-SHA256" & vbLf & _
            CURRENT_UTC_DATETIME & vbLf & _
            Left(CURRENT_UTC_DATETIME, 8) & "/" & S3_REGION & "/" & S3 & "/aws4_request" & vbLf & _
            GetHashedCanonicalRequest(s3ObjectKey)
            
        Dim SigningKey : SigningKey = GetSigningKey()

        GetSignature = BytesToHex(HMACSHA256(SigningKey, StringToSign))
    End Function

    ' S3 서명 키 획득
    Private Function GetSigningKey()
        Dim kSecret, kDate, kRegion, kService, kSigning

        kSecret = "AWS4" & S3_SECRET_KEY
        kDate = HMACSHA256(kSecret, Left(CURRENT_UTC_DATETIME, 8))
        kRegion = HMACSHA256(kDate, S3_REGION)
        kService = HMACSHA256(kRegion, S3)
        kSigning = HMACSHA256(kService, "aws4_request")

        GetSigningKey = kSigning
    End Function

    ' S3 해시된 정규화된 요청 획득
    Private Function GetHashedCanonicalRequest(ByVal s3ObjectKey)
        Dim CanonicalRequest : CanonicalRequest = GetCanonicalRequest(s3ObjectKey)
        GetHashedCanonicalRequest = BytesToHex(SHA256Managed(CanonicalRequest))
    End Function

    ' S3 정규화된 요청 획득
    Private Function GetCanonicalRequest(ByVal s3ObjectKey)
        Dim CanonicalRequest : CanonicalRequest = _
            "GET" & vbLf & _
            s3ObjectKey & vbLf & _
            vbLf & _
            "host:" & LCase(S3_HOST) & vbLf & _
            "x-amz-content-sha256:UNSIGNED-PAYLOAD" & vbLf & _
            "x-amz-date:" & CURRENT_UTC_DATETIME & vbLf & _
            vbLf & _
            GetSignedHeaders() & vbLf & _
            "UNSIGNED-PAYLOAD"
            
        GetCanonicalRequest = CanonicalRequest
    End Function

    ' SHA256 관리 객체 획득
    Private Function SHA256Managed(ByVal varValue)
        With Server.CreateObject("System.Security.Cryptography.SHA256Managed")
            SHA256Managed = .ComputeHash_2(UTF8Bytes(varValue))
        End With
    End Function

    ' HMACSHA256 객체 획득
    Private Function HMACSHA256(ByVal varKey, ByVal varValue)
        With Server.CreateObject("System.Security.Cryptography.HMACSHA256")
            .Key = IIF(IsArray(varKey), varKey, UTF8Bytes(varKey))
            HMACSHA256 = .ComputeHash_2(UTF8Bytes(varValue))
        End With
    End Function

    ' UTF8 바이트 획득
    Private Function UTF8Bytes(ByVal str)
        Dim stream, bytes

        Set stream = Server.CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 2 'Text
        stream.Charset = "utf-8"
        stream.WriteText str
        stream.Position = 0 ' 처음부터 읽기 위해 0으로 설정
        stream.Type = 1 'Binary
        stream.Position = 3 ' Skip BOM (EF BB BF)
        bytes = stream.Read
        stream.Close
        Set stream = Nothing

        UTF8Bytes = bytes
    End Function

    ' 바이트를 16진수로 변환
    Private Function BytesToHex(ByVal bytes)
        Dim i, hexStr
        hexStr = ""

        For i = 1 To LenB(bytes)
            hexStr = hexStr & LCase(Right("0" & Hex(AscB(MidB(bytes, i, 1))), 2))
        Next

        BytesToHex = hexStr
    End Function

    ' 정규화된 URI 인코딩
    Private Function CanonicalUriEncode(ByVal uri)
        Dim parts, i, encodedUri
        parts = Split(uri, "/")
        encodedUri = ""

        For i = 0 To UBound(parts)
            If i > 0 Then encodedUri = encodedUri & "/"
            encodedUri = encodedUri & UriEncode(parts(i))
        Next

        CanonicalUriEncode = encodedUri
    End Function

    ' URI 인코딩
    Private Function UriEncode(ByVal str)
        Dim i, ch, code, encoded
        encoded = ""
        For i = 1 To Len(str)
            ch = Mid(str, i, 1)
            code = Asc(ch)
            If (code >= 48 And code <= 57) Or _
            (code >= 65 And code <= 90) Or _
            (code >= 97 And code <= 122) Or _
            ch = "-" Or ch = "_" Or ch = "." Or ch = "~" Then
                encoded = encoded & ch
            Else
                encoded = encoded & "%" & Right("0" & Hex(code), 2)
            End If
        Next
        UriEncode = encoded
    End Function

    ' 2자리 패딩 함수
    Private Function pad2(n)
        If n < 10 Then
            pad2 = "0" & n
        Else
            pad2 = CStr(n)
        End If
    End Function

    ' 현재 UTC 시간 획득
    Private Function GetCustomUTCDateTime()
        ' 한국 시간 기준으로 시간 조정 (UTC+9)
        Dim utcTime : utcTime = DateAdd("h", -9, Now())
        GetCustomUTCDateTime = _
            Year(utcTime) & _
            pad2(Month(utcTime)) & _
            pad2(Day(utcTime)) & "T" & _
            pad2(Hour(utcTime)) & _
            pad2(Minute(utcTime)) & _
            pad2(Second(utcTime)) & "Z"
    End Function
End Class
%>