<%
' =====================================================
' 点薪官网 · 反馈数据库公共模块（被 feedback.asp / feedback-list.asp 包含）
' =====================================================

' ---------- 配置 ----------
Const FEEDBACK_ADMIN_PASS = "dianxin2026"   ' 后台查看密码，上线前请修改！
Const FEEDBACK_COOLDOWN   = 15              ' 同一访客两次反馈的最小间隔（秒）

Dim gConnStr, gConn, gDbFile

gDbFile = Server.MapPath("db/feedback.mdb")

' 优先使用新驱动 ACE（Access 2007+），回退老驱动 Jet 4.0（Access 2000-2003）
Function BuildConnStr(path)
    Dim s
    s = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";"
    On Error Resume Next
    Dim t : Set t = Server.CreateObject("ADODB.Connection")
    t.ConnectionString = s
    t.Open
    If Err.Number = 0 Then
        t.Close : Set t = Nothing
        BuildConnStr = s
        Exit Function
    End If
    Err.Clear
    Set t = Nothing
    On Error GoTo 0
    BuildConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";"
End Function

Function DbConnect()
    gConnStr = BuildConnStr(gDbFile)
    Set DbConnect = Server.CreateObject("ADODB.Connection")
    DbConnect.Open gConnStr
End Function

' ---------- 数据库与表的自动创建 ----------

' 建库：不存在 feedback.mdb 时用 ADOX 创建一个空的 Access 2000 格式库
Function CreateDatabaseFile()
    On Error Resume Next
    Dim cat
    Set cat = Server.CreateObject("ADOX.Catalog")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        CreateDatabaseFile = False
        Exit Function
    End If
    cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gDbFile & ";"
    If Err.Number <> 0 Then
        Err.Clear
        Set cat = Nothing
        On Error GoTo 0
        CreateDatabaseFile = False
        Exit Function
    End If
    Set cat = Nothing
    On Error GoTo 0
    CreateDatabaseFile = True
End Function

' 建表：库存在但 Feedback 表缺失时执行 DDL（Jet 支持 AUTOINCREMENT）
Function EnsureTable(conn)
    On Error Resume Next
    Dim rs : Set rs = conn.OpenSchema(20)   ' adSchemaTables
    Dim found : found = False
    Do While Not rs.EOF
        If StrComp(rs("TABLE_NAME"), "Feedback", vbTextCompare) = 0 Then
            found = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close : Set rs = Nothing
    If found Then
        On Error GoTo 0
        EnsureTable = True
        Exit Function
    End If
    conn.Execute "CREATE TABLE Feedback (" & _
        "ID AUTOINCREMENT PRIMARY KEY, " & _
        "NickName VARCHAR(30), " & _
        "Contact VARCHAR(60), " & _
        "Category VARCHAR(20), " & _
        "Content MEMO, " & _
        "ClientIP VARCHAR(40), " & _
        "AddTime DATETIME, " & _
        "Status INTEGER NOT NULL DEFAULT 0)"
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        EnsureTable = False
        Exit Function
    End If
    On Error GoTo 0
    EnsureTable = True
End Function

' 连接并保证库与表都存在；失败返回 Nothing
Function OpenFeedbackDb()
    Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(gDbFile) Then
        If Not CreateDatabaseFile() Then
            Set OpenFeedbackDb = Nothing
            Exit Function
        End If
    End If
    Set fso = Nothing

    Dim conn : Set conn = DbConnect()
    If Not EnsureTable(conn) Then
        conn.Close : Set conn = Nothing
        Set OpenFeedbackDb = Nothing
        Exit Function
    End If
    Set OpenFeedbackDb = conn
End Function

' ---------- 安全工具 ----------

Function SafeText(v, maxLen)
    Dim s
    s = Trim("" & v)
    s = Replace(s, "'", "''")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, " ")
    If Len(s) > maxLen Then s = Left(s, maxLen)
    SafeText = s
End Function

Function SafeHtml(v)
    SafeHtml = Server.HTMLEncode(Trim("" & v))
End Function

Function ClientIP()
    Dim ip
    ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    If ip = "" Or IsNull(ip) Then
        ip = Request.ServerVariables("REMOTE_ADDR")
    Else
        ip = Split(ip, ",")(0)
    End If
    ClientIP = SafeText(ip, 40)
End Function
%>
