<%@ Language=VBScript CodePage=65001 %>
<% Option Explicit %>
<%
' =====================================================
' 点薪官网 · 反馈后台（管理用）
'   访问 feedback-list.asp，输入密码登录
'   默认密码在 feedback-db.asp 顶部 FEEDBACK_ADMIN_PASS，上线前务必修改！
' =====================================================
Response.Charset = "utf-8"
%>
<!--#include file="feedback-db.asp"-->
<%
Dim action, isLogin
isLogin = (Session("fbAdmin") = "yes")
action = LCase(Request.QueryString("action") & "")

If action = "logout" Then
    Session.Contents.Remove("fbAdmin")
    isLogin = False
End If

If Not isLogin And UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" And Request.Form("pass") <> "" Then
    If Request.Form("pass") = FEEDBACK_ADMIN_PASS Then
        Session("fbAdmin") = "yes"
        isLogin = True
    Else
        Response.Write "<script>alert('密码不对');history.back();</script>"
        Response.End
    End If
End If

' 登录后的操作：标记已处理 / 待处理 / 删除
If isLogin And action <> "" Then
    Dim id
    id = SafeText(Request.QueryString("id"), 10)
    If id <> "" And IsNumeric(id) Then
        Dim conn : Set conn = OpenFeedbackDb()
        If Not conn Is Nothing Then
            If action = "done" Then
                conn.Execute "UPDATE Feedback SET Status = 1 WHERE ID = " & CLng(id)
            ElseIf action = "undone" Then
                conn.Execute "UPDATE Feedback SET Status = 0 WHERE ID = " & CLng(id)
            ElseIf action = "del" Then
                conn.Execute "DELETE FROM Feedback WHERE ID = " & CLng(id)
            End If
            conn.Close : Set conn = Nothing
        End If
        Response.Redirect "feedback-list.asp"
        Response.End
    End If
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>点薪 · 反馈后台</title>
<style>
    * { box-sizing: border-box; }
    body {
        margin: 0; background: #f4efe4; color: #3f3a2e;
        font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif;
        font-size: 15px; line-height: 1.7;
    }
    .box { max-width: 860px; margin: 30px auto; padding: 0 16px; }
    h1 { font-size: 20px; letter-spacing: 1px; }
    .login { max-width: 380px; margin: 80px auto; background: #fffdf8; border: 1px solid rgba(231,222,204,.8);
             border-radius: 16px; padding: 30px 28px; box-shadow: 0 8px 24px rgba(120,95,40,.12); text-align: center; }
    .login input { width: 100%; padding: 11px 13px; border: 1.5px solid rgba(198,154,61,.4); border-radius: 10px;
                   font-size: 15px; margin: 14px 0; }
    .login button { width: 100%; padding: 12px; border: none; border-radius: 999px; cursor: pointer;
                    background: linear-gradient(145deg, #e2b153, #c98d33); color: #fff; font-size: 15px; font-weight: 600; }
    table { width: 100%; border-collapse: collapse; background: #fffdf8; border-radius: 14px; overflow: hidden;
            box-shadow: 0 6px 20px rgba(120,95,40,.1); }
    th, td { padding: 10px 12px; border-bottom: 1px solid rgba(231,222,204,.7); text-align: left; vertical-align: top; font-size: 14px; }
    th { background: #f0e7d4; font-size: 13px; color: #6d6450; white-space: nowrap; }
    .tag { display: inline-block; padding: 1px 10px; border-radius: 999px; font-size: 12px; white-space: nowrap; }
    .tag-pending { background: #fdeede; color: #a9702a; }
    .tag-done { background: #e7f1e9; color: #4d7c5c; }
    .cat { color: #a9702a; font-size: 12px; white-space: nowrap; }
    .cont { white-space: pre-wrap; word-break: break-all; }
    .ops a { color: #a97f2a; text-decoration: none; font-size: 13px; margin-right: 8px; white-space: nowrap; }
    .ops a:hover { text-decoration: underline; }
    .del { color: #a5452a !important; }
    .head-row { display: flex; align-items: center; gap: 14px; margin-bottom: 14px; }
    .head-row a { color: #a97f2a; text-decoration: none; font-size: 14px; }
    .meta { color: #a59b82; font-size: 12px; }
</style>
</head>
<body>
<% If Not isLogin Then %>
    <div class="login">
        <h1>🔒 反馈后台</h1>
        <p style="color:#a59b82;font-size:13px;margin:6px 0 0">请输入管理密码</p>
        <form method="post" action="feedback-list.asp">
            <input type="password" name="pass" placeholder="管理密码" />
            <button type="submit">进入</button>
        </form>
    </div>
<% Else %>
    <div class="box">
        <div class="head-row">
            <h1 style="margin:0">用户反馈列表</h1>
            <span style="margin-left:auto"></span>
            <a href="feedback-list.asp?action=logout">退出</a>
            <a href="index.html" target="_blank">回官网</a>
        </div>
<%
    Dim rs, cnt
    Set conn = OpenFeedbackDb()
    If conn Is Nothing Then
        Response.Write "<p style='color:#a5452a'>数据库不可用：请确认 db 目录存在且有写入权限，服务器支持 Jet/ACE 驱动。</p>"
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT ID, NickName, Contact, Category, Content, ClientIP, AddTime, Status FROM Feedback ORDER BY ID DESC", conn, 1, 1
        cnt = 0
%>
        <table>
            <tr>
                <th style="width:56px">编号</th>
                <th style="width:110px">称呼</th>
                <th style="width:120px">联系</th>
                <th>内容</th>
                <th style="width:96px">时间 / IP</th>
                <th style="width:120px">操作</th>
            </tr>
<%
        Do While Not rs.EOF And cnt < 300
%>
            <tr>
                <td><%= rs("ID") %></td>
                <td><%= Server.HTMLEncode(rs("NickName") & "") %></td>
                <td><%= Server.HTMLEncode(rs("Contact") & "") %><% If rs("Contact") = "" Then %><span class="meta">未留</span><% End If %></td>
                <td>
                    <span class="cat"><%= Server.HTMLEncode(rs("Category") & "") %></span>
                    <% If CLng(rs("Status")) = 1 Then %><span class="tag tag-done">已处理</span><% Else %><span class="tag tag-pending">待处理</span><% End If %>
                    <div class="cont"><%= Server.HTMLEncode(rs("Content") & "") %></div>
                </td>
                <td>
                    <div class="meta"><%= Server.HTMLEncode(FormatDateTime(rs("AddTime"), 2) & " " & FormatDateTime(rs("AddTime"), 4)) %></div>
                    <div class="meta"><%= Server.HTMLEncode(rs("ClientIP") & "") %></div>
                </td>
                <td class="ops">
                    <% If CLng(rs("Status")) = 1 Then %>
                        <a href="feedback-list.asp?action=undone&id=<%= rs("ID") %>">标为待处理</a>
                    <% Else %>
                        <a href="feedback-list.asp?action=done&id=<%= rs("ID") %>">标为已处理</a>
                    <% End If %>
                    <a class="del" href="feedback-list.asp?action=del&id=<%= rs("ID") %>" onclick="return confirm('确定删除这条反馈？')">删除</a>
                </td>
            </tr>
<%
            cnt = cnt + 1
            rs.MoveNext
        Loop
        If cnt = 0 Then
            Response.Write "<tr><td colspan='6' style='text-align:center;color:#a59b82;padding:30px'>还没有收到反馈</td></tr>"
        End If
        rs.Close : Set rs = Nothing
        conn.Close : Set conn = Nothing
%>
        </table>
<%  End If %>
    </div>
<% End If %>
</body>
</html>
