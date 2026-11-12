<%@ Language=VBScript CodePage=65001 %>
<% Option Explicit %>
<%
' =====================================================
' 点薪官网 · 用户反馈页
'   独立访问 feedback.asp        → 完整页面
'   iframe 嵌入 feedback.asp?embed=1 → 仅表单（供官网首页引用）
' 数据写入同目录 db/feedback.mdb（Access），库和表首次访问自动创建
' =====================================================
Response.Charset = "utf-8"
%>
<!--#include file="feedback-db.asp"-->
<%
Dim isEmbed
isEmbed = (Request.QueryString("embed") = "1")

Dim posted, done, errMsg
posted = (UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST")
done = False
errMsg = ""

If posted Then
    Dim nick, contact, category, content, conn

    nick     = SafeText(Request.Form("nick"), 30)
    contact  = SafeText(Request.Form("contact"), 60)
    category = SafeText(Request.Form("category"), 20)
    content  = SafeText(Request.Form("content"), 1000)

    If nick = "" Then
        errMsg = "请填写称呼，方便我们回复你。"
    ElseIf Len(Replace(content, " ", "")) < 5 Then
        errMsg = "请把问题或建议描述得稍微详细一点（至少 5 个字）。"
    ElseIf Session("fbLast") <> "" And DateDiff("s", CDate(Session("fbLast")), Now()) < FEEDBACK_COOLDOWN Then
        errMsg = "提交太频繁了，休息 " & FEEDBACK_COOLDOWN & " 秒再试。"
    Else
        Set conn = OpenFeedbackDb()
        If conn Is Nothing Then
            errMsg = "服务器数据库不可用（请确认 db 目录有写入权限，且支持 Access/Jet 驱动）。"
        Else
            On Error Resume Next
            conn.Execute "INSERT INTO Feedback (NickName, Contact, Category, Content, ClientIP, AddTime, Status) VALUES ('" & _
                nick & "', '" & contact & "', '" & category & "', '" & content & "', '" & ClientIP() & "', Now(), 0)", , 129
            If Err.Number <> 0 Then
                errMsg = "写入失败：" & Server.HTMLEncode(Err.Description)
                Err.Clear
            Else
                done = True
                Session("fbLast") = CStr(Now())
            End If
            On Error GoTo 0
            conn.Close
            Set conn = Nothing
        End If
    End If
End If

' 保持表单里已填的内容（失败时不丢）
Dim vNick, vContact, vCategory, vContent
If posted And Not done Then
    vNick = SafeHtml(Request.Form("nick"))
    vContact = SafeHtml(Request.Form("contact"))
    vCategory = SafeHtml(Request.Form("category"))
    vContent = SafeHtml(Request.Form("content"))
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>点薪 · 问题反馈</title>
<style>
    * { box-sizing: border-box; }
    body {
        margin: 0;
        background: #f4efe4;
        color: #3f3a2e;
        font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Noto Sans SC",
            "Microsoft YaHei", "Segoe UI", system-ui, sans-serif;
        font-size: 15px;
        line-height: 1.7;
    }
    .fb { max-width: 640px; margin: 0 auto; padding: 24px 18px 34px; }
    .fb-head h1 { font-size: 21px; margin: 0 0 6px; letter-spacing: 1px; }
    .fb-head p { margin: 0 0 20px; color: #857c66; font-size: 14px; }
    .card {
        background: #fffdf8;
        border: 1px solid rgba(231, 222, 204, 0.8);
        border-radius: 16px;
        box-shadow: 0 8px 24px rgba(120, 95, 40, 0.1);
        padding: 22px 20px;
    }
    label { display: block; font-weight: 600; font-size: 14px; margin: 14px 0 6px; }
    label:first-child { margin-top: 0; }
    label .opt { color: #a59b82; font-weight: 400; font-size: 12px; }
    input[type=text], textarea, select {
        width: 100%;
        padding: 10px 12px;
        border: 1.5px solid rgba(198, 154, 61, 0.35);
        border-radius: 10px;
        background: #fff;
        font-size: 15px;
        color: #3f3a2e;
        font-family: inherit;
        outline: none;
        transition: border-color 0.15s;
    }
    input[type=text]:focus, textarea:focus, select:focus { border-color: #c98d33; }
    textarea { min-height: 130px; resize: vertical; }
    .cats { display: flex; gap: 8px; flex-wrap: wrap; }
    .cats label { margin: 0; }
    .cats input { display: none; }
    .cats span {
        display: inline-block;
        padding: 7px 16px;
        border-radius: 999px;
        border: 1.5px solid rgba(198, 154, 61, 0.4);
        font-size: 13.5px;
        color: #857c66;
        cursor: pointer;
        transition: all 0.15s;
    }
    .cats input:checked + span {
        background: linear-gradient(145deg, #e2b153, #c98d33);
        border-color: transparent;
        color: #fff;
        font-weight: 600;
    }
    .btn {
        display: block;
        width: 100%;
        margin-top: 20px;
        padding: 13px;
        border: none;
        border-radius: 999px;
        background: linear-gradient(145deg, #e2b153, #c98d33);
        color: #fff;
        font-size: 16px;
        font-weight: 600;
        cursor: pointer;
        letter-spacing: 2px;
        box-shadow: 0 6px 16px rgba(170, 120, 40, 0.3);
    }
    .btn:hover { box-shadow: 0 8px 22px rgba(170, 120, 40, 0.45); }
    .tip { color: #a59b82; font-size: 12.5px; margin-top: 12px; text-align: center; }
    .err {
        background: #fdf2ee;
        border: 1px solid #ecc7b3;
        color: #a5452a;
        border-radius: 10px;
        padding: 10px 14px;
        font-size: 14px;
        margin-bottom: 14px;
    }
    .ok {
        background: #eef6ef;
        border: 1px solid #bcdcc3;
        border-radius: 14px;
        padding: 30px 24px;
        text-align: center;
    }
    .ok .ok-ico {
        width: 52px; height: 52px; margin: 0 auto 12px;
        border-radius: 50%; background: #6f9c7d; color: #fff;
        font-size: 26px; line-height: 52px;
    }
    .ok h2 { font-size: 18px; margin: 0 0 8px; }
    .ok p { color: #6d937a; font-size: 14px; margin: 0; }
    @media (max-width: 480px) {
        .fb { padding: 18px 12px 26px; }
        .card { padding: 18px 14px; }
    }
</style>
</head>
<body>
<div class="fb">
    <div class="fb-head">
        <h1>问题反馈与功能建议</h1>
        <p>遇到问题、算薪不对、想加新功能？写下来告诉我们，每一条都会被认真阅读。</p>
    </div>

<% If done Then %>
    <div class="ok">
        <div class="ok-ico">✓</div>
        <h2>反馈已收到，谢谢！</h2>
        <p>我们会尽快查看你的反馈。如果留下了联系方式，问题解决后会尽快回复你。</p>
<% If isEmbed Then %>
        <p style="margin-top:14px"><a href="feedback.asp?embed=1" target="_self" style="color:#c98d33">再提一条 →</a></p>
<% Else %>
        <p style="margin-top:14px"><a href="feedback.asp" style="color:#c98d33">再提一条 →</a></p>
        <p><a href="index.html" style="color:#c98d33">返回官网首页</a></p>
<% End If %>
    </div>
<% Else %>
    <div class="card">
        <% If errMsg <> "" Then %><div class="err"><%= Server.HTMLEncode(errMsg) %></div><% End If %>
        <form method="post" action="feedback.asp<% If isEmbed Then %>?embed=1<% End If %>">
            <label>这是什么类型的反馈？</label>
            <div class="cats">
                <label><input type="radio" name="category" value="问题报告" <% If vCategory = "问题报告" Then %>checked<% End If %> /><span>🐞 问题报告</span></label>
                <label><input type="radio" name="category" value="功能建议" <% If vCategory = "功能建议" Then %>checked<% End If %> /><span>💡 功能建议</span></label>
                <label><input type="radio" name="category" value="算薪咨询" <% If vCategory = "算薪咨询" Then %>checked<% End If %> /><span>¥ 算薪咨询</span></label>
                <label><input type="radio" name="category" value="其他" <% If vCategory <> "问题报告" And vCategory <> "功能建议" And vCategory <> "算薪咨询" Then %>checked<% End If %> /><span>聊聊其他</span></label>
            </div>

            <label>称呼 <span class="opt">（必填，最多 30 字）</span></label>
            <input type="text" name="nick" maxlength="30" placeholder="怎么称呼你？例如：张师傅" value="<%= vNick %>" />

            <label>联系方式 <span class="opt">（选填，QQ / 微信 / 邮箱）</span></label>
            <input type="text" name="contact" maxlength="60" placeholder="方便我们回复你，不填也可以" value="<%= vContact %>" />

            <label>问题描述 <span class="opt">（必填，最多 1000 字）</span></label>
            <textarea name="content" maxlength="1000" placeholder="例如：周六加班 8 小时，月结单里没有按 2 倍计算；或者：希望增加导出 Excel 的功能……"><%= vContent %></textarea>

            <button class="btn" type="submit">提交反馈</button>
            <p class="tip">反馈内容仅用于改进软件 · 提交间隔请大于 <%= FEEDBACK_COOLDOWN %> 秒</p>
        </form>
    </div>
<% End If %>
</div>
</body>
</html>
