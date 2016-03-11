<!-- #include file="RSS2HTMLProForMultiRSS2HTML.asp" -->
<%  
    On Error Resume Next
    Response.Expires = -1

    ' =========== MultiRSS2HTML.ASP for ASP/ASP.NET ==========
    ' copyright 2005-2007 (c) www.Bytescout.com
    '  version 1.06 (11 october 2007)
    ' =========== configuration =====================

    ' ##### Array RSS Sources #########
    URLs = Array( _
             "http://digg.com/rss/indexprogramming.xml", _
             "http://feeds.feedburner.com/MonkeyBites")

    ' ##### Predefined Words for searching #########
    'Keywords = "Windows -Linux"
    'Keywords = "OSI"
    'Keywords = "Intel"
    Keywords = ""

    ' ================================================

    ' ##### Show Items for each RSS #########
    For Each url in URLs
        ' Set rss source
        URLToRSS = url

        ' Show RSS
        ShowRSS

        If Err.Number <> 0 Then
            Response.Write "Cannot load " & url & "<br />"
            Response.Write "Error: " & Err.Description
            Err.Clear 
        End If

    Next
%>