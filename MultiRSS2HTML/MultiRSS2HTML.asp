<!-- #include file="RSS2HTMLProForMultiRSS2HTML.asp" -->
<%  
    On Error Resume Next
    Response.Expires = -1 ' change to another value to use built-in ASP caching (timeout is in minutes), for example set Response.Expires = 20 set to cache the content for 20 minutes


    ' =========== MultiRSS2HTML.ASP for ASP/ASP.NET ==========
    ' copyright 2005-2009 (c) www.Bytescout.com
    '  version 1.18 (22 January 2009)
    ' =========== configuration =====================

    ' ##### Header & Footer for each RSS source #########
    MultiRSSHTMLHeader = "<br/>MULTIPLE RSS FEEDS HEADER (MultiRSSHTMLHeader variable)<br/>"
    MultiRSSHTMLFooter = "<br/>MULTIPLE RSS FEEDS FOOTER (MultiRSSHTMLFooter variable)<br/>"

    ' MultiRSSHTMLHeader = "<table border='1'>"
    ' MultiRSSHTMLFooter = "</table>"

    ' ##### Array RSS Sources #########
    URLs = Array( _
          "http://feeds.feedburner.com/Bytescout", _
          "http://feeds.feedburner.com/MonkeyBites", _
          "http://sixstringbliss.libsyn.com/rss")

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
        
        Response.Write MultiRSSHTMLHeader
        ShowRSS
        Response.Write MultiRSSHTMLFooter 

        If Err.Number <> 0 Then
            Response.Write "Cannot load " & url & "<br />"
            Response.Write "Error: " & Err.Description
            Err.Clear 
        End If

    Next
%>
