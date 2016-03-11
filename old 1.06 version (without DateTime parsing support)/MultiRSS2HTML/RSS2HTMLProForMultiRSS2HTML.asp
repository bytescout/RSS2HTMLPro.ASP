<%
 On Error Resume Next
 Response.Expires = -1

 ' =========== RSS2HTMLProForMultiRSS2HTML.ASP For ASP/ASP.NET ==========
 ' copyright 2005-2007 (c) www.Bytescout.com
 '  version 1.06 (11 october 2007)
 ' =========== configuration =====================

 ' ##### Get URLToRSS and Keywords from url #########
 UseParametersFromURL = False

 ' ##### Predefined URL to RSS Feed to display #########
 URLToRSS = ""

 ' ##### Predefined Words For searching #########
 Keywords = ""

 ' ##### Use search in Title, Description and URL #########
 FilterByTitle = False
 FilterByDescription = True
 FilterByURL = False
 ' ##### Description Length #########
 DescriptionLengthLimit = -1
 ' ##### Remove Timezone Info #########
 RemoveTimezoneInfoFromDateTime = True

 ' ##### Max number of displayed items #####
 MaxNumberOfItems = 70

 ' ##### Main template constants
 ' ##### {CHANNELTITLE} will be replaced with item Channel Title
 ' ##### {CHANNELURL} will be replaced with item Channel Url
 MainTemplateHeader = "<table><tr><td><a href=" & """{CHANNELURL}""" & ">{CHANNELTITLE}</a></td></tr>"
 MainTemplateFooter = "</table>"
 ' ##### 

 ' ##### Item template.
 ' ##### {LINK} will be replaced with item link
 ' ##### {TITLE} will be replaced with item title
 ' ##### {DESCRIPTION} will be replaced with item description
 ' ##### {DATE} will be replaced with item date and time
 ' ##### {COMMENTSLINK} will be replaced with link to comments (If you use RSS feed from blog)
 ' ##### {CATEGORY} will be replaced with item category
 ItemTemplate = "<tr><td><strong>{DATE}</strong><br/><strong>{CATEGORY}<br/></strong><a href=" & """{LINK}""" & ">{TITLE}</a><BR>{DESCRIPTION}</td></tr>"

 ' ##### The title of the channel #####
 Dim ChannelTitle
 ' ================================================

 ' ##### Load XmlDOM from URL #########
 Function GetXmlDOM(URLToRSS)
    Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")

    If UseParametersFromURL Then
        If Request.QueryString.Item("URLToRSS") <> Empty Then
            URLToRSS = Request.QueryString.Item("URLToRSS")
        Else
            URLToRSS = ""
        End If
    End If

    ' If URLToRSS is empty return Nothing
    If Len(URLToRSS) <= 0 Then
        Set GetXmlDOM = Nothing
        Set xmlHttp = Nothing

        Exit Function
    End If

    ' Send Request
    xmlHttp.Open "GET", URLToRSS, False
    xmlHttp.Send()
    RSSXML = Trim(xmlHttp.ResponseText)

    Set xmlDOMObj = Server.CreateObject("MSXML2.DomDocument.3.0")
    xmlDOMObj.async = False
    xmlDOMObj.validateOnParse = False
    xmlDOMObj.resolveExternals = False

    If Not xmlDOMObj.LoadXml(RSSXML) Then
        ErrorMessage = "Can Not load XML:" & vbCRLF & xmlDOMObj.parseError.reason & vbCRLF & ErrorMessage
    End If 

    Set xmlHttp = Nothing ' Clear HTTP object

    Set GetXmlDOM = xmlDOMObj
 End Function

 ' ##### Write filtered Items to Response #########
 Function ShowItems(RSSItems)

     ' Set Filter beFore show RSSItems
     rsFilterString = GetFilter

     RSSItemsCount = RSSItems.Length - 1
     j = -1

     For i = 0 To RSSItemsCount
      Set RSSItem = RSSItems.Item(i)

      For Each child in RSSItem.childNodes

       Select Case LCase(child.nodeName)
         Case "title"
               RSStitle = child.text
         Case "link"
	       If RSSLink = "" Then
		    If child.Attributes.length>0 Then
		     RSSLink = child.GetAttribute("href")
		 	     If (RSSLink <> "") Then
 			      If child.GetAttribute("rel") <> "alternate" Then
	   			     RSSLink = ""
			      End If
 	   		     End If
		    End If ' If has attributes	   	
 		    If RSSLink = "" Then
		 	    RSSlink = child.text
	   	    End If
	       End If
         Case "description"
               RSSdescription = child.text
         Case "content" ' atom Format
               RSSdescription = child.text
         Case "published"' atom Format
               RSSDate = child.text
         Case "pubdate"
               RSSDate = child.text
         Case "comments"
               RSSCommentsLink = child.text
         Case "category"
	  	    Set CategoryItems = RSSItem.getElementsByTagName("category")
		    RSSCategory = ""
	  		    For Each categoryitem in CategoryItems
           			If RSSCategory <> "" Then 
					    RSSCategory = RSSCategory & ", "
				    End If

					RSSCategory = RSSCategory & categoryitem.text
			    Next
       End Select
      Next

      j = J + 1

      If J < MaxNumberOfItems Then 
          If CheckItem(RSSTitle, RSSDescription, rsFilterString) Then
              ItemContent = Replace(ItemTemplate, "{LINK}", RSSlink)
              ItemContent = Replace(ItemContent, "{TITLE}", RSSTitle)

              If RemoveTimezoneInfoFromDateTime Then
                index = InStr(RSSDate, "+") - 2
                If index > 1 Then
                  RSSDate = Mid(RSSDate, 1, index)
                End If
                ItemContent = Replace(ItemContent, "{DATE}", RSSDate)
              Else
                ItemContent = Replace(ItemContent, "{DATE}", RSSDate)
              End If

              ItemContent = Replace(ItemContent, "{COMMENTSLINK}", RSSCommentsLink)
              ItemContent = Replace(ItemContent, "{CATEGORY}", RSSCategory)

              If DescriptionLengthLimit = -1 Then
                Response.Write Replace(ItemContent, "{DESCRIPTION}", RSSDescription)
              Else
                Response.Write Replace(ItemContent, "{DESCRIPTION}", Left(RSSDescription, DescriptionLengthLimit))
              End If

              ItemContent = ""
              RSSLink = ""
          End If
      End If

     Next

 End Function

 ' ##### Execute filter string, return False If RSSItem does Not satisfy Keywords #########
 Function CheckItem(RSSTitle, RSSDescription, rsFilterString)
    Execute rsFilterString
 End Function

 ' ##### Build filter string to check RSS Items #########
 Function GetFilter()
    Dim KeywordsString
    Dim FilterString
    Dim KeywordsArr

    If UseParametersFromURL Then
        If Request.QueryString.Item("Keywords") <> Empty Then
            KeywordsString = Server.HTMLEncode(Request.QueryString.Item("Keywords"))
        Else
            KeywordsString = ""
        End If
    Else
        KeywordsString = Keywords
    End If

    ' Search is not case sensitivity
    KeywordsString = LCase(KeywordsString)

    KeywordsString = Replace(KeywordsString, "+", " ")
    KeywordsString = Replace(KeywordsString, "-", " -")

    KeywordsArr = Split(KeywordsString, " ")
    FilterString = "CheckItem = ("

    For Each str in KeywordsArr
        If FilterString <> "CheckItem = (" Then
            FilterString = FilterString & " and ("
        End If

        If Left(str, 1) = "-" Then
            UseAnd = False

            If FilterByTitle Then
                FilterString = FilterString & "(InStr(LCase(RSSTitle), """ & Mid(str, 2) & """) = 0)"
                UseAnd = True
            End If

            If FilterByDescription Then
                If UseAnd Then
                    FilterString = FilterString & " and (InStr(LCase(RSSDescription), """ & Mid(str, 2) & """) = 0)"
                Else
                    FilterString = FilterString & "(InStr(LCase(RSSDescription), """ & Mid(str, 2) & """) = 0)"
                End If

                UseAnd = True
            End If

            If FilterByURL Then
                If UseAnd Then
                    FilterString = FilterString & " and (InStr(LCase(RSSlink), """ & Mid(str, 2) & """) = 0)"
                Else
                    FilterString = FilterString & "(InStr(LCase(RSSlink), """ & Mid(str, 2) & """) = 0)"
                End If
            End If
        Else
            UseOr = False

            If FilterByTitle Then
                FilterString = FilterString & "(InStr(LCase(RSSTitle), """ & str & """) > 0)"
                UseOr = True
            End If

            If FilterByDescription Then
                If UseOr Then
                    FilterString = FilterString & " or (InStr(LCase(RSSDescription), """ & str & """) > 0)"
                Else
                    FilterString = FilterString & "(InStr(LCase(RSSDescription), """ & str & """) > 0)"
                End If
                UseOr = True
            End If

            If FilterByURL Then
                If UseOr Then
                    FilterString = FilterString & " or (InStr(LCase(RSSlink), """ & str & """) > 0)"
                Else
                    FilterString = FilterString & "(InStr(LCase(RSSlink), """ & str & """) > 0)"
                End If
            End If

        End If

        FilterString = FilterString & ") "
    Next

    If FilterString = "CheckItem = (" Then
        FilterString = "CheckItem = True"
    End If

    GetFilter = FilterString
 End Function

 ' ##### Get RSSItems from xmlDOM #########
 Function GetRSS(xmlDOM)

     ' Collect filtered "items" from downloaded RSS
     Set RSSItems = xmlDOM.getElementsByTagName("item")

     'If Not <item>..</item> entries, Then try to get <entry>..</entry>
     If RSSItems.Length <= 0 Then 
        Set RSSItems = xmlDOM.getElementsByTagName("entry") 
        If xmlDOM.childNodes.Length >= 3 Then 
            ChannelTitle = xmlDOM.childNodes.item(3).childNodes.item(0).text
        End If
     Else
        If xmlDOM.childNodes.length >= 1 and _
            xmlDOM.childNodes.item(1).childNodes.length >= 1 Then

            ChannelTitle = xmlDOM.childNodes.item(1).childNodes.item(0).childNodes.item(0).text
        End If
     End If 

     Set GetRSS = RSSItems

 End Function
 
 ' ##### If RSSItems available, write to Response Header, RSSItems and Footer #########
 Function ShowRSS()
                  
    Set xmlDOM = GetXmlDOM(URLToRSS)

    If Not(xmlDOM is Nothing) Then 
        Set RSSItems = GetRSS(xmlDOM)

        Set xmlDOM = Nothing ' clear XML

        If RSSItems.Length > 0 Then 
            Dim TemplateHeader
            TemplateHeader = Replace(MainTemplateHeader, "{CHANNELTITLE}", ChannelTitle)
            TemplateHeader = Replace(TemplateHeader, "{CHANNELURL}", URLToRSS)

            ' writing Header
            Response.Write TemplateHeader

            ' Show RSSItems
            ShowItems RSSItems

            ' writing Footer
            Response.Write MainTemplateFooter

            Set RSSItems = Nothing ' clear RSS
        End If
    End If

 End Function

 ' ##### Entry point to script #########
 ShowRSS

 ' Show Error message, If available
 If Err.Number <> 0 Then
     Response.Write "Error: " & Err.Description
     Err.Clear 
 End If

 ' Response.End ' uncomment this For use in on-the-fly output
%>