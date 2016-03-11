<%
 ' On Error Resume Next
 Response.Expires = -1 ' change to another value to use built-in ASP caching (timeout is in minutes), for example set Response.Expires = 20 set to cache the content for 20 minutes

 StandaloneScript = True  ' set to TRUE if used as standalone script, set to FALSE when used with MultiRSS2HTML.asp

 ' =========== RSS2HTMLPro.ASP For ASP/ASP.NET ==========
 ' copyright 2005-2010 (c) www.Bytescout.com
 '  version 1.24 (22 June 2010)
 ' =========== configuration =====================

 ' ##### Get "URLToRSS", "Keywords" and "CategoryKeywords" parameters from the original URL parameters #########
 UseParametersFromURL = False

 ' ##### Predefined URL to RSS Feed to display #########
 ' As ASP script relies on MSXML parser so some feeds containing incompatible characters may fail in the script
 ' You should try to use feeburner.com to convert RSS feed into MSXML compatible format
'' For example: URLToRSS = "http://rssnewsapps.ziffdavis.com/tech.xml"

URLToRSS = "http://feeds.feedburner.com/Bytescout"

 ' ##### Predefined Keywords Filtering  #########
 ' Keywords operators: "+" = keywords required; "-" = negative keyword (i.e. keyword should not exist in one of the filtered strings (title, description or URL - see below); "+-" = negative keyword required (i.e. keyword should not exist in any way);
 ' Examples: 
 ' "bytescout-movies" = "bytescout" keyword should exist OR "movies" keyword should NOT exist
 ' "bytescout movies" = "bytescout" keyword should exist OR "movies" keyword should exist
 ' "+bytescout+movies" = "bytescout" keyword should exist AND "movies" keyword should exist
 ' "bytescout+-movies" = "bytescout" keyword should exist AND "movies" keyword should NOT exist
 ' you can also pass keywords using URL using "keywords" parameter (for example: http://mysite.com/RSS2HTMLPro.asp?keywords=time-widget ) but please set UseParametersFromURL = TRUE

Keywords = ""

 ' optional keywords filtering by Categories (if any)
 ' you can also pass CategoryKeywords using URL using "CategoryKeywords" parameter (for example: http://mysite.com/RSS2HTMLPro.asp?CategoryKeywords=tips-business ) but please set UseParametersFromURL = TRUE

 CategoryKeywords = ""


 ' ##### Switch On/Off filtering by Title, Description and URL #########
 FilterByTitle = True
 FilterByDescription = True
 FilterByURL = True
 FilterByCategory = True

 ' ##### Description Length #########
 DescriptionLengthLimit = -1 ' if you have problems with limiting the description and RSS contains rich formatting then set DescriptionStripHTMLTags (see below) to True

 ' ##### Strip rich formatting from description fields ################
 DescriptionStripHTMLTags = False ' change to True if you converts RSS with HTML (images, formatting) in the description

 ' ##### Remove Timezone Info #########
 RemoveTimezoneInfoFromDateTime = True

 ' ##### Max number of displayed items #####
 MaxNumberOfItems = 70

 ' ##### Main template constants
 ' ##### {CHANNELTITLE} will be replaced with item Channel Title
 ' ##### {CHANNELURL} will be replaced with item Channel Url
 ' ##### {CHANNELDESCRIPTION} will be replaced with item Channel description
 MainTemplateHeader = "<table><tr><td><a href="& """{CHANNELURL}"" title=""{CHANNELDESCRIPTION}""" & ">{CHANNELTITLE}</a></td></tr>"
 MainTemplateFooter = "</table>"
 ' ##### 

 ' ##### Item template.
 ' ##### {LINK} will be replaced with item link
 ' ##### {TITLE} will be replaced with item title
 ' ##### {DESCRIPTION} will be replaced with item description
 ' ##### {DATE} will be replaced with item date and time
 ' ##### {COMMENTSLINK} will be replaced with link to comments (If you use RSS feed from blog)
 ' ##### {CATEGORY} will be replaced with item category
 ' ##### PODCASTING
 ' ##### {MEDIA_URL} will be replaced with Media URL
 ' ##### {MEDIA_TITLE} will be replaced with Media title
 ' ##### {MEDIA_SIZE} will be replaced with Media file size in KB
 ' ##### {AUTHOR} will be replaced with Author name
 ' ##### {SOURCE_TITLE} will be replaced with source title 
 ' ##### {SOURCE_URL} will be replaced with source URL link

 ' HTML template for single item in a feed
' ItemTemplate = "<tr><td><strong>{DATE} {TIME}</strong><br/><i>Categories: {CATEGORY}</i><br/><i>Author: {AUTHOR}, via <a href=" & """{SOURCE_URL}""" & ">{SOURCE_TITLE}</a></i><br><a href=" & """{LINK}""" & ">{TITLE}</a><BR>{DESCRIPTION}<br><i>Media download (for podcasts): <a href=" & """{MEDIA_URL}""" & ">{MEDIA_TITLE} ({MEDIA_SIZE} KB)</a></i></td></tr>"
' ItemTemplate = "<tr><td><a href=" & """{SOURCE_URL}""" & ">{SOURCE_TITLE}</a><a href=" & """{LINK}""" & ">{TITLE}</a><BR>{DATE} {AUTHOR}<br>{DESCRIPTION}<br><hr></td></tr>"

ItemTemplate = "<tr><td><a href=" & """{SOURCE_URL}""" & ">{SOURCE_TITLE}</a><a href=" & """{LINK}""" & "target='_blank'>{TITLE}</a><BR>{DATE} {AUTHOR}<br>{DESCRIPTION}<br><hr></td></tr>"

 '#### Date Time format
 '#### dd - Day number
 '#### DD - Full Day Name (for example, Friday)
 '#### mm - Month number (for example, 10 as October)
 '#### MMMM - Full Month Name
 '#### yyyy - Year (number)
 '#### hh - 12 Hour (for example 8 PM)
 '#### H - 24 Hour
 '#### MM - Minute (as number)
 '#### SS - Second (as number)
 '#### Offset - Time offset from UTC
 DateTemplate = "DD, dd, MMMM, yyyy "
 TimeTemplate = "hh:MM:SS AMPM (Offset)" ' for French format please use "HhMM", so it will be displayed like, for example, "15h35"

' ####### variables to control different parameters of date time output #########3
AddLeadingZeroToHours = True    ' set to True to display time "9:35" as "09:35" (add the leading zero to hours when 24-hours format is used)
AddLeadingZeroToMonths = True  ' set to True to display month number "8" as "08" (add the leading zero to month number)
AddLeadingZeroToDays = True     ' set to True to display day number  "7" as "07" (add the leading zero to day number)

 ' ##### channel properties internal variables #####
Dim ChannelTitle, ChannelURL, ChannelDescription

' ###### show empty feeds (when no items in the feed) ###########
ShowEmptyFeeds = False

' ###### User Agent for feed downloads ###########
UserAgent = "Mozilla/5.0 (Windows; U; MSIE 7.0; Windows NT 6.0; en-US)" 

' ###### Connection timeout (ms) ###################
ConnectionTimeout = 30000

' ####### debug mode - disabled by default ############
DebugMode = False

' ######## different private internal variables #########
Dim MonthsNamesSystem
MonthsNamesSystem = Array("","","","","","","","","","","","")
Dim MonthsNamesLatin
MonthsNamesLatin = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

' ########## IMPLEMENTATION #############

 ' Load XmlDOM from URL 
 Function GetXmlDOM(URLToRSS)
    Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")

  xmlHTTP.SetTimeouts ConnectionTimeout, ConnectionTimeout, ConnectionTimeout, ConnectionTimeout

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

 if DebugMode Then Response.Write "<br>xmlHttp.Open GET" End If

    ' Send Request
    xmlHttp.Open "GET", URLToRSS, False

 if DebugMode Then Response.Write "<br>xmlHttp.sendRequestHeader" End If

    xmlhttp.setRequestHeader "User-Agent",UserAgent

 if DebugMode Then Response.Write "<br>xmlHttp.send()" End If

    xmlHttp.Send()

 if DebugMode Then Response.Write "<br>xmlHttp.send() - SUCCESS" End If

    RSSXML = Trim(xmlHttp.ResponseText)

 if DebugMode Then Response.Write "<br>xmlHttp.send() - RSSXML = " & RSSXML End If

    Set xmlDOMObj = Server.CreateObject("MSXML2.DomDocument.3.0")
    xmlDOMObj.async = False
    xmlDOMObj.validateOnParse = False
    xmlDOMObj.resolveExternals = False


    If Not xmlDOMObj.LoadXml(RSSXML) Then
        If DebugMode Then 
		ErrorMessage = "Can Not load XML (see XML content below):" & vbCRLF & xmlDOMObj.parseError.reason & vbCRLF & ErrorMessage & vbCRLF & RSSXML 
 	Else
        		ErrorMessage = "Can Not load Feed from " & URLToRSS
	End If
        Response.Write ErrorMessage
    Else
	 if DebugMode Then Response.Write "<br>xmlDOMObj.LoadXml(RSSXML) - Success" End If
    End If 

    Set xmlHttp = Nothing ' Clear HTTP object

    Set GetXmlDOM = xmlDOMObj
 End Function
 
 ' ##### Write filtered Items to Response #########
 Function ShowItems(RSSItems)

     ' Set filter by categories before outputing RSSItems
     rsCategoriesFilterString = GetCategoriesKeywordsFilter

     ' Set Filter by keywords before outputing RSSItems
     rsKeywordsFilterString = GetKeywordsFilter

     RSSItemsCount = RSSItems.Length - 1
     j = -1


   For i = 0 To RSSItemsCount

''    For i = RSSItemsCount To 0 Step -1 ' uncomment to show items in the reversed order

      Set RSSItem = RSSItems.Item(i)

	' prepare variables to store item information
              ItemContent = ""
              RSSLink = ""
              RSSDate = ""
              RSSAtomDate = ""
              RSSAuthor = ""
              RSSSourceTitle = ""
              RSSSourceURL = ""
              RSSTime = ""
              RSSMediaUrl = ""
              RSSMediaSize = ""
              RSSMediaTitle = ""
              DateTime = Empty

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

         Case "summary"
               RSSdescription = child.text

         Case "published" ' atom Format
               RSSAtomDate = child.text

         Case "pubdate"
               RSSDate = child.text

         Case "comments"
               RSSCommentsLink = child.text


         Case "dc:creator"
               RSSAuthor = child.text


         Case "author" ' atom format

	      For Each tempChild in child.childNodes
	       
	       if RSSAuthor <> "" Then
		Exit For
	       End If
	
	        Select Case LCase(tempChild.nodeName)
	          Case "name"
		 RSSAuthor= tempChild.Text
  	        End Select

	      Next
	if RSSAuthor = "" Then
		RSSAuthor = child.Text
	End If


         Case "source" ' atom format

	      For Each tempChild in child.childNodes
	       
	       if (RSSSourceTitle <> "") and (RSSSourceURL<>"") Then
		Exit For
	       End If
	
	        Select Case LCase(tempChild.nodeName)
	          Case "title"
		 RSSSourceTitle= tempChild.Text
	          Case "link"
		 if RSSSourceURL = "" Then
     		    If tempChild.Attributes.length>0 Then
		      	RSSSourceURL = tempChild.GetAttribute("href")
		    End If ' If has attributes	   	
 		    If RSSSourceURL = "" Then
		 	RSSSourceURL= tempChild.Text
	   	    End If
		 End If

  	        End Select

	      Next

         Case "enclosure"
               RSSMediaUrl = child.attributes.getNamedItem("url").text
	Set tempObj = child.attributes.getNamedItem("length")
 	If Not tempObj is Nothing Then
               		RSSMediaSize = child.attributes.getNamedItem("length").text               
	End If

         Case "itunes:author"
               RSSMediaTitle = child.text

         Case "category"
	  	    Set CategoryItems = RSSItem.getElementsByTagName("category")
		    RSSCategories  = ""
	  		    For Each categoryitem in CategoryItems
           			If RSSCategories  <> "" Then 
					    RSSCategories = RSSCategories & ", "
				    End If

					RSSCategories  = RSSCategories & categoryitem.text
			    Next
       End Select
      Next

      J = J + 1

      If J < MaxNumberOfItems Then 
          If CheckItem(RSSTitle, RSSDescription, RSSCategories, rsKeywordsFilterString, rsCategoriesFilterString) Then
              ItemContent = Replace(ItemTemplate, "{LINK}", RSSLink)
              ItemContent = Replace(ItemContent, "{TITLE}", RSSTitle)

              If RSSAtomDate <> "" Then
	' parse atom date
                  if DebugMode Then Response.Write "DEBUG: parse ATOM date time" End If
	            DateTime = FormatAtomDateTime(RSSAtomDate)
	Else
	' parse RFC822 date (RSS format)
                  if DebugMode Then Response.Write "DEBUG: parse RFC822 date time" End If
	            DateTime = FormatRFC822DateTime(RSSDate)
 	 If IsEmpty(DateTime) Then ' if empty then try to parse as Atom date time
	            DateTime = FormatAtomDateTime(RSSDate)
	 End If		
              End If

              If IsEmpty(DateTime) Then
                  if DebugMode Then RSSTime = "DEBUG: empty date time!" End If
              Else
                  RSSDate = DateTime(0)
                  RSSTime = DateTime(1)
              End If

              If RemoveTimezoneInfoFromDateTime Then
                index = InStr(RSSDate, "&") - 2
                If index > 1 Then
                  RSSDate = Mid(RSSDate, 1, index)
                End If
                ItemContent = Replace(ItemContent, "{DATE}", RSSDate)
              Else
                ItemContent = Replace(ItemContent, "{DATE}", RSSDate)
              End If

              ItemContent = Replace(ItemContent, "{TIME}", RSSTime)

              ItemContent = Replace(ItemContent, "{COMMENTSLINK}", RSSCommentsLink)
              ItemContent = Replace(ItemContent, "{CATEGORY}", RSSCategories)

              ItemContent = Replace(ItemContent, "{AUTHOR}", RSSAuthor)
              ItemContent = Replace(ItemContent, "{SOURCE_TITLE}", RSSSourceTitle)
              ItemContent = Replace(ItemContent, "{SOURCE_URL}", RSSSourceURL)

              ItemContent = Replace(ItemContent, "{MEDIA_URL}", RSSMediaUrl)
              If Len(RSSMediaSize) > 0 Then
                ItemContent = Replace(ItemContent, "{MEDIA_SIZE}", Int(RSSMediaSize / 1024))
              Else
                ItemContent = Replace(ItemContent, "{MEDIA_SIZE}", "")
              End If
              ItemContent = Replace(ItemContent, "{MEDIA_TITLE}", RSSMediaTitle)

	If DescriptionStripHTMLTags Then ' strip HTML formatting from description if needed
		RSSDescription = stripHTML(RSSDescription)
	End If


              If DescriptionLengthLimit = -1 Then
                Response.Write Replace(ItemContent, "{DESCRIPTION}", RSSDescription)
              Else
                Response.Write Replace(ItemContent, "{DESCRIPTION}", Left(RSSDescription, DescriptionLengthLimit) & "...")
              End If

          End If
      End If

     Next

 End Function
 
function GetMonthNumber(sMonth)
 GetMonthNumber = 0 ' default
 sMonth = UCase(sMonth)

 If sMonth = "JAN" Then 
  GetMonthNumber = 1
 ElseIf sMonth = "FEB" Then 
  GetMonthNumber = 2
 ElseIf sMonth = "MAR" Then
  GetMonthNumber = 3
 ElseIf sMonth = "APR" Then
  GetMonthNumber = 4
 ElseIf sMonth = "MAY" Then
  GetMonthNumber = 5
 ElseIf sMonth = "JUN" Then
  GetMonthNumber = 6
 ElseIf sMonth = "JUL" Then
  GetMonthNumber = 7
 ElseIf sMonth = "AUG" Then
  GetMonthNumber = 8
 ElseIf sMonth = "SEP" Then
  GetMonthNumber = 9
 ElseIf sMonth = "OCT" Then
  GetMonthNumber = 10
 ElseIf sMonth = "NOV" Then
  GetMonthNumber = 11
 ElseIf sMonth = "DEC" Then
  GetMonthNumber = 12
 End If
  
End Function


Function GetLatinMonthName (sMonth)

  GetLatinMonthName = sMonth

  If MonthsNamesSystem(0) = "" Then
   ' fill with values for the local month names
    For i = 0 to 11
     MonthsNamesSystem(i) = UCase(MonthName(i+1))
    Next
  End If

    sMonth = UCase(sMonth)

    MonthNumber = -1

    For i = 0 to 11
     if InStr(MonthsNamesSystem(i), sMonth) > 0 Then
	MonthNumber = i
	Exit For
     End If
    Next

    If MonthNumber > -1 Then 
    	GetLatinMonthName = MonthsNamesLatin(MonthNumber)
    End If  
 	
End Function

 Function FormatRFC822DateTime(RSSDate)

    Dim dateRet
    Dim arrRet(1)

    Dim rawDate
    rawDate = Split(RSSDate, " ")
    
    uBoundrawDate = UBound(rawDate) 

    If (uBoundrawDate <4) OR (uBoundrawDate>5) Then
	if DebugMode Then Response.Write "DEBUG: UBound(rawDate)<>5, value=" & CStr(UBound(rawDate)) &"<br>" End If
        Exit Function
    End If

    Dim rawTime
    rawTime = Split(rawDate(4), ":")

    Dim sDay, sMonth, sYear, sOffset
    Dim sHour, sMinute, sSecond
    sDay = rawDate(1)
    sMonth = rawDate(2)
    sYear = rawDate(3)
    
    if UBound(rawDate)<5 Then
     sOffset = ""
    Else     
      sOffset = rawDate(5)
    End If

    sHour = rawTime(0)
    sMinute = rawTime(1)
	If 1<ubound(rawTime) Then
    		sSecond = rawTime(2)
	Else
		sSecond = "00"
	End If

    Dim sDate

   sDate = "#" & sDay & " " & sMonth & " " & sYear & " " & sHour & ":" & sMinute & ":" & sSecond & "#"
   
  if DebugMode Then Response.Write "DEBUG: (1)sDate" & sDate &"<br>" End If   

   If not IsDate(sDate) Then 
 	sMonth = GetLatinMonthName (sMonth)
	sDate = "#" & sDay & " " & sMonth & " " & sYear & " " & sHour & ":" & sMinute & ":" & sSecond & "#"

    if DebugMode Then Response.Write "DEBUG: (2)sDate" & sDate &"<br>" End If	

   End If	
   

    dateRet = CDate(Eval(sDate))

    arrRet(0) = DateTemplate

   ' check if day is 0-9 and add leading zero to output as "04" instead of "4"
   sDay= Day(dateRet)
      If AddLeadingZeroToDays Then 
	If Len(sDay)=1 Then
		sDay = "0" & sDay
	End If
      End If

    arrRet(0) = Replace(arrRet(0), "dd", sDay)

    Dim sWeekdayName
    sWeekdayName = WeekdayName(Weekday(dateRet))

    If FixCapitalLetters Then
        sWeekdayName = UCase(Left(sWeekdayName, 1)) + Right(sWeekdayName, Len(sWeekdayName) - 1)
    End If

    arrRet(0) = Replace(arrRet(0), "DD", sWeekdayName)

   ' check if month is 0-9 and add leading zero to output as "04" instead of "4"
   sMonth= Month(dateRet)
      If AddLeadingZeroToMonths Then 
	If Len(sMonth)=1 Then
		sMonth = "0" & sMonth
	End If
      End If

    arrRet(0) = Replace(arrRet(0), "mm", sMonth)
    arrRet(0) = Replace(arrRet(0), "MMMM", MonthName(Month(dateRet)))
    ' To display short month names (for example Jan instead of January) uncomment the line below
    ' arrRet(0) = Replace(arrRet(0), "MMMM", MonthName(Month(dateRet), True))

    arrRet(0) = Replace(arrRet(0), "yyyy", Year(dateRet))

    arrRet(1) = TimeTemplate
    Dim sHours
    sHours = Hour(dateRet)

    Dim sAMPM
    If Int(sHours) > 12 Then
        sAMPM = "PM"
    Else
        sAMPM = "AM"
    End If

    arrRet(1) = Replace(arrRet(1), "AMPM", sAMPM)

    Dim hh
    hh = Int(sHours)
    If hh > 12 Then hh = hh - 12
    
   ' check if hours is 0-9 and add leading zero to output as "04" instead of "4"
   sHour = sHours
      if AddLeadingZeroToHours Then 
	If Len(sHour)=1 Then
		sHour = "0" & sHour
	End If
      End If


    arrRet(1) = Replace(arrRet(1), "H", sHour)    
    arrRet(1) = Replace(arrRet(1), "hh", hh)

   ' check if minutes is 0-9 and add leading zero to output as "04" instead of "4"
   sMinute = FormatTime(Minute(dateRet))
	If Len(sMinute)=1 Then
		sMinute = "0" & sMinute
	End If
      
    arrRet(1) = Replace(arrRet(1), "MM", sMinute)
    
   ' check if seconds is 0-9 and add leading zero to output as "04" instead of "4"
   sSecond = FormatTime(Second(dateRet))
	If Len(sSecond)=1 Then
		sSecond = "0" & sSecond
	End If

    arrRet(1) = Replace(arrRet(1), "SS", sSecond)

    arrRet(1) = Replace(arrRet(1), "Offset", sOffset)
    
      if DebugMode Then Response.Write "DEBUG: FormatRFC822DateTime = " & arrRet(0) & " ## " & arrRet(1) & "<br>" End If
       
    FormatRFC822DateTime = arrRet
 End Function

 Function FormatAtomDateTime(RSSDate)
    Dim dateRet
    Dim arrRet(1)

	if DebugMode Then Response.Write "DEBUG: enter FormatAtomDateTime" End If

    Dim RawData
    rawData = Split(RSSDate, "T")
    If UBound(rawData) <> 1 Then
	if DebugMode Then Response.Write "DEBUG: UBound(rawData)<>1" End If
        Exit Function
    End If

	if DebugMode Then Response.Write "DEBUG: rawDate" & rawData(0) End If

    Dim rawDate
    rawDate = Split(rawData(0), "-")

    If UBound(rawDate) <> 2 Then
	if DebugMode Then Response.Write "DEBUG: UBound(rawDate) <> 2" End If
        Exit Function
    End If

    Dim rawTime

	if DebugMode Then Response.Write "DEBUG: rawData" & rawData(1) End If

    rawTime = Split(rawData(1), ":")


    Dim sDay, sMonth, sYear, sOffset
    Dim sHour, sMinute, sSecond
    sDay = rawDate(2)
    sMonth = rawDate(1)
    sYear = rawDate(0)
    sOffset = "UTC"

	if DebugMode Then Response.Write "DEBUG: sMonth=" & sMonth & " sDay=" & sDay & " sYear=" & sYear End If

    sHour = rawTime(0)
    sMinute = rawTime(1)
	If 1<ubound(rawTime) Then
    		sSecond = rawTime(2)  
		sSecond = Replace(sSecond, "Z", "")
	Else
		sSecond = "00"
	End If

    Dim sDate

    if Len(sSecond)>2 Then
 	sSecond = "00"
    End If
 

  iMonth = GetMonthNumber(sMonth) 

   If iMonth = 0 Then 

      if DebugMode Then Response.Write "<br>DEBUG: GetLatinName BEFORE , sMonth= " & sMonth & "<br>" End If
	
      sMonth = GetLatinMonthName (sMonth)
      iMonth = GetMonthNumber(sMonth)
 	
      if DebugMode Then Response.Write "<br>DEBUG: GetLatinName call AFTER, sMonth= " & sMonth & "<br>" End If
   	
   End If	
   
    sDate = "#" & sYear & "-" & sMonth & "-" & sDay & " " & sHour & ":" & sMinute & ":" & sSecond & "#"

    dateRet = CDate(Eval(sDate))    
    

    arrRet(0) = DateTemplate

   ' check if day is 0-9 and add leading zero to output as "04" instead of "4"
   sDay= Day(dateRet)
      If AddLeadingZeroToDays Then 
	If Len(sDay)=1 Then
		sDay = "0" & sDay
	End If
      End If

    arrRet(0) = Replace(arrRet(0), "dd", sDay)

    Dim sWeekdayName
    sWeekdayName = WeekdayName(Weekday(dateRet))

    If FixCapitalLetters Then
        sWeekdayName = UCase(Left(sWeekdayName, 1)) + Right(sWeekdayName, Len(sWeekdayName) - 1)
    End If

    arrRet(0) = Replace(arrRet(0), "DD", sWeekdayName)

   ' check if month is 0-9 and add leading zero to output as "04" instead of "4"
   sMonth= Month(dateRet)
      If AddLeadingZeroToMonths Then 
	If Len(sMonth)=1 Then
		sMonth = "0" & sMonth
	End If
      End If

    arrRet(0) = Replace(arrRet(0), "mm", sMonth)

    arrRet(0) = Replace(arrRet(0), "MMMM", MonthName(Month(dateRet)))

    ' To display short month names (for example Jan instead of January) uncomment the line below
    ' arrRet(0) = Replace(arrRet(0), "MMMM", MonthName(Month(dateRet), True))

    arrRet(0) = Replace(arrRet(0), "yyyy", Year(dateRet))

    arrRet(1) = TimeTemplate
    Dim sHours
    sHours = Hour(dateRet)

    Dim sAMPM
    If Int(sHours) > 12 Then
        sAMPM = "PM"
    Else
        sAMPM = "AM"
    End If

    arrRet(1) = Replace(arrRet(1), "AMPM", sAMPM)

    Dim h
    h = Int(sHours)
    If h > 12 Then h = h - 12

   ' check if hours is 0-9 and add leading zero to output as "04" instead of "4"
   sHour = sHours
      if AddLeadingZeroToHours Then 
	If Len(sHour)=1 Then
		sHour = "0" & sHour
	End If
      End If


    arrRet(1) = Replace(arrRet(1), "H", sHour)   
    arrRet(1) = Replace(arrRet(1), "h", h)


   ' check if minutes is 0-9 and add leading zero to output as "04" instead of "4"
   sMinute = FormatTime(Minute(dateRet))
	If Len(sMinute)=1 Then
		sMinute = "0" & sMinute
	End If
      
    arrRet(1) = Replace(arrRet(1), "MM", sMinute)
    
   ' check if seconds is 0-9 and add leading zero to output as "04" instead of "4"
   sSecond = FormatTime(Second(dateRet))
	If Len(sSecond)=1 Then
		sSecond = "0" & sSecond
	End If

    arrRet(1) = Replace(arrRet(1), "SS", sSecond)
    
    arrRet(1) = Replace(arrRet(1), "Offset", sOffset)

	if DebugMode Then Response.Write "DEBUG: " & arrRet(0) End If

    FormatAtomDateTime = arrRet
 End Function

 Function FormatTime(t)
    If t = 0 Then t = "00"
    If Len(t) = 1 Then t = "0" + t

    FormatTime = t
 End Function

 ' ##### Execute filter string, return False If RSSItem does Not satisfy Keywords #########
 Function CheckItem(RSSTitle, RSSDescription, RSSCategories, rsKeywordsFilterString, rsCategoriesFilterString)
  ' check filter by categories
  Execute rsCategoriesFilterString
  ' if passed by categories then run filter by keywords (url, title, description)
  If CheckItem Then 
     Execute rsKeywordsFilterString
  End If
 End Function

 ' ##### Build filter string to check RSS Items #########
 Function GetKeywordsFilter()
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

    KeywordsString = Replace(KeywordsString, "+-", " *") ' means obligatory negative keyword
    KeywordsString = Replace(KeywordsString, "+", " +") ' means obligatory keyword
    KeywordsString = Replace(KeywordsString, "&", " +") ' means obligatory keyword
    KeywordsString = Replace(KeywordsString, "-", " -") ' means non obligatory negative keyword

    KeywordsArr = Split(KeywordsString, " ")
    FilterString = "CheckItem = ("

    For Each str in KeywordsArr
        MinusSign = False
        PlusSign = False
       AsteriskSign = False

        If Left(str, 1) = "-" Then
            MinusSign = True
           str = Mid(str, 2)
        End If

        If Left(str, 1) = "+" Then
            PlusSign = True
           str = Mid(str, 2)
        End If

        If Left(str, 1) = "*" Then
           AsterikSign = True
           str = Mid(str, 2)
        End If


      If str<>"" Then

        If FilterString <> "CheckItem = (" Then
 
         	If PlusSign or AsterikSign Then
            		FilterString = FilterString & " AND ("	
         	Else 
            		FilterString = FilterString & " OR ("	
         	End If

        End If

        If MinusSign or AsterikSign Then
            ConditionsExist = False

            If FilterByTitle Then
                FilterString = FilterString & "(InStr(LCase(RSSTitle), """ & str & """) = 0)"
                ConditionsExist = True
            End If

            If FilterByDescription Then
                If ConditionsExist Then
                    FilterString = FilterString & " AND (InStr(LCase(RSSDescription), """ & str & """) = 0)"
                Else
                    FilterString = FilterString & "(InStr(LCase(RSSDescription), """ & str & """) = 0)"
                End If

                ConditionsExist= True
            End If

            If FilterByURL Then
                If ConditionsExist Then
                    FilterString = FilterString & " AND (InStr(LCase(RSSlink), """ & str & """) = 0)"
                Else
                    FilterString = FilterString & "(InStr(LCase(RSSlink), """ & str & """) = 0)"
                End If
            End If

        Else

            ConditionsExist = False

            If FilterByTitle Then
                FilterString = FilterString & "(InStr(LCase(RSSTitle), """ & str & """) > 0)"
                ConditionsExist = True
            End If

            If FilterByDescription Then
                If ConditionsExist Then
                    FilterString = FilterString & " OR (InStr(LCase(RSSDescription), """ & str & """) > 0)"
                Else
                    FilterString = FilterString & "(InStr(LCase(RSSDescription), """ & str & """) > 0)"
                End If
                ConditionsExist = True
            End If

            If FilterByURL Then
                If ConditionsExist Then
                    FilterString = FilterString & " OR (InStr(LCase(RSSlink), """ & str & """) > 0)"
                Else
                    FilterString = FilterString & "(InStr(LCase(RSSlink), """ & str & """) > 0)"
                End If
            End If

        End If

        FilterString = FilterString & ") "
     End If ' If str<>"" condition
    Next

    If FilterString = "CheckItem = (" Then
        FilterString = "CheckItem = True"
    End If

    GetKeywordsFilter = FilterString
 End Function

 ' ##### Build filter string to check RSS Items by Categories #########
 Function GetCategoriesKeywordsFilter()
    Dim CategoriesKeywordsString
    Dim FilterString
    Dim CategoriesKeywordsArr

    If UseParametersFromURL Then
        If Request.QueryString.Item("CategoryKeywords") <> Empty Then
            CategoriesKeywordsString = Server.HTMLEncode(Request.QueryString.Item("CategoryKeywords"))
        Else
            CategoriesKeywordsString = ""
        End If
    Else
        CategoriesKeywordsString = CategoryKeywords
    End If

    ' Search is not case sensitivity
    CategoriesKeywordsString = LCase(CategoriesKeywordsString)

    CategoriesKeywordsString = Replace(CategoriesKeywordsString, "+-", " *")
    CategoriesKeywordsString = Replace(CategoriesKeywordsString, "+", " +")
    CategoriesKeywordsString = Replace(CategoriesKeywordsString, "&", " +")
    CategoriesKeywordsString = Replace(CategoriesKeywordsString, "-", " -")

    CategoriesKeywordsArr = Split(CategoriesKeywordsString, " ")
    FilterString = "CheckItem = ("

    For Each str in CategoriesKeywordsArr

        MinusSign = False
        PlusSign = False
        AsteriskSign = False

        If Left(str, 1) = "-" Then
            MinusSign = True
           str = Mid(str, 2)
        End If

        If Left(str, 1) = "+" Then
            PlusSign = True
           str = Mid(str, 2)
        End If

        If Left(str, 1) = "*" Then
           AsterikSign = True
           str = Mid(str, 2)
        End If

      If str<>"" Then

        If FilterString <> "CheckItem = (" Then
 
         	If PlusSign or AsterikSign Then
            		FilterString = FilterString & " AND ("	
         	Else 
            		FilterString = FilterString & " OR ("	
         	End If

        End If


        If MinusSign or AsterikSign Then
            ConditionsExist = False

            If FilterByCategory Then
                If ConditionsExist Then
                    FilterString = FilterString & " AND (InStr(LCase(RSSCategories), """ & str & """) = 0)"
                Else
                    FilterString = FilterString & "(InStr(LCase(RSSCategories), """ & str & """) = 0)"
                End If

                ConditionsExist = True
            End If

        Else

            ConditionsExist = False

            If FilterByCategory Then
                If ConditionsExist Then
                    FilterString = FilterString & " OR (InStr(LCase(RSSCategories), """ & str & """) > 0)"
                Else
                    FilterString = FilterString & "(InStr(LCase(RSSCategories), """ & str & """) > 0)"
                End If
                ConditionsExist = True
            End If

        End If

        FilterString = FilterString & ") "
     End If ' If str<>"" condition
    Next

    If FilterString = "CheckItem = (" Then
        FilterString = "CheckItem = True"
    End If

    GetCategoriesKeywordsFilter = FilterString
 End Function

' ######## get channel information
Sub GetChannel(xmlDOM)

ChannelTitle = ""
ChannelURL = ""
ChannelDescription = ""

	' get channel title 
	Set tempObj = xmlDOM.getElementsByTagName("title")
 	If Not tempObj is Nothing Then
		Set tempObj = tempObj.item(0)
	 	If Not tempObj is Nothing Then
			Set tempObj = tempObj.childNodes.item(0)
		End If
		 	If Not tempObj is Nothing Then
				ChannelTitle = tempObj.text
			End If
	End If

	' get channel url
	Set tempObj = xmlDOM.getElementsByTagName("link")
 	If Not tempObj is Nothing Then
		Set tempObj = tempObj.item(0)
	 	If Not tempObj is Nothing Then
			Set tempObj = tempObj.childNodes.item(0)
		End If
		 	If Not tempObj is Nothing Then
				ChannelURL = tempObj.text
			End If
	End If


	' get channel description
	Set tempObj = xmlDOM.getElementsByTagName("description")
 	If Not tempObj is Nothing Then
		Set tempObj = tempObj.item(0)
	 	If Not tempObj is Nothing Then
			Set tempObj = tempObj.childNodes.item(0)
		End If
		 	If Not tempObj is Nothing Then
				ChannelDescription = tempObj.text
			End If
	End If

    If ChannelURL = "" Then
	ChannelURL = URLToRSS
    End If

    If ChannelDescription = "" Then
	ChannelDescription = ChannelTitle
    End If

End Sub

 ' ##### Get RSSItems from xmlDOM #########
 Function GetRSS(xmlDOM)



     ' Collect filtered "items" from downloaded RSS
     Set RSSItems = xmlDOM.getElementsByTagName("item")

     ' /// If Not <item>..</item> entries, Then try to get <entry>..</entry>
     If RSSItems.Length <= 0 Then 
        Set RSSItems = xmlDOM.getElementsByTagName("entry") ' get <entry> items
     End If 

  GetChannel(xmlDOM) ' get information about channel

     Set GetRSS = RSSItems

 End Function
 
 ' ##### If RSSItems available, write to Response Header, RSSItems and Footer #########
 Function ShowRSS()
                  
    Set xmlDOM = GetXmlDOM(URLToRSS)

    If Not(xmlDOM is Nothing) Then 
        Set RSSItems = GetRSS(xmlDOM)

        Set xmlDOM = Nothing ' clear XML
	
	bShow = True

        If RSSItems.Length < 1 Then 
     		if DebugMode Then Response.Write "<br>ShowItems() - WARNING - items collection is empty" End If
		If ShowEmptyFeeds Then
			bShow = True
		Else 
			bShow = False
		End If
	End If
   	
        If bShow Then

            Dim TemplateHeader
            TemplateHeader = Replace(MainTemplateHeader, "{CHANNELTITLE}", ChannelTitle)
            TemplateHeader = Replace(TemplateHeader, "{CHANNELURL}", ChannelURL)
            TemplateHeader = Replace(TemplateHeader, "{CHANNELDESCRIPTION}", ChannelDescription)

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

 ' ##### strip HTML tags from description if set #########
Function stripHTML(ByVal sText)
   stripHTML= ""
   bFound = False
   Do While InStr(sText, "<")
      bFound = True
      stripHTML= stripHTML& " " & Left(sText, InStr(sText, "<")-1)
      sText = MID(sText, InStr(sText, ">") + 1)
   Loop
   Do While InStr(sText, "&nbsp;")
      bFound = True
      stripHTML= stripHTML& " " & Left(sText, InStr(sText, "&nbsp;")-1)
      sText = MID(sText, InStr(sText, "&nbsp;") + 1)
   Loop
   stripHTML= stripHTML& sText
   If not bFound Then stripHTML= sText
End Function


 ' ##### Entry point to script #########
Sub Main ' this function is intended to be called only when RSS2HTMLPro.asp is used standalone
 ShowRSS

 ' Show Error message, If available
 If Err.Number <> 0 Then
     Response.Write "Error: " & Err.Description
     Err.Clear 
 End If

 ' Response.End ' uncomment this For use in on-the-fly output
End Sub

If StandaloneScript Then
	Main() ' calling this only if the script is standalone. Set StandaloneScript variable to FALSE if script is called from MultiRSS2HTML.asp
End If


%>