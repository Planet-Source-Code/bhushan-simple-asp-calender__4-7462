<div align="center">

## Simple ASP Calender


</div>

### Description

Simple ASP Calender
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bhushan\-](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bhushan.md)
**Level**          |Intermediate
**User Rating**    |4.9 (44 globes from 9 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bhushan-simple-asp-calender__4-7462/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<%
'Function to Return the number of Days in a month
function findMonth(strDate, strYear)
	dim days
	if strDate = 4 or strDate = 6 or strDate = 9 or strDate = 11 then
		days = 30
	elseif strDate = 2 AND strYear/4 = int(strYear/4) then
		days = 29
	elseif strDate = 2 then
		days = 28
	else
		days = 31
	end if
	findMonth = days
end function
'Function will return the numeric value last or Next Month
function fnChangeMonth(strMonth, strDirection)
	if strDirection = "previous" then
		if strMonth = 1 then
			tempstrMonth = 12
		else
			tempstrMonth = strMonth - 1
		end if
	else
		if strMonth > 11 then
			tempstrMonth = 1
		else
			tempstrMonth = strMonth + 1
		end if
	end if
	fnChangeMonth = tempstrMonth
end function
'Function will return a date format from the qstring dd
'I use querystring called dd in this format 01012000 this just makes that
'into a date format
function formatQstring(strQstring)
	ddLength = Len(strQstring)
	tempYear = Right(strQstring,4)
	tempDay = Right(strQstring,6)
	tempDay = Left(tempday,2)
	tempMonth = Left(strQstring,ddLength - 6)
	strQstring = tempMonth & "/" & tempDay & "/" & tempYear
	formatQstring = formatdatetime(strQstring,2)
end function
'Find the numeric value of the first day in the month (Monday = 2...)
function formatFirstDay(strFirstDay)
	strFirstDay = WeekDay(Left(strFirstDay,2) & "/1/" & Right(strFirstDay,4))
	formatFirstDay = strFirstDay
end function
'Make the Hyperlink for Previous or Next Month
function makeLink(strDate, strLinkType)
	if strdate = "" then
		strdate = Month(DisplayDate) & "01" & Year(DisplayDate)
	end if
	theLength = len(strdate)
	theYear = Right(strdate,4)
	theMonth = Left(strdate, theLength-6)
	if strLinkType = "previous" then
		theMonth = fnChangeMonth((Left(theMonth,2)),"previous")
		if theMonth = 12 then
			theYear = Right(strDate,4) - 1
		else
			theYear = Right(strDate,4)
		end if
	else
		theMonth = fnChangeMonth((Left(theMonth,2)),"Next")
		if theMonth = 1 then
			theYear = Right(strDate,4) + 1
		else
			theYear = Right(strDate,4)
		end if
	end if
	if len(theMonth) <> 2 then
		theMonth = "0" & theMonth
	end if
	strdate = theMonth & "01" & theYear
	makelink = strdate
end function
'Determine if there is a Calendar Request to show a month otherwise show this month
if Request("dd") = "" then
	DisplayDate = Date()
	ShowYear = Year(Date)
	FirstDayofMonth = WeekDay(Month(Date) & "/1/" & ShowYear)
else
	ShowYear = Right(Request("dd"),4)
	DisplayDate = formatQstring(Request("dd"))
	FirstDayofMonth = WeekDay(DisplayDate)
end if
previousMonth = findMonth(fnChangeMonth(Month(DisplayDate),"previous"), ShowYear) - FirstDayofMonth + 1
thisMonth = 0
nextMonth = 0
weekdaynum = 0
DisplayMonth = Month(DisplayDate)
If len(DisplayMonth) <> 2 then
	DisplayMonth = "0" & DisplayMonth
end if
DisplayYear = Right((DisplayDate),4)
html = "<TR><TD colspan=""7""><center><b>" & MonthName(month(DisplayDate), 0) & "&nbsp;" & ShowYear & "</b></center></TD></TR>" & vbcr
html = html & "<TR><TD align=""center"" class=""date"">Su</TD><TD align=""center"" class=""date"">Mo</TD><TD align=""center"" class=""date"">Tu</TD><TD align=""center"" class=""date"">We</TD><TD align=""center"" class=""date"">Th</TD><TD align=""center"" class=""date"">Fr</TD><TD align=""center"" class=""date"">Sa</TD></TR>"
for tablecell = 1 to 42
	if weekdaynum = 7 then
		weekdaynum = 0
	end if
	weekdaynum = weekdaynum + 1
	inc = inc + 1
	if inc < FirstDayofMonth then
		previousMonth = previousMonth + 1
		html = html & "<TD align=""center"" class=""dateother"">" & previousMonth & "</TD>" & vbcr
	elseif thisMonth < findMonth(DisplayMonth, ShowYear) then
		thisMonth = thisMonth + 1
		html = html & "<TD align=""center"" class=""date""><A HREF=""day.asp?at=" & Request("at") & "&sguid=" & Request("sguid") & "&dd=" & DisplayMonth & thisMonth & DisplayYear & "&wd=" & weekdaynum & """>" & thisMonth & "</A></TD>" & vbcr
	else
		nextMonth = nextMonth + 1
		html = html & "<TD align=""center"" class=""dateother"">" & nextMonth & "</TD>" & vbcr
	end if
	if tablecell/7 = int(tablecell/7) then
		html = html & "</tr><tr>" & vbcr
	end if
Next
html = html & "<TR><TD align=""center"" colspan=""7""><A HREF=""cal_small.asp?dd=" & makeLink(Request("dd"),"previous") & """>Previous</A>&nbsp;&nbsp;&nbsp;<A HREF=""cal_small.asp?dd=" & makeLink(Request("dd"),"next") & """>Next</A></TD><TR>"
%>
<HTML>
<HEAD>
</HEAD>
<BODY>
<TABLE WIDTH="200px" BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<%=html%>
</TABLE>
</BODY>
</HTML>
```

