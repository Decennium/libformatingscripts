dtmTargetDate = #10/1/2016#
for i =0 to 35 step 7
	d = dateadd("d",i,dtmTargetDate)
	wscript.echo d & " Is " & WeekdayOfMonth(d) & " of " & DatePart("m",d)
next

Function WeekdayOfMonth(dtmTargetDate)
	dtmDay = DatePart("d", dtmTargetDate)
	dtmMonth = DatePart("m", dtmTargetDate)
	dtmYear = DatePart("yyyy", dtmTargetDate)

	dtmStartDate = dtmMonth & "/1/" & dtmYear
	dtmStartDate = CDate(dtmStartDate)

	intWeekday = Weekday(dtmStartDate)
	intAddon = 8 - intWeekday

	intWeek1 = intAddOn
	intWeek2 = intWeek1 + 7
	intWeek3 = intWeek2 + 7
	intWeek4 = intWeek3 + 7
	intWeek5 = intWeek4 + 7
	intWeek6 = intWeek5 + 7

	If dtmDay > intWeek5 And dtmDay <= intWeek6 Then
		strWeek = 6
	ElseIf dtmDay > intWeek4 And dtmDay <= intWeek5 Then
		strWeek = 5
	ElseIf dtmDay > intWeek3 And dtmDay <= intWeek4 Then
		strWeek = 4
	ElseIf dtmDay > intWeek2 And dtmDay <= intWeek3 Then
		strWeek = 3
	ElseIf dtmDay > intWeek1 And dtmDay <= intWeek2 Then
		strWeek = 2
	ElseIf dtmDay <= intWeek1 Then
		strWeek = 1
	End If

	WeekdayOfMonth = strWeek
End Function