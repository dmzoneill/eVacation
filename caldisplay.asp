<!--#include virtual="/eVacation/common/appglobal.asp" -->
<!-- #include virtual="/evacation/common/objects/calendar.asp" -->
<!--#include virtual="/eVacation/common/calendarfunctions.asp" -->
<%
strCurrentPageName = "Calendar"

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser

	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName
%>
<script language="javascript">
var monthStart;
var monthEnd;
var yearOfCal;


//FUNCTIONS TO CHOOSE SECTIONS TO VIEW
function qtr1() {
	var now = new Date();
	
	monthStart = 1;
	monthEnd = 3;
	yearOfCal = now.getFullYear();
	
	document.form1.startMonth.value = monthStart;
	document.form1.monthsToDisplay.value = monthEnd;
	
	if (document.form1.yearOfCal.value == '')  {
		document.form1.yearOfCal.value = yearOfCal;
	}

}

function qtr2() {
	var now = new Date();
	
	monthStart = 4;
	monthEnd = 3;
	yearOfCal = now.getFullYear();
	
	document.form1.startMonth.value = monthStart;
	document.form1.monthsToDisplay.value = monthEnd;
	
	if (document.form1.yearOfCal.value == '')  {
		document.form1.yearOfCal.value = yearOfCal;
	}

}

function qtr3() {
	var now = new Date();
	
	monthStart = 7;
	monthEnd = 3;
	yearOfCal = now.getFullYear();
	
	document.form1.startMonth.value = monthStart;
	document.form1.monthsToDisplay.value = monthEnd;
	
	if (document.form1.yearOfCal.value == '')  {
		document.form1.yearOfCal.value = yearOfCal;
	}

}

function qtr4() {
	var now = new Date();
	
	monthStart = 10;
	monthEnd = 3;
	yearOfCal = now.getFullYear();
	
	document.form1.startMonth.value = monthStart;
	document.form1.monthsToDisplay.value = monthEnd;
	
	if (document.form1.yearOfCal.value == '')  {
		document.form1.yearOfCal.value = yearOfCal;
	}

}

//FUNCTION TO CHOOSE TO VIEW THE FULL YEAR
function fullYear() {
	var now = new Date();
	
	monthStart = 1;
	monthEnd = 12;
	yearOfCal = now.getFullYear();
	
	document.form1.startMonth.value = monthStart;
	document.form1.monthsToDisplay.value = monthEnd;
	
	if (document.form1.yearOfCal.value == '') {
		document.form1.yearOfCal.value = yearOfCal;
	}

}


</script>
<table class='pageContentWidth'>
<form method="post" name="form1" action="yearlycalendar.asp" >
<tr>
<!--BUTTONS TO FILL THE FIELDS WITH INFORMATION-->
	<td colspan="2" align="center"><input type="button" value="1st QTR" onclick="qtr1()"  class="tbcontent">
	<input type="button" value="2nd QTR" onclick="qtr2()" class="tbcontent">
	<input type="button" value="3rd QTR" onclick="qtr3()" class="tbcontent">
	<input type="button" value="4th QTR" onclick="qtr4()" class="tbcontent">
	<input type="button" value="Full Year" onclick="fullYear()" class="tbcontent"></td>
</tr>	
<tr>
	<td style='width:50%'>First Month: <tt>(As Number, e.g. January = 1)</tt></td>	
	<td style='width:50%'><input type="text" name="startMonth" ID="Text1"></td>
</tr>
<tr>	
	<td>Year: <tt>(As Number, e.g. 2007)</tt></td>
	<td><input type="text" name="yearOfCal" ID="Text2"></td>
</tr>
<tr>	
	<td>Months to Display: <tt>(e.g. 12 = Full Year)</tt></td>
	<td><input type="text" name="monthsToDisplay" ID="Text3"></td>
</tr>
<tr>
<!--BUTTON TO SUBMIT FORM-->
<td colspan="2" align="center">
<input type="submit" value="Submit"  class="tbcontent">
</td>
</tr>
</form>
</table>
<br class="small">
<%
mWritePageFooter

%>
