Browser("InsuranceWeb: Home").Sync
Browser("InsuranceWeb: Home").Page("InsuranceWeb: Home").Sync
varComa = ","
varSparator = "|"
For i = 1 To DataTable.GetSheet("Login").GetRowCount
	Browser("InsuranceWeb: Home").Page("InsuranceWeb: Home").WebEdit("login-form:email").Set DataTable("email", "Login")
	Browser("InsuranceWeb: Home").Page("InsuranceWeb: Home").WebEdit("login-form:password").Set DataTable("password", "Login")
	Browser("InsuranceWeb: Home").Page("InsuranceWeb: Home").Image("Login").Click
	Browser("InsuranceWeb: Home").Page("InsuranceWeb: Home_2").Sync
	Browser("InsuranceWeb: Home").Page("InsuranceWeb: Home_2").Image("details").Click
	Browser("InsuranceWeb: Home").Page("InsuranceWeb: Account").Sync @@ script infofile_;_ZIP::ssf6.xml_;_
	Browser("InsuranceWeb: Home").Page("InsuranceWeb: Account").WebElement("fnameObj").WaitProperty "visible", True, 5000 ' using Synch point
	fullname = fullname + Browser("InsuranceWeb: Home").Page("InsuranceWeb: Account").WebElement("fnameObj").GetROProperty("outertext") + varComa + varSparator ' collect data
	
	Browser("InsuranceWeb: Home").Page("InsuranceWeb: Account").Link("Home").Click
	Browser("InsuranceWeb: Home").Page("InsuranceWeb: Account").Image("logout").Click

	DataTable.SetNextRow

	Set x = New oTest
	x.test
	
Next
Parameter("oFullName") = fullname
Parameter("oRowDataLogin") = DataTable.GetSheet("Login").GetRowCount
Browser("InsuranceWeb: Home").Close

Class oTest

	Sub test()
		Print "testingClass"
	End Sub
End Class
