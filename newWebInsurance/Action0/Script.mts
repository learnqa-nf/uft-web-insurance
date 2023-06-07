RunAction "Login", oneIteration
RunAction "insertExcel", oneIteration, Parameter("Login", "oFullName"), Parameter("Login", "oRowDataLogin")
