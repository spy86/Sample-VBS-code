Set updateSession = CreateObject(“Microsoft.Update.Session”) 
Set updateSearcher = updateSession.CreateupdateSearcher() 
Set searchResult = _ updateSearcher.Search(“IsInstalled=0 and Type='Software'”) 
WScript.Echo searchResult.Updates.Count 