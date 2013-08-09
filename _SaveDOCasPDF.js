var fileSystem = new ActiveXObject("Scripting.FileSystemObject");
var docPath = WScript.Arguments(0);
docPath = fileSystem.GetAbsolutePathName(docPath);

var pdfPath = docPath.replace(/\.doc[^.]*$/, ".pdf");
var objWord = null;
var strBrokenReport = null;

try
{
    objWord = new ActiveXObject("Word.Application");
    objWord.Visible = false;

    var objDoc = objWord.Documents.Open(docPath);
	
	// this sentinel variable checks if there is a dangling reference.
	var runSentinel = false
	try
	{
		runSentinel = objDoc.Application.Run("isRefBroke");
	}
	catch(e) {}
	
	WScript.Echo("Abrindo o arquivo: '" + objDoc.Name + "' e iniciando conversão.");
	
	// action if there is a problem.
	if(runSentinel != false)
	{
		WScript.Echo("WARNING!! THE DOCUMENT "+ objDoc.Name + " HAS INVALID REFERENCES!!!");
		WScript.Echo("")
	}
	
    var wdFormatPdf = 17;
	
    objWord.DisplayAlerts = false // deactivating prompts
	objDoc.SaveAs(pdfPath, wdFormatPdf); // saving
	objWord.DisplayAlerts = true // activating prompts
	
	WScript.Echo("Conversão realizada com Sucesso.");
	WScript.Echo("==========");
}
catch(e){}

finally
{
	if(objDoc != null)
	{
		objDoc.close(0)
	}

	// closing the word application
	if (objWord != null)
    {
		objWord.Quit();
    }
}