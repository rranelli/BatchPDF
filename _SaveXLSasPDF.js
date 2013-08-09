var fileSystem = new ActiveXObject("Scripting.FileSystemObject");
var xlsPath = WScript.Arguments(0);
xlsPath = fileSystem.GetAbsolutePathName(xlsPath);

var pdfPath = xlsPath.replace(/\.xls[^.]*$/, ".pdf");
var objExcel = null;

try
{
	objExcel = new ActiveXObject("Excel.Application");
    objExcel.Visible = false;
	
	var objExcel = objExcel.Workbooks.Open(xlsPath);

	numSheets = objExcel.Worksheets.Count
	iterac = 1
	
	while(iterac <= numSheets)
	{
		try
		{
//			the false tells the worksheet object to ADD
//			to the selection the worksheet 'iterac'
			objExcel.Worksheets(iterac).Select(false);
		}
		catch(err) {} // no action on error.
		//finally {}
		
		iterac = iterac + 1;
	}
	
	var wdFormatPdf = 57;
	WScript.Echo("Abrindo o arquivo: '" + objExcel.Name + "' e iniciando conversão.");

	objExcel.SaveAs(pdfPath, wdFormatPdf);
	WScript.Echo("Conversão realizada com Sucesso.");
	WScript.Echo("==========");
}
catch(err) {}

finally
{
    if (objExcel != null)
    {
		objExcel.Close(0);
	}
}