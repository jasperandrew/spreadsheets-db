function add()
{
    var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
    var addForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Add Student");
    var newStudentInfo = addForm.getRange("C2:C57").getValues();
    
    for(var i = 0; i < 3; i++){
	if(newStudentInfo[i][0] == ""){
	    errorCell(addForm, "D1", "Error: B-Number, First Name, and Last Name required");
	    return "Error: B-Number, First Name, and Last Name required"
	}
    }
    
    newStudentInfo = getNewRow(newStudentInfo);
    dataSheet.appendRow(newStudentInfo);
    
    addForm.getRange("C2:C57").clearContent();
    
    addForm.getRange("D1").setFontColor("green");
    addForm.getRange("D1").setValue("Success!");
}

function getNewRow(values)
{
    var row = new Array(values.length);
    for(var i = 0; i < values.length; i++)
	row[i] = values[i][0];
    return row;
}

function errorCell(sheet, a1, msg)
{
    sheet.getRange(a1).setFontColor("red");
    sheet.getRange(a1).setValue(msg); 
};
