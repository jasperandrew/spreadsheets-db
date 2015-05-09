function query()
{
    var querySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Query Form");
    var sheetHeaders = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data").getRange("A1:ZZ1").getValues();
    
    var headerA1 = getHeaderA1(sheetHeaders);
    
    var selection = getSelection(querySheet, "A2:A11", headerA1);
    if(selection == null){
	errorCell(querySheet, "C8", "Error: Invalid Selection");
	return "Error: Invalid Selection(s)";
    }
    
    var conditions = getConditions(querySheet, "B2:D4", headerA1);
    if(conditions == null){
	errorCell(querySheet, "C8", "Error: Invalid Condition(s)");
	return "Error: Invalid Condition(s)";
    }

    var sort = "";
    if(getVal(querySheet, "E2") == "ascending")
	sort = " order by " + headerA1[getVal(querySheet, "E3")];
    else if(getVal(querySheet, "E2") == "descending")
	sort = " order by " + headerA1[getVal(querySheet, "E3")] + " desc";
    
    var queryString = "=QUERY(Data!A1:ZZ, \"Select " + selection + " where " + conditions + sort + "\", 1)";
    
    var resultsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Query Results");
    querySheet.getRange("C8").setFontColor("green");
    querySheet.getRange("C8").setValue("Success!");
    resultsSheet.getRange("A3").setValue(queryString);
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(resultsSheet);
};

function resizeCols()
{
    var resultsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Query Results");
    var range = resultsSheet.getRange("A3:ZZ3").getValues();
    
    var i = 0;
    while(range[0][i] != null && range[0][i] != "") i++;

    resizeColumns(resultsSheet, 1, i);
}

function getVal(sheet, a1)
{
    if(sheet.getRange(a1).getValue() != "")
	return sheet.getRange(a1).getValue();
    else
	return null;
};

function getSelection(sheet, a1Range, headerA1)
{
    var selStr = "";
    var empty = true;
    var selectVals = sheet.getRange(a1Range).getValues();
    for(var i = 0; i < selectVals.length; i++){
	if(selectVals[i][0] == ""){
	}else if(selectVals[i][0] == "ALL"){
	    return "*";
	}else{
	    if(empty == true)
		selStr = headerA1[selectVals[i][0]];
	    else
		selStr = selStr + ", " + headerA1[selectVals[i][0]];
	    empty = false;
	}
    }
    
    if(empty == true)
	return null;
    
    return selStr;
};

function getConditions(sheet, a1Range, headerA1)
{
    var condArr = [null, null, null];
    var empty = true;
    var conditionVals = sheet.getRange(a1Range).getValues();
    for(var i = 0; i < conditionVals.length; i++){
	var hasVal = false;
	for(var j = 0; j < 3; j++){
	    if(conditionVals[i][j] === ""){
		if(hasVal) return null;
	    }else
		hasVal = true;
	}
	
	if(!hasVal) continue;
	
	var field = headerA1[conditionVals[i][0]];
	var operator = getOperator(conditionVals[i][1]);
	var value = conditionVals[i][2];
	
	if(!isNaN(value))
	    condArr[i] = field + " " + operator + " " + value;
	else
	    condArr[i] = field + " " + operator + " '" + value + "'";
	
	empty = false;
    }
    
    if(empty)
	return null;
    
    var condStr = "";
    hasVal = false;
    for(var i = 0; i < 3; i++){
	if(condArr[i]){
	    condStr += (hasVal ? " and " + condArr[i] : condArr[i]);
	    hasVal = true;
	}
    }
    return condStr;
}

function getOperator(val)
{
    switch(val){
    case "equals":
	return "=";
    case "does not equal":
	return "!=";
    case "is more than":
	return ">";
    case "is less than":
	return "<";
    case "is more than/equal to":
	return ">=";
    case "is less than/equal to":
	return "<=";
    default:
	return null;
    } 
};

function getHeaderA1(headerVals)
{
    retObj = new Object();
    for(var i = 0; i < headerVals[0].length; i++){
	if(headerVals[0][i] != "")
	    retObj[headerVals[0][i]] = columnToLetter(i+1);
    }
    return retObj;
};

function errorCell(sheet, a1, msg)
{
    sheet.getRange(a1).setFontColor("red");
    sheet.getRange(a1).setValue(msg);
};

function resizeColumns(sheet, startCol, endCol)
{
    for(var i = startCol; i < endCol+1; i++)
	sheet.autoResizeColumn(i);
};

function columnToLetter(colIndex)
{
    var tmp, letter = '';
    while (colIndex > 0)
    {
	tmp = (colIndex-1) % 26;
	letter = String.fromCharCode(tmp + 65) + letter;
	colIndex = (colIndex - tmp - 1) / 26;
    }
    return letter;
};
