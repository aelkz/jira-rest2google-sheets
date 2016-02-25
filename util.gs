function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getFieldName(heading,fields) {
  var index = fields.names.indexOf(heading);
  if ( index > -1) {
     return fields.ids[index]; 
  }
  return "";
}                 

function convertToMinutes(value,column,spreadsheet) {
  // columns: G,H,I,J
  if (value == null ||value == '') {
    value = column+'2:'+column+'100';
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet);
  var sheet = ss;
  var range = sheet.getRange(value);
  var status = '';
  
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var countNull = 0;
  
  for (var i = 1; i <= numRows; i++) {
    if (countNull >= 10) {
      break;
    }

    for (var j = 1; j <= numCols; j++) {
      var currentValue = range.getCell(i,j).getValue();
      status = currentValue;

      if (countNull >= 10) {
        break;
      }
      
      if (range.getCell(i, j).getValue() == '' || range.getCell(i, j).getValue() == null || range.getCell(i, j).getValues() == null) {
        countNull = countNull+1;
      }else {
        countNull = 0;
        var totalSeconds = range.getCell(i, j).getValue();
        var totalMinutes = totalSeconds / 60;
        var totalHours = totalMinutes / 60;

        range.getCell(i, j).setValue(totalHours);
      }
    }
  }
  countNull = 0;
}; 

function cleanCell(column,spreadsheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet);
  var sheet = ss;
  
  var value = column+2+':'+column+1000;
  var range = sheet.getRange(value);
  var icon_issue_value = 'H2'+':'+'H1000';
  var icon_issue_range = sheet.getRange(icon_issue_value);
  var icon_priority_value = 'K2'+':'+'K1000';
  var icon_priority_range = sheet.getRange(icon_priority_value);
  var icon_status_value = 'Q2'+':'+'Q1000';
  var icon_status_range = sheet.getRange(icon_status_value);

  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var countNull = 0;

  switch(column) {
    case "I":
      for (var i = 1; i <= numRows; i++) {
        if (countNull >= 3) {
          break;
        }
        
        for (var j = 1; j <= numCols; j++) {
          var currentValue = range.getCell(i,j).getValue();
          status = currentValue;
          
          if (countNull >= 3) {
            break;
          }
          
          if (range.getCell(i, j).getValue() == '' || range.getCell(i, j).getValue() == null || range.getCell(i, j).getValues() == null) {
            countNull = countNull+1;
          }else {
            countNull = 0;

            var data = range.getCell(i, j).getValue();
            var description = data;
            description = description.slice((description.indexOf("name=")+5),(description.indexOf("self=")-2));
            var iconUrl = data;
            iconUrl = iconUrl.slice((iconUrl.indexOf("iconUrl=")+8),(iconUrl.indexOf("subtask=")-2));
            range.getCell(i, j).setValue(description);
            icon_issue_range.getCell(i, j).setFormula("=image(\""+iconUrl+"\";4;16;16)");
          }
        }
      }
      break;
    case "L":
      for (var i = 1; i <= numRows; i++) {
        if (countNull >= 3) {
          break;
        }
        
        for (var j = 1; j <= numCols; j++) {
          var currentValue = range.getCell(i,j).getValue();
          status = currentValue;
          
          if (countNull >= 3) {
            break;
          }
          
          if (range.getCell(i, j).getValue() == '' || range.getCell(i, j).getValue() == null || range.getCell(i, j).getValues() == null) {
            countNull = countNull+1;
          }else {
            countNull = 0;

            var data = range.getCell(i, j).getValue();
            var description = data;
            description = description.slice((description.indexOf("name=")+5),(description.indexOf("self=")-2));
            var iconUrl = data;
            iconUrl = iconUrl.slice((iconUrl.indexOf("iconUrl=")+8),(iconUrl.indexOf("id=")-2));
            range.getCell(i, j).setValue(description);
            icon_priority_range.getCell(i, j).setFormula("=image(\""+iconUrl+"\";4;16;16)");
          }
        }
      }
      break;
    case "R":
      for (var i = 1; i <= numRows; i++) {
        if (countNull >= 3) {
          break;
        }
        
        for (var j = 1; j <= numCols; j++) {
          var currentValue = range.getCell(i,j).getValue();
          status = currentValue;
          
          if (countNull >= 3) {
            break;
          }
          
          if (range.getCell(i, j).getValue() == '' || range.getCell(i, j).getValue() == null || range.getCell(i, j).getValues() == null) {
            countNull = countNull+1;
          }else {
            countNull = 0;

            var data = range.getCell(i, j).getValue();
            var description = data;
            description = description.slice((description.indexOf("name=")+5),(description.indexOf("self=")-2));
            var iconUrl = data;
            iconUrl = iconUrl.slice((iconUrl.indexOf("iconUrl=")+8),(iconUrl.indexOf("id=")-2));
            range.getCell(i, j).setValue(description);
            icon_status_range.getCell(i, j).setFormula("=image(\""+iconUrl+"\";4;16;16)");
          }
        }
      }
      break;
    case "B":
      formatCellDate(numRows,numCols,range);
      break;
    case "C":
      formatCellDate(numRows,numCols,range);      
      break;
    case "D":
      formatCellDate(numRows,numCols,range);
      break;
    case "E":
      formatCellDate(numRows,numCols,range);      
      break;
    case "F":
      formatCellDate(numRows,numCols,range);      
      break;
  }
}

function getDateString(date_string) {
    var date = date_string.split("-");

    var dd = date[0];
    var mm = date[1];
    var yyyy = date[2];

    if(dd.length==1){
      dd="0"+dd;
    } 
    if(mm.length==1){
      mm=parseInt(mm)+1;
      mm="0"+mm;
    }
    return dd+"/"+mm+"/"+yyyy;
}

function formatCellDate(numRows,numCols,range) {
  var countNull = 0;
  for (var i = 1; i <= numRows; i++) {
    if (countNull >= 3) {
      break;
    }
    
    for (var j = 1; j <= numCols; j++) {
      var currentValue = range.getCell(i,j).getValue();
      status = currentValue;
      
      if (countNull >= 3) {
        break;
      }
      
      if (range.getCell(i, j).getValue() == '' || range.getCell(i, j).getValue() == null || range.getCell(i, j).getValues() == null) {
        countNull = countNull+1;
      }else {
        countNull = 0;
        
        var data = range.getCell(i, j).getValue();
        // 2015-08-01T11:45:00.000+0200
        data = data.slice(0,(data.indexOf("T")));
        range.getCell(i, j).setValue(getDateString(data));
      }
    }
  }
}

function formatFont(range) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var cells = sheet.getRange(range);
    
  cells.setFontFamily("Verdana");
  cells.setFontSize(8);
  cells.setWrap(true);
  cells.setHorizontalAlignment("center");  
};

function setAlign(range,align) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var cells = sheet.getRange(range);

  cells.setHorizontalAlignment(align);
};

function setBold(range) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var cells = sheet.getRange(range);

  cells.setFontWeight("bold");
};

function setUnBold(range) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var cells = sheet.getRange(range);

  cells.setFontWeight("normal");
};

function getActiveCellValue(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  var cell = ss.getActiveCell();
  var data = ss.getDataRange().getValues();
  
  var activeR = cell.getRow();
  var activeC = cell.getColumn();
 
  var activeCell = sheet.getRange("A"+activeR);
  //activeCell.setBackground('#ffff55');
  //activeCell.setFontSize(12);
  //activeCell.setFontColor('#ffff55');
  
  return activeCell.getValues();
};

function setCreationDateSingleRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var index = ss.getActiveCell().getRow();

  // DT. CRIAÇÃO
  var value = 'S'+index+':S'+index;

  var range = sheet.getRange(value);
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  
  var timezone = "GMT-3";
  var timestamp_format = "dd-MM-yyyy HH:mm"; // Timestamp Format. 
  
  for (var i = 1; i <= numRows; i++) {
    for (var j = 1; j <= numCols; j++) {
      var currentValue = range.getCell(i,j).getValue();
      if (currentValue == '' || currentValue == null || currentValue == undefined) {
        // define a data de criação do registro
        var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
        range.getCell(i,j).setValue(date);
      }else {
        // atualiza a última data de alteração (last update)
        setLastUpdateDateSingleRow(i,j,index);
        //
        //if (statusRange.getValue() == 'Aguardando Validação' || statusRange.getValue() == 'Aguardando Validação do Demandante' || statusRange.getValue() == 'INDRA/Negociar Solução de Contorno') {
        //  setConclusionDateSingleRow(i,j,index);
        //  setProductSingleRow();
        //  setStatusSingleRow();
        //}else if (statusRange.getValue() == 'Em Atendimento' || statusRange.getValue() == 'Em Análise') {
        //  clearConclusionDateSingleRow(i,j,index);
        //}
      }
    }
 }
};

function setLastUpdateDateSingleRow(i,j,index) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  if (index != 0 && index != 1) {
    // DT. ÚLTIMA ATUALIZAÇÃO
    var value = 'T'+index+':T'+index;
    
    var range = sheet.getRange(value);
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
    
    var timezone = "GMT-3";
    var timestamp_format = "dd-MM-yyyy HH:mm"; // Timestamp Format. 
    
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    range.getCell(i,j).setValue(date);
  }
  
};