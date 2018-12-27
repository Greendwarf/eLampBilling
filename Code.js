/**
* @NotOnlyCurrentDoc
**/

/* Checks if range1 is the same as range2 i.e. covers the same cells */
function eLampIsSameRange(range1, range2) {
  var isSameRange = range1.getRow() == range2.getRow() && range1.getLastRow() == range2.getLastRow() && range1.getColumn() == range2.getColumn() && range1.getLastColumn() == range2.getLastColumn() ? true : false;
  return isSameRange;
}

/* Checks if range 2 encompasses range1 */
function eLampIsInRange(range1, range2): boolean {
  return isInRange = range1.getRow() >= range2.getRow() && range1.getLastRow() <= range2.getLastRow() && range1.getColumn() >= range2.getColumn() && range1.getLastColumn() <= range2.getLastColumn() ? true : false;
}

/* Checks if range 1 runs across range 2, i.e. might start in range 2 and end outside range 2, or might start outside range 2 and end in range 2, or might start and end outside range 2 but run across range 2, or might be included in range2 */
function eLampIsAcrossRange(range1, range2) {
  
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  
  ui.createAddonMenu()
    .addItem("Gérer Facturation", 'manageBill')
    .addItem("Gérer Commandes", "manageOrders")
    .addItem("Gérer Clients", "manageAccounts")
    .addToUi();
}

function manageBill() {
  let html = HtmlService.createTemplateFromFile("IndexManage");
  html.objectType = "Bills";
  html = html.evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
}

function manageOrders() {
  var html = HtmlService.createTemplateFromFile("IndexManage");
  html.objectType = "Orders";
  html = html.evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
}

function manageAccounts() {
  var html = HtmlService.createTemplateFromFile("IndexManage");
  html.objectType = "Accounts";
  html = html.evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
}

function createNewBill() {
  var html = HtmlService.createTemplateFromFile("IndexCreate");
  html.objectType = "Bills";
  html = html.evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Créer une nouvelle facture");
}

function createNewOrder() {
  var html = HtmlService.createTemplateFromFile("IndexCreate");
  html.objectType = "Orders";
  html = html.evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Créer une nouvelle facture");
}

function createNewAccount() {
  var html = HtmlService.createTemplateFromFile("IndexCreate");
  html.objectType = "Accounts";
  html = html.evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Créer une nouvelle facture");
}

function getParameterData() {
  var billParametersSpreadsheetValues = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Paramètres").getValues();
  var parameterData = [];
  
  billParametersSpreadsheetValues.forEach(function(element) {
    parameterData.push({
      'column-title' : element[0],
      'parameter-title' : element[1],
      'parameter-type' : element[2]
    });
  });
  
  return parameterData;
}

function loadBills() {

  var billsSpreadsheetValues = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Bills").getValues();
  var billsKeysSpreadsheetValues = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Paramètres").getValues();
  var billsData = [];
  
  billsSpreadsheetValues[0] = billsSpreadsheetValues
    .filter(function(element, index) { return index == 0; })
    .reduce(function(bout1, bout2) { return bout1.concat(bout2); })
    .map(function(header) {
      var newHeader = header;
      billsKeysSpreadsheetValues.some(function(keyPairing) {
        if(header == keyPairing[0]) { newHeader = keyPairing[1]; console.log(newHeader); return true; }
        return false;
      });
      return newHeader;
    });
    
    console.log("headers : " +billsSpreadsheetValues[0]);
  
  billsData = billsSpreadsheetValues
    .filter(function(element, index) { return index > 0; })
    .map(function(line, lineIndex) {
      var newObject = {};
      line.forEach(function(cell, cellIndex) {
        newObject[billsSpreadsheetValues[0][cellIndex]] = Object.prototype.toString.call(cell) === '[object Date]' ? cell.toString() : cell;
      });
      return newObject;
    });
  
  return billsData;

}

function etablirBill(formObject) {
  
  var url = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Bill_Template").getValues()[0][0];
  var billTemplateID = getIdFrom(url);
  var billTemplateFile = DriveApp.getFileById(billTemplateID);
  var billTemplateFolder = getImmediateFolder(billTemplateFile);
  var existingBills = billTemplateFolder.getFilesByName(formObject['nom-facture']);
  var toBeFormatedAsDate = ["debut-prestation", "fin-prestation", "debut-licences", "fin-licences"];
  var toBeFormatedAsNumber = ["prix-prestation", "prix-unitaire-licences", "prix-total-licences", ];
  
  if(getFileIteratorSize(existingBills) > 0) {
    
    SpreadsheetApp.getUi().alert("La facture existe déjà : " + billTemplateFolder.getFilesByName(formObject['nom-facture']).next().getUrl());
  
  } else {
    
    var body = DocumentApp.openById(billTemplateFile.makeCopy(formObject['nom-facture']).getId()).getBody();
    var text = body.editAsText();
    var markupsBills = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Paramètres").getValues();
    
    for( i=0; i < markupsBills.length; i++) {
      text.replaceText("<\\["+markupsBills[i][1]+"\\]>", formObject[markupsBills[i][1]]);
    }
    
    var prixTotal = 0;
    var TVA = 0;
    var totalTTC = 0;
    
    if(formObject['type-facture'] == "Prestation") {
    
      removeDocSection(body, "<\\[zone-licences\\]>", "<\\[zone-licences\\]/>");
      removeDocMarkups(body,["<\\[zone-prestation]>", "<\\[zone-prestation]/>"]);
      prixTotal = parseFloat(formObject['prix-prestation']);
          
    } else if (formObject['type-facture'] == "Licences") {
    
      removeDocSection(body, "<\\[zone-prestation\\]>", "<\\[zone-prestation\\]/>");
      removeDocMarkups(body,["<\\[zone-licences]>", "<\\[zone-licences]/>"]);
      prixTotal = parseFloat(formObject['prix-total-licences']);
    
    } else {
    
      prixTotal = parseFloat(formObject['prix-prestation']) + parseFloat(formObject['prix-total-licences']);
    
    }
    
    TVA = prixTotal * 0.2;
    totalTTC = prixTotal + TVA;
    
    text.replaceText("<\\[prix-total-HT\\]>", prixTotal);
    text.replaceText("<\\[tva-total\\]>", TVA);
    text.replaceText("<\\[total-ttc\\]>", totalTTC);       

  }
}

function envoyerBill(formObject) {
  console.log('blob_envoyer');
  console.log(formObject);
}

function payerBill(formObject) {

  var billsSpreadsheetValues = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Bills").getValues();
  var billParametersSpreadsheetValues = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Paramètres").getValues();
  var billsData = [];
  var headerValues = [];

  billParametersSpreadsheetValues.forEach(function(parameter) {
    billsSpreadsheetValues[0].some(function(header, index) {
      if(header == parameter[0]) {
        headerValues.push({
          'spreadsheet-title': header,
          'parameter-title': parameter[1],
          'column-position': index
        });
        return true;
      }
      return false;
    });
  });
  
  billsSpreadsheetValues.forEach(function(line, lineIndex) {
    if(lineIndex != 0) {
      var dataLine = {};
      line.forEach(function(cell, cellIndex) {
        headerValues.some(function(header) {
          if(header['column-position']==cellIndex) { 
            var parameterTitle = header['parameter-title'];
            dataLine[parameterTitle] = cell; 
            return true;
          }
          return false;
        });
      });
      billsData.push(dataLine);
    }
  });
  
  var billReturnValues = [];
  billReturnValues[0] = [];
  
  headerValues.forEach(function(header) {
    billReturnValues[0][header['column-position']] = header['spreadsheet-title'];
  });

  billsData.forEach(function(dataLine, dataLineIndex) {
    if(dataLine['nom-facture'] == formObject['nom-facture']) {
      dataLine['facture-payee'] = true;
    }
    var tempArr = [];
    headerValues.forEach(function(header) {
      tempArr[header['column-position']] = dataLine[header['parameter-title']];
    });
    billReturnValues.push(tempArr);
  });
  
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Bills").setValues(billReturnValues);
  
  var returnData = {};
  
  billsData.forEach(function(dataLine) {
    returnData[dataLine['nom-facture']] = {};
    Object.keys(dataLine).map(function(dataIndex) {
      returnData[dataLine['nom-facture']][dataIndex] = dataLine[dataIndex];
    });
  });
  
  return billsData[0];
  
}

function successAndUpdate(action) {
  SpreadsheetApp.getUi().alert("lol " + action);
  return loadBills();
}

function getImmediateFolder(file) {
  var parentFolders = file.getParents();
  while(parentFolders.hasNext()) {
    var folder = parentFolders.next();
  }
  return folder;
}

function getFileIteratorSize(fileIterator) {
  var i =0;
  while(fileIterator.hasNext()) {
    i++;
    fileIterator.next();
  }
  return i;
}

function removeDocMarkups(body, markups) {

  for(i=0; i < markups.length; i++) {
  
    body.editAsText().replaceText(markups[i], "");
    
  }
}

function removeDocSection(body, startRegExp, endRegExp) {

  var totalchildren = body.getNumChildren();
  var found = 0;
  
  for (var i=totalchildren-1; i >= 0; i--) {
    
    console.log("body child : " + body.getChild(i));
    
    if(body.getChild(i).getType() == DocumentApp.ElementType.TABLE) {
      var thisBodyChild = body.getChild(i).asTable();
    } else {
      var thisBodyChild = body.getChild(i).asText();
    }
    
    if (found == 0) {      
      
      if (thisBodyChild.findText(endRegExp) !== null) {
        
        thisBodyChild.replaceText(".*"+endRegExp,"");
        thisBodyChild.appendText(" ");
        
        if (thisBodyChild.findText("^ *$") !== null) {
          thisBodyChild.removeFromParent();
        }
        
        found = 1;
        
      }
      
    } else {
      
      if (thisBodyChild.findText(startRegExp) !== null) {
        
        thisBodyChild.replaceText(startRegExp+".*","");
        thisBodyChild.appendText(" ");
        
        if (thisBodyChild.findText("^ *$") !== null) {
          thisBodyChild.removeFromParent();
        }
        
        found = 0;
        
      } else {
        
        thisBodyChild.removeFromParent();
        
      }
    }  
  }
}

function include(filename) {
    console.log("filename babilobabd : " + filename);
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getIdFrom(url) {
  var id = "";
  var parts = url.split(/^(([^:\/?#]+):)?(\/\/([^\/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/);
  if (url.indexOf('?id=') >= 0){
     id = (parts[6].split("=")[1]).replace("&usp","");
     return id;
   } else {
   id = parts[5].split("/");
   //Using sort to get the id as it is the longest element. 
   var sortArr = id.sort(function(a,b){return b.length - a.length});
   id = sortArr[0];
   return id;
   }
 }