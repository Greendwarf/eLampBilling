/**
* @NotOnlyCurrentDoc
**/

interface elmHtmlService extends GoogleAppsScript.HTML.HtmlTemplate {
    objectType?: string;
    authorizationUrl?: string;
}

const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

/* Checks if range1 is the same as range2 i.e. covers the same cells */
function eLampIsSameRange(range1, range2) {
  return range1.getRow() == range2.getRow() && range1.getLastRow() == range2.getLastRow() && range1.getColumn() == range2.getColumn() && range1.getLastColumn() == range2.getLastColumn() ? true : false;
}

/* Checks if range 2 encompasses range1 */
function eLampIsInRange(range1, range2) {
  return range1.getRow() >= range2.getRow() && range1.getLastRow() <= range2.getLastRow() && range1.getColumn() >= range2.getColumn() && range1.getLastColumn() <= range2.getLastColumn() ? true : false;
}

/* Checks if range 1 runs across range 2, i.e. might start in range 2 and end outside range 2, or might start outside range 2 and end in range 2, or might start and end outside range 2 but run across range 2, or might be included in range2 */
function eLampIsAcrossRange(range1, range2) {
  
}

function uniq(a) {
  var seen = {};
  return a.filter(function(item) {
      return seen.hasOwnProperty(item) ? false : (seen[item] = true);
  });
}


function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  let ui = SpreadsheetApp.getUi();
  
  ui.createAddonMenu()
    .addItem("Gérer Facturation", 'manageBill')
    .addItem("Gérer Commandes", "manageOrders")
    .addItem("Gérer Clients", "manageAccounts")
    .addToUi();
}

function manageBill(): void {
  let html: elmHtmlService = HtmlService.createTemplateFromFile("indexhtml/IndexManage");
  html.objectType = "Bills";
  SpreadsheetApp.getUi().showSidebar(html.evaluate());
}

function manageOrders(): void {
  let html: elmHtmlService = HtmlService.createTemplateFromFile("indexhtml/IndexManage");
  html.objectType = "Orders";
  SpreadsheetApp.getUi().showSidebar(html.evaluate());
}

function manageAccounts(): void {
  var pipedriveService = getPipedriveService();
  if(!pipedriveService.hasAccess()) {
    let authorizationUrl = pipedriveService.getAuthorizationUrl();
    let template: elmHtmlService = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Reopen the sidebar when the authorization is complete.');
    template.authorizationUrl = authorizationUrl;
    let page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    var html: elmHtmlService = HtmlService.createTemplateFromFile("indexhtml/IndexManage");
    html.objectType = "Accounts";
    SpreadsheetApp.getUi().showSidebar(html.evaluate());
  }
}

function createNewBill(): void {
  let html: elmHtmlService = HtmlService.createTemplateFromFile("indexhtml/IndexCreate");
  html.objectType = "Bills";
  SpreadsheetApp.getUi().showModalDialog(html.evaluate(), "Créer une nouvelle facture");
}

function createNewOrder(): void {
  let html: elmHtmlService = HtmlService.createTemplateFromFile("indexhtml/IndexCreate");
  html.objectType = "Orders";
  SpreadsheetApp.getUi().showModalDialog(html.evaluate(), "Créer une nouvelle facture");
}

function createNewAccount(): void {
  let html: elmHtmlService = HtmlService.createTemplateFromFile("indexhtml/IndexCreate");
  html.objectType = "Accounts";
  SpreadsheetApp.getUi().showModalDialog(html.evaluate(), "Créer une nouvelle facture");
}

function getCreateBillsFormData(): object {
  let billSpreadsheetValues = currentSpreadsheet.getRangeByName("Bills").getValues();
  let billParametersSpreadsheetValues: any[] = currentSpreadsheet.getRangeByName("Paramètres").getValues();
  let formData: object = {};
  formData['parameters-data'] = [];
  formData['contracts-data']= [];
  formData['bills-data'] = [];

  billParametersSpreadsheetValues.forEach(function(element) {
    formData['parameters-data'].push({
      'column-title' : element[0],
      'parameter-title' : element[1],
      'parameter-type' : element[2]
    });
  });

  let billColumnFound = false;
  let contractColumnFound = false;

  billSpreadsheetValues[0].some(function(header, headerIndex) {
    if(billColumnFound && contractColumnFound){
      return true;
    } else {
      if(header=="Contrat") {
        billSpreadsheetValues.forEach(element => {
          formData['contracts-data'].push(element[headerIndex]);
        });
        contractColumnFound = true;
      } else if (header=="Nom de la facture") {
        billSpreadsheetValues.forEach(element=> {
          formData['bills-data'].push(element[headerIndex]);
        });
        billColumnFound = true;
      }
    }
    return false;
  });

  formData['contracts-data'] = uniq(formData['contracts-data']);
  formData['bills-data'] = uniq(formData['bills-data']);
  
  return formData;
}

function loadBills() {

  let billsSpreadsheetValues = currentSpreadsheet.getRangeByName("Bills").getValues();
  let billsKeysSpreadsheetValues = currentSpreadsheet.getRangeByName("Paramètres").getValues();
  let billsData = [];
  
  billsSpreadsheetValues[0] = billsSpreadsheetValues
    .filter(function(element, index) { return index == 0; })
    .reduce(function(bout1, bout2) { return bout1.concat(bout2); })
    .map(function(header) {
      let newHeader = header;
      billsKeysSpreadsheetValues.some(function(keyPairing) {
        if(header == keyPairing[0]) { newHeader = keyPairing[1]; Logger.log(newHeader); return true; }
        return false;
      });
      return newHeader;
    });
    
    Logger.log("headers : " + billsSpreadsheetValues[0]);
  
  billsData = billsSpreadsheetValues
    .filter(function(element, index) { return index > 0; })
    .map(function(line, lineIndex) {
      let newObject = {};
      line.forEach(function(cell, cellIndex) {
        newObject[billsSpreadsheetValues[0][cellIndex].toString()] = Object.prototype.toString.call(cell) === '[object Date]' ? cell.toString() : cell;
      });
      return newObject;
    });
  
  return billsData;

}

function etablirBill(formObject) {
  
  let url = currentSpreadsheet.getRangeByName("Bill_Template").getValues()[0][0];
  let billTemplateID = getIdFrom(url);
  let billTemplateFile = DriveApp.getFileById(billTemplateID);
  let billTemplateFolder = getImmediateFolder(billTemplateFile);
  let existingBills = billTemplateFolder.getFilesByName(formObject['nom-facture']);
  let toBeFormatedAsDate = ["debut-prestation", "fin-prestation", "debut-licences", "fin-licences"];
  let toBeFormatedAsNumber = ["prix-prestation", "prix-unitaire-licences", "prix-total-licences", ];
  
  if(getFileIteratorSize(existingBills) > 0) {
    
    SpreadsheetApp.getUi().alert("La facture existe déjà : " + billTemplateFolder.getFilesByName(formObject['nom-facture']).next().getUrl());
  
  } else {
    
    let body: GoogleAppsScript.Document.Body = DocumentApp.openById(billTemplateFile.makeCopy(formObject['nom-facture']).getId()).getBody();
    let text: GoogleAppsScript.Document.Text = body.editAsText();
    let markupsBills: Object[][] = currentSpreadsheet.getRangeByName("Paramètres").getValues();

    markupsBills.forEach((element)=>{
      text.replaceText("<\\["+element[1]+"\\]>", formObject[element[1].toString()]);
    });
    
    let prixTotal = 0;
    let TVA = 0;
    let totalTTC = 0;
    
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
    
    text.replaceText("<\\[prix-total-HT\\]>", prixTotal.toString());
    text.replaceText("<\\[tva-total\\]>", TVA.toString());
    text.replaceText("<\\[total-ttc\\]>", totalTTC.toString());       

  }
}

function envoyerBill(formObject) {
  Logger.log('blob_envoyer');
  Logger.log(formObject);
}

function payerBill(formObject) {

  let billsSpreadsheetValues = currentSpreadsheet.getRangeByName("Bills").getValues();
  let billParametersSpreadsheetValues = currentSpreadsheet.getRangeByName("Paramètres").getValues();
  let billsData = [];
  let headerValues = [];

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
  
  currentSpreadsheet.getRangeByName("Bills").setValues(billReturnValues);
  
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

    let i: number = 0;

  for(i=0; i < markups.length; i++) {
  
    body.editAsText().replaceText(markups[i], "");
    
  }
}

function removeDocSection(body, startRegExp, endRegExp) {

  var totalchildren = body.getNumChildren();
  var found = 0;
  
  for (var i=totalchildren-1; i >= 0; i--) {
    
    Logger.log("body child : " + body.getChild(i));
    
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
    Logger.log("filename babilobabd : " + filename);
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getIdFrom(url): string {
  let id: string & string[];
  let parts = url.split(/^(([^:\/?#]+):)?(\/\/([^\/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/);
  if (url.indexOf('?id=') >= 0){
     id = (parts[6].split("=")[1]).replace("&usp","");
     return <string>id;
   } else {
        id = parts[5].split("/");
        //Using sort to get the id as it is the longest element. 
        let sortArr: string[] = <string[]>id.sort(function(a,b){return b.length - a.length});
        return sortArr[0];
   }
 }