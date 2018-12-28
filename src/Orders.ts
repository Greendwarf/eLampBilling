import * as accounts from "./Accounts";
import { currentSpreadsheet as currentSpreadsheet } from "./Code";

function getOrderData() {
    let orderData = {};
    let billSpreadsheetValues = currentSpreadsheet.getRangeByName("Bills").getValues();
    let orderSpreadsheetValues = currentSpreadsheet.getRangeByName("Orders").getValues();
    let billData = [];
    let orderDataFromSpreadsheet = [];

    billData = billSpreadsheetValues
    .filter(function(element, index) { return index > 0; })
    .map(function(line, lineIndex) {
        let newObject = {};
        line.forEach(function(cell, cellIndex) {
        newObject[billSpreadsheetValues[0][cellIndex].toString()] = Object.prototype.toString.call(cell) === '[object Date]' ? cell.toString() : cell;
        });
        return newObject;
    });

    orderDataFromSpreadsheet = orderSpreadsheetValues
    .filter(function(element, index) { return index > 0; })
    .map(function(line, lineIndex) {
        let newObject = {};
        line.forEach(function(cell, cellIndex) {
        newObject[orderSpreadsheetValues[0][cellIndex].toString()] = Object.prototype.toString.call(cell) === '[object Date]' ? cell.toString() : cell;
        });
        return newObject;
    });

    orderData['bill-data'] = billData;
    orderData['order-data'] = orderDataFromSpreadsheet;

    return orderData;
}