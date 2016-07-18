var FIRST_ROW = 1;
var FIRST_COLUMN = 1;
var LATITUDE_COLUMN = 2
var LONGITUDE_COLUMN = 3


function onOpen(){
var ss = SpreadsheetApp.getActiveSpreadsheet();

var subMenus = [
  {name : "Convert coordinates (DMS do DD)", functionName : "convertDMStoDD"}
]; 
  
  ss.addMenu("Converter", subMenus);
}


function convertDMStoDD() {
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spread.getActiveSheet();
  
  var coordinates = getCoordinates_(sheet);
  var convertedLatitude = convertLatitude_(coordinates);
  var latitudeRange = getLatitudeRange_(sheet);
  writeConvertedCoordinates(latitudeRange, convertedLatitude);

  var coordinates = getCoordinates_(sheet);
  var convertedLongitude = convertLongitude_(coordinates);
  var longitudeRange = getLongitudeRange_(sheet);
  writeConvertedCoordinates(longitudeRange, convertedLongitude);
 
}

function getCoordinates_(sheet){
  var coordinatesRange = sheet.getRange(FIRST_ROW, FIRST_COLUMN, sheet.getDataRange().getNumRows());
  var coordinates = coordinatesRange.getValues(); 
  return coordinates;
}

function convertLatitude_(coordinates){
  
  for(var i = 0; i<coordinates.length; ++i){
    var coordinatesToConvert = coordinates[i][0].toString();
    var latitudeToConvert = getLatitude_(coordinatesToConvert);
   
    var latitudeDD = convertToDD_(latitudeToConvert);    
    coordinates[i][0] = latitudeDD;
  }
  
  return coordinates;
}

function convertLongitude_(coordinates){
  
  for(var i = 0; i<coordinates.length; ++i){
    var coordinatesToConvert = coordinates[i][0].toString();
    var longitudeToConvert = getLongitude_(coordinatesToConvert);
    var longitudeDD = convertToDD_(longitudeToConvert);
    
    coordinates[i][0] = longitudeDD;
  }
  
  return coordinates;
}


function convertToDD_(coordinateToConvert){
  var degrees = getDegrees_(coordinateToConvert);
  var minutes = getMinutes_(coordinateToConvert);
  var seconds = getSeconds_(coordinateToConvert);
  
  var coordinateDD = parseFloat(degrees.replace(",", ".")) + parseFloat(minutes.replace(",", "."))/60 + parseFloat(seconds.replace(",", "."))/3600
  
  return coordinateDD.toFixed(6);
}

function getDegrees_(coordinateToConvert){
  var degreesIdx = coordinateToConvert.indexOf("°");
  var degrees = coordinateToConvert.substring(0, degreesIdx);
  
  return degrees;
}

function getMinutes_(coordinateToConvert){
  var degreesIdx = coordinateToConvert.indexOf("°");
  var minutesIdx = coordinateToConvert.indexOf("\'");
  var minutes = coordinateToConvert.substring(degreesIdx+1, minutesIdx);
  
  return minutes;
}

function getSeconds_(coordinateToConvert){
  
  var minutesIdx = coordinateToConvert.indexOf("\'");
  var secondsIdx = coordinateToConvert.indexOf("\"");
  var seconds = coordinateToConvert.substring(minutesIdx+1, secondsIdx);
  
  return seconds;
}

function getLatitude_(coordinatesToConvert){
  var spaceBetweenLatitudeAndLongitudeIdx = coordinatesToConvert.indexOf(" ");
  var latitude = coordinatesToConvert.substring(0, spaceBetweenLatitudeAndLongitudeIdx);
  return latitude;
}

function getLongitude_(coordinatesToConvert){
  var spaceBetweenLatitudeAndLongitudeIdx = coordinatesToConvert.indexOf(" ");
  var longitude = coordinatesToConvert.substring(spaceBetweenLatitudeAndLongitudeIdx+1);
  return longitude;
}


function getLatitudeRange_(sheet){
  return sheet.getRange(FIRST_ROW, LATITUDE_COLUMN, sheet.getDataRange().getNumRows());
}

function getLongitudeRange_(sheet){
  return sheet.getRange(FIRST_ROW, LONGITUDE_COLUMN, sheet.getDataRange().getNumRows());
}

function writeConvertedCoordinates(coordinatesRange, convertedCoordinates){
  coordinatesRange.setValues(convertedCoordinates);
}

