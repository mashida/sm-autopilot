//Это глобальные переменные, их можно использовать внутри других функций
var startDT = Date.now(); //Записали текущую дату и время по UNIX в миллисекундах
var report = {}; //Создали объект, в который запишем отчёт

//+------------------------------------------------------------------+
//| Принимаем запрос из Автопилота                                   |
//+------------------------------------------------------------------+
function doPost(e){
    //Читаем параметры запроса и пишем их в одноимённые переменные
    var SpreadsheetID = e.parameter.SpreadsheetID; //Записали в переменную ID гугл-таблицы
    var SheetName = e.parameter.SheetName; //Записали название нужного листа
    var Col = e.parameter.Col; //Записали номер столбца
    var Row = e.parameter.Row; //Записали номер строки

    var jsonString = e.parameter.Str.replace(/'/g, '"'); // Конвертируем строку в валидный JSON объект
    var Arr = parseResponseTo2DArray(jsonString); // Парсим JSON объект в двумерный массив

    //Открываем таблицу и записываем в неё данные
    var SS = SpreadsheetApp.openById(SpreadsheetID); //Открываем таблицу по ID
    var ST = SS.setActiveSheet(SS.getSheetByName(SheetName)); //Вызываем лист по названию из таблицы, которая в переменной Spreadsheet
    var RowIndex = Row; //Получили номер новой строки, которая будет после последней заполненной
    ST.getRange(RowIndex, Col, Arr.length, Arr[0].length).setValues(Arr) // Вставляем двумерный массив

    //Формируем объект-отчёт
    report["runtime"] = (Date.now() - startDT)/1000; //Расчитали время выполнения Гугл-скрипта в секундах
    report["success"] = "1";
    report["rowline"] = RowIndex;

    //Завершаем работу скрипта и отправляем ответ
    SpreadsheetApp.flush();
    
    return ContentService.createTextOutput(JSON.stringify(report));
}

//+------------------------------------------------------------------+
//| парсим JSON в двумерный массив                                   |
//+------------------------------------------------------------------+
function parseResponseTo2DArray(responseObject) {  
  var amountOfresponses = responseObject['response'].length;
  var response = responseObject['response'];
  var result = new Array(amountOfresponses);
 
  for (var counter = 0; counter < amountOfresponses; counter = counter + 1) {
    // здесь определяем сколько будет параметров в массиве и их значения по умолчанию
    result[counter] = ["", "", "", 0, 0, "", "", "", 0, "", "", ""]; 

    // добавляем дату
    result[counter][0] = Utilities.formatDate(new Date(), "GMT+3", "dd.MM.yyyy");
    // добавляем название
    result[counter][1] = "";

    // проверяем есть ли в stats сколько-нибудь данных. Бывает ни сколько.
    if (response[counter].stats.length > 0) {
      var stats = response[counter].stats[0];
      
      // добавляем потрачено    
      result[counter][2] = getValueOfTheObject(stats, "spent", "");
      // добавялем показы
      result[counter][3] = getValueOfTheObject(stats, "impressions", 0);
      // добавляем клики
      result[counter][4] = getValueOfTheObject(stats, "clicks", 0);
      // добавляем ctr
      result[counter][5] = getValueOfTheObject(stats, "ctr", "");
      // добавляем ecpc
      result[counter][6] = getValueOfTheObject(stats, "ecpc", "");
      // добавляем ecpm
      result[counter][7] = getValueOfTheObject(stats, "ecpm", "");
      // добавляем ID объявления
      result[counter][8] = getValueOfTheObject(response[counter], "id", 0);
    }

    // добавляем дата и время создания объявления
    result[counter][9] = "";
    // добавляем кабинет клиента
    result[counter][10] = "";
    // добавляем ID компании
    result[counter][11] = "";
  }
  return result;
}

//+------------------------------------------------------------------+
//| получаем значение параметра key объекта object                   |
//+------------------------------------------------------------------+
function getValueOfTheObject(obj, key, defaultValue){
  return obj.hasOwnProperty(key) ? obj[key] : defaultValue;
}