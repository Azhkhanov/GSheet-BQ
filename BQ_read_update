function onOpen() { // Триггер на открытие файла
  SpreadsheetApp.getUi() 
      .createMenu('Аптека') 
      .addItem('Отобразить панель контрактов', 'showSidebar') 
      .addToUi();
}


function showSidebar() {
  var html = HtmlService.createTemplateFromFile('page')
      .evaluate()
      .setTitle('Отбор')
      .setWidth(500);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}


/**
 * Получить запрос для фильтрования по товару и контракту
 */
function getFiltersSql(med, contr) {
  // подправить запрос если требуется
  var string ='SELECT medicine_name, group_name '+
                  'FROM `ELT.detected_empty` '+
                  'WHERE '+
                      'UPPER(medicine_name) LIKE UPPER(\'%'+med+'%\') ';
  Logger.log(string);
  return string;
}

/**
 * Отправить запрос на bq и перезаписать ячейки
 */
function runQuery (med, contr) {
  var projectId = 'apteka-211807';

  var reqStr = getFiltersSql(med, contr);
  var request = {
    query: reqStr,
    useLegacySql: false
  };
  console.log(">>> request:");
  console.log(request);
  Logger.log(request);
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var jobId = queryResults.jobReference.jobId;
  Logger.log(jobId);
  // Прверка статуса Query Job.
  var sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }

  // Получаем все строки из result.
  var rows = queryResults.rows;
  while (queryResults.pageToken) {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
      pageToken: queryResults.pageToken
    });
    rows = rows.concat(queryResults.rows);
  }

  if (rows) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Distinct contracts");
    sheet.clearContents(); 
    
    //var spreadsheet = SpreadsheetApp.create('BiqQuery Results');
    //var sheet = spreadsheet.getActiveSheet();

    // Объединение заголовков.
    var headers = queryResults.schema.fields.map(function(field) {
      return field.name;
    });
    sheet.appendRow(headers);

    // Объединение результатов.
    var data = new Array(rows.length);
    for (var i = 0; i < rows.length; i++) {
      var cols = rows[i].f;
      data[i] = new Array(cols.length);
      for (var j = 0; j < cols.length; j++) {
        data[i][j] = cols[j].v;
      }
    }
    sheet.getRange(2, 1, rows.length, headers.length).setValues(data);

    Logger.log('Results spreadsheet created: %s',
        spreadsheet.getUrl());
  } else {
    Logger.log('No rows returned.');
  }
  return "Data loaded";
}

/**
 * Отправить данные на bq
 */
function pushToBQ () {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Distinct contracts");
  var projectId = 'apteka-211807';
  var datasetId = 'ELT';
  var tableId = 'test_push';
  
  // взять все значения с листа начиная с ячейки 1,1 и 
  // -1,-1 для последней строки и столбца с контентом
  var data = sheet.getSheetValues(1, 1, -1, -1);
  
  // получаем данные из таблицы для csv файла
  var csvFile = undefined;
  if (data.length > 1) {
    var csv = "";
    for (var row = 0; row < data.length; row++) {
      for (var col = 0; col < data[row].length; col++) {
        if (data[row][col].toString().indexOf(",") != -1) {
          data[row][col] = "\"" + data[row][col] + "\"";
        }
      }
      // join each row's columns
      // add a carriage return to end of each row, except for the last one
      if (row < data.length-1) {
        csv += data[row].join(",") + "\r\n";
      }
      else {
        csv += data[row];
      }
    }
    csvFile = csv;
  }
  // создаем csv файл с полученными данными 
  var csv_name = 'temp_' + new Date().getTime()+'.csv';
  Logger.log("csv name: " + csv_name);
  DriveApp.createFile(csv_name, csvFile);
  Logger.log("created file in drive app ");
  
  // обработка файла
  var files = DriveApp.getFilesByName(csv_name);
  while (files.hasNext()) {
    var file = files.next(); 
    Logger.log("processing file: " + csv_name);
    var table = {
      tableReference: {
        projectId: projectId,
        datasetId: datasetId,
        tableId: tableId
      },
    };
    // берем данные из csv
    var data = file.getBlob().setContentType('application/octet-stream'); 
    // создаем объект job_data с конфигурацией на загрузку данных 
    var job_data = {
      configuration: {
        load: {
          destinationTable: {
            projectId: projectId,
            datasetId: datasetId,
            tableId: tableId
          },
          // пропускаем заголовки
          skipLeadingRows: 1,
          autodetect: false,
          // write_append для добавления строк
          writeDisposition: 'WRITE_APPEND'
        }
      }
    };
    // создаем и запускаем job 
    var job = null;
    try {
      job = BigQuery.Jobs.insert(job_data, projectId, data);
    }
    catch(e) {
      Logger.log(e);
      console.log(e);
      throw e;
    }
    Logger.log("job description: " + job);
    // помечаем файл на удаление
    file.setTrashed(true);
  }
  return "Данные отправлены";
}
