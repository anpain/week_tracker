const WEEEK_API_BASE_URL = "https://api.weeek.net/public/v1/";
const WEEEK_TOKEN_KEY = "WEEEK_API_TOKEN"; 
const GOOGLE_SHEET_ID_KEY = "GOOGLE_SHEET_ID"; 

const SHEET_NAME = "Отчет WEEEK"; 

const PRIORITY_MAP = {0: "Низкий", 1: "Средний", 2: "Высокий", 3: "Замороженный"};
const PRIORITY_COLORS_GSHEETS = {
  "Замороженный": "#5CB8FF", "Высокий": "#FFA500", "Средний": "#FFFF66", "Низкий": "#B3D980"
};

const HEADER_BACKGROUND_COLOR = "#4A568D"; 
const HEADER_FONT_COLOR = "#FFFFFF";       
const FONT = "Roboto";      

const STATUS_COLORS_GSHEETS = {
  "Бэклог": "#73CCFF", "В работе": "#793AFF", "Тестирование": "#FFCC48", 
  "Демо": "#93CE48", "Тестирование пользователем": "#FFCC48", "Завершено": "#3AC648"
};
const OTHER_STATUS_BACKGROUND_COLOR = "#F0F0F0";

function getWeeekApiToken_() {
  const token = PropertiesService.getScriptProperties().getProperty(WEEEK_TOKEN_KEY);
  if (!token) {
    const errorMessage = `WEEEK API токен не найден. Запустите 'Настроить конфигурацию' из меню.`;
    Logger.log(errorMessage);
    throw new Error("WEEEK API токен не настроен.");
  }
  return token;
}

function getGoogleSheetId_() {
  const sheetId = PropertiesService.getScriptProperties().getProperty(GOOGLE_SHEET_ID_KEY);
  if (!sheetId) {
    const errorMessage = `Google Sheet ID не найден. Запустите 'Настроить конфигурацию' из меню.`;
    Logger.log(errorMessage);
    throw new Error("Google Sheet ID не настроен.");
  }
  return sheetId;
}

function makeApiRequest_(endpoint, apiToken, queryParameters) {
  const requestOptions = {
    'method': 'get', 
    'headers': {'Authorization': `Bearer ${apiToken}`}, 
    'muteHttpExceptions': true
  };

  let targetUrl = WEEEK_API_BASE_URL + endpoint;
  if (queryParameters) {
    const queryString = Object.keys(queryParameters)
      .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(queryParameters[key])}`)
      .join('&');
    if (queryString) targetUrl += (targetUrl.includes('?') ? '&' : '?') + queryString;
  }
  
  try {
    const httpResponse = UrlFetchApp.fetch(targetUrl, requestOptions);
    const responseCode = httpResponse.getResponseCode();
    const responseText = httpResponse.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      try { return JSON.parse(responseText); } 
      catch (jsonError) { Logger.log(`Ошибка парсинга JSON от ${targetUrl}: ${jsonError}. Ответ: ${responseText.substring(0, 200)}`); return null; }
    } else { 
      Logger.log(`Ошибка HTTP для ${targetUrl}: Код ${responseCode}. Ответ: ${responseText.substring(0, 200)}`); return null; 
    }
  } catch (fetchError) { 
    Logger.log(`Исключение при запросе к ${targetUrl}: ${fetchError}`); return null; 
  }
}

function fetchWorkspaceMembers_(apiToken) {
  const apiResponse = makeApiRequest_("ws/members", apiToken);
  const membersMap = {};
  if (apiResponse && apiResponse.success && apiResponse.members) {
    apiResponse.members.forEach(member => { if (member.id) membersMap[member.id] = member.firstName || member.id; });
  } else { 
    Logger.log("Не удалось загрузить участников рабочей области."); 
  }
  return membersMap;
}

function fetchBoardColumnsForBoard_(boardId, apiToken) {
  const apiResponse = makeApiRequest_(`tm/board-columns?boardId=${boardId}`, apiToken);
  if (apiResponse && apiResponse.success && apiResponse.boardColumns) return apiResponse.boardColumns;
  Logger.log(`Не удалось загрузить колонки для доски ID: ${boardId}.`);
  return [];
}

function fetchAllTasks_(apiToken) {
  const allTasks = []; 
  let apiCursor = null; 
  let pageCounter = 0; 
  let hasMoreTasks = true;

  while (hasMoreTasks) {
    pageCounter++;
    const requestParams = apiCursor ? { "cursor": apiCursor } : {};
    const apiResponse = makeApiRequest_("tm/tasks", apiToken, requestParams);

    if (apiResponse && apiResponse.success && apiResponse.tasks) {
      allTasks.push(...apiResponse.tasks);
      if (apiResponse.hasMore && apiResponse.cursor) apiCursor = apiResponse.cursor; 
      else hasMoreTasks = false;
    } else { 
      Logger.log(`Ошибка при загрузке ${pageCounter}-й страницы задач.`); 
      hasMoreTasks = false; 
    }
  }
  return allTasks;
}

function fetchWeeekData_(apiToken) {
  Logger.log("Загрузка данных из WEEEK...");
  const workspaceMembers = fetchWorkspaceMembers_(apiToken);
  const allTasksData = fetchAllTasks_(apiToken);
  const boardColumnsMap = {};

  if (allTasksData && allTasksData.length > 0) {
    const uniqueBoardIdsRegistry = {}; 
    allTasksData.forEach(task => { if(task.boardId != null) uniqueBoardIdsRegistry[task.boardId] = true; });
    
    Object.keys(uniqueBoardIdsRegistry).forEach(boardIdString => {
      const boardIdNumber = parseInt(boardIdString, 10);
      if (!isNaN(boardIdNumber)) {
        fetchBoardColumnsForBoard_(boardIdNumber, apiToken).forEach(column => {
          if (column.id != null) {
            const columnIdNumber = parseInt(column.id, 10);
            if(!isNaN(columnIdNumber)) boardColumnsMap[`${boardIdNumber}_${columnIdNumber}`] = column.name || String(columnIdNumber);
          }
        });
      }
    });
  }
  return { 
    membersMap: workspaceMembers, 
    tasksData: allTasksData, 
    boardColumnNamesMap: boardColumnsMap 
  };
}

function parseDateStringSafe_(dateString) {
  if (!dateString) return null;
  try { return new Date(dateString); } 
  catch (error) { Logger.log(`Не удалось преобразовать '${dateString}' в дату: ${error}`); return null; }
}

function addMonthsToDate_(initialDate, numberOfMonths) {
  const newDate = new Date(initialDate); 
  newDate.setMonth(newDate.getMonth() + numberOfMonths);
  if (newDate.getDate() < initialDate.getDate() && numberOfMonths > 0) newDate.setDate(0); 
  return newDate;
}

function getReportingPeriods_(currentDate) {
  const dayOfMonth = currentDate.getDate(); 
  const period1StartDay = 6, period1EndDay = 20;
  const period2StartDay = 21, period2EndDay = 5;
  
  let period1BaseDate = new Date(currentDate); 
  if (dayOfMonth <= 5) period1BaseDate = addMonthsToDate_(period1BaseDate, -1);
  
  const p1StartDate = new Date(period1BaseDate.getFullYear(), period1BaseDate.getMonth(), period1StartDay);
  const p1EndDate = new Date(period1BaseDate.getFullYear(), period1BaseDate.getMonth(), period1EndDay);
  
  let period2BaseDate = new Date(period1BaseDate);
  const p2StartDate = new Date(period2BaseDate.getFullYear(), period2BaseDate.getMonth(), period2StartDay);
  
  let period2EndMonthBase = addMonthsToDate_(new Date(period2BaseDate), 1);
  const p2EndDate = new Date(period2EndMonthBase.getFullYear(), period2EndMonthBase.getMonth(), period2EndDay);
  
  return [[p1StartDate, p1EndDate], [p2StartDate, p2EndDate]];
}

function formatValueForSheetCell_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "dd-MM-yyyy");
  if (value === null || typeof value === 'undefined') return "";
  return value;
}

function convertColumnIndexToLetter_(columnIndexOneBased) {
  let temp, letter = ''; 
  while (columnIndexOneBased > 0) { 
    temp = (columnIndexOneBased - 1) % 26; 
    letter = String.fromCharCode(temp + 65) + letter; 
    columnIndexOneBased = (columnIndexOneBased - temp - 1) / 26; 
  } 
  return letter;
}

function processTasksToSheetRows_(
    allTasks, 
    workspaceMembersMap, 
    boardColumnsNameMap, 
    reportingPeriod1Dates, 
    reportingPeriod2Dates
  ) {
  const sheetRows = []; 
  const [period1StartDate, period1EndDate] = reportingPeriod1Dates; 
  const [period2StartDate, period2EndDate] = reportingPeriod2Dates;

  allTasks.forEach(taskItem => {
    const priorityValue = taskItem.priority;
    const priorityString = priorityValue != null ? (PRIORITY_MAP[priorityValue] || String(priorityValue)) : "";
    
    const userId = taskItem.userId;
    const executorName = userId ? (workspaceMembersMap[userId] || userId) : "";
    
    let statusName = "";
    const taskBoardId = taskItem.boardId != null ? parseInt(taskItem.boardId, 10) : null;
    const taskColumnId = taskItem.boardColumnId != null ? parseInt(taskItem.boardColumnId, 10) : null;

    if (taskBoardId != null && !isNaN(taskBoardId) && taskColumnId != null && !isNaN(taskColumnId)) {
      statusName = boardColumnsNameMap[`${taskBoardId}_${taskColumnId}`] || String(taskColumnId);
    } else if (taskColumnId != null && !isNaN(taskColumnId)) {
      statusName = String(taskColumnId);
    } else if (taskItem.boardColumnId != null) {
      statusName = String(taskItem.boardColumnId);
    }
    
    const creationDate = parseDateStringSafe_(taskItem.createdAt); 
    let completionDate = null;
    if (taskItem.isCompleted) {
      const timeEntries = (taskItem.timeEntries || taskItem.workloads || []).filter(entry => entry.date);
      if (timeEntries.length > 0) {
        completionDate = parseDateStringSafe_(timeEntries.sort((a, b) => new Date(b.date) - new Date(a.date))[0].date);
      } else if (taskItem.updatedAt) {
        completionDate = parseDateStringSafe_(taskItem.updatedAt);
      }
    }
    
    const estimatedHours = (taskItem.duration && taskItem.duration > 0) ? Math.ceil(taskItem.duration / 60) : null;
    
    let timePeriod1Minutes = 0, timePeriod2Minutes = 0; 
    const commentsPeriod1 = [], commentsPeriod2 = [];
    (taskItem.workloads || []).forEach(workloadEntry => {
      const workloadDate = parseDateStringSafe_(workloadEntry.date);
      const durationMinutes = workloadEntry.duration || 0;
      const commentText = workloadEntry.comment || "";
      if (workloadDate && durationMinutes > 0) {
        if (workloadDate >= period1StartDate && workloadDate <= period1EndDate) { 
          timePeriod1Minutes += durationMinutes; 
          if (commentText) commentsPeriod1.push(commentText); 
        }
        if (workloadDate >= period2StartDate && workloadDate <= period2EndDate) { 
          timePeriod2Minutes += durationMinutes; 
          if (commentText) commentsPeriod2.push(commentText); 
        }
      }
    });
    
    const commentsPeriod1String = commentsPeriod1.map((comment, index) => `${index + 1}. ${comment}`).join("\n");
    const commentsPeriod2String = commentsPeriod2.map((comment, index) => `${index + 1}. ${comment}`).join("\n");
    const timePeriod1Hours = timePeriod1Minutes > 0 ? Math.ceil(timePeriod1Minutes / 60) : null;
    const timePeriod2Hours = timePeriod2Minutes > 0 ? Math.ceil(timePeriod2Minutes / 60) : null;
    
    sheetRows.push([
      `${taskItem.title || ""} (${taskItem.id || ""})`, priorityString, executorName, statusName,
      creationDate, completionDate, estimatedHours, timePeriod1Hours, commentsPeriod1String, 
      timePeriod2Hours, commentsPeriod2String,
      null, null 
    ].map(formatValueForSheetCell_));
  });
  return sheetRows;
}

function getOrCreateSheet_(spreadsheetObject, sheetNameForReport) {
  let targetSheet = spreadsheetObject.getSheetByName(sheetNameForReport);
  if (!targetSheet) {
    targetSheet = spreadsheetObject.insertSheet(sheetNameForReport);
    Logger.log(`Лист '${sheetNameForReport}' создан.`);
  } else {
    Logger.log(`Лист '${sheetNameForReport}' найден. Очистка содержимого и форматов...`);
    targetSheet.clearContents().clearFormats(); 
  }
  return targetSheet;
}

function updateGoogleSheet_(
    googleSheetFileId, 
    targetSheetNameForReport, 
    reportHeadersArray, 
    reportDataRows
  ) {
  Logger.log(`Обновление Google Таблицы ID: ${googleSheetFileId}, Лист: ${targetSheetNameForReport}`);
  try {
    const spreadsheet = SpreadsheetApp.openById(googleSheetFileId);
    const worksheet = getOrCreateSheet_(spreadsheet, targetSheetNameForReport);
    
    const dataToWriteToSheet = [reportHeadersArray, ...reportDataRows];
    
    let trackedTime1ColIdx = -1, trackedTime2ColIdx = -1, hourlyRateColIdx = -1, salaryCalcColIdx = -1;
    let priorityColIdx = -1, statusColIdx = -1; 

    reportHeadersArray.forEach((headerText, index) => {
      if (headerText.startsWith("Затрекано (") && trackedTime1ColIdx === -1) trackedTime1ColIdx = index;
      else if (headerText.startsWith("Затрекано (")) trackedTime2ColIdx = index;
      if (headerText === "Стоимость в час") hourlyRateColIdx = index;
      if (headerText === "Расчет зп") salaryCalcColIdx = index;
      if (headerText === "Приоритет") priorityColIdx = index;
      if (headerText === "Статус") statusColIdx = index;
    });
    const canCalculateSalary = trackedTime1ColIdx !== -1 && trackedTime2ColIdx !== -1 && 
                             hourlyRateColIdx !== -1 && salaryCalcColIdx !== -1;
    if (!canCalculateSalary) Logger.log("Не все колонки для формулы расчета ЗП найдены. Формула не будет применена.");

    for (let rowIndex = 1; rowIndex < dataToWriteToSheet.length; rowIndex++) {
        const sheetRowNumber = rowIndex + 1;
        if (canCalculateSalary) {
            const colHLetter = convertColumnIndexToLetter_(trackedTime1ColIdx + 1);
            const colJLetter = convertColumnIndexToLetter_(trackedTime2ColIdx + 1);
            const colLLetter = convertColumnIndexToLetter_(hourlyRateColIdx + 1);
            dataToWriteToSheet[rowIndex][salaryCalcColIdx] = 
                `= (N(${colHLetter}${sheetRowNumber}) + N(${colJLetter}${sheetRowNumber})) * N(${colLLetter}${sheetRowNumber})`;
        }
    }
    
    if (dataToWriteToSheet.length > 0) {
      const fullRange = worksheet.getRange(1, 1, dataToWriteToSheet.length, reportHeadersArray.length);
      fullRange.setValues(dataToWriteToSheet)
               .setFontFamily(FONT)
               .setFontSize(12)
               .setVerticalAlignment("middle")
               .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    }

    worksheet.getRange(1, 1, 1, reportHeadersArray.length)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground(HEADER_BACKGROUND_COLOR)
      .setFontColor(HEADER_FONT_COLOR);

    if (reportDataRows.length > 0) {
      reportHeadersArray.forEach((headerName, columnIndexZeroBased) => {
        const columnDataRange = worksheet.getRange(2, columnIndexZeroBased + 1, reportDataRows.length, 1);
        
        if (headerName === "Задача (Номер)" || headerName.startsWith("Комментарии")) {
          columnDataRange.setHorizontalAlignment("left");
        } else {
          columnDataRange.setHorizontalAlignment("center");
        }

        if (headerName === "Приоритет" && priorityColIdx !== -1) {
          reportDataRows.forEach((dataRowArray, rowIndexZeroBased) => {
            const priorityText = String(dataRowArray[priorityColIdx]);
            const backgroundColorHex = PRIORITY_COLORS_GSHEETS[priorityText];
            if (backgroundColorHex) {
              worksheet.getRange(rowIndexZeroBased + 2, priorityColIdx + 1).setBackground(backgroundColorHex);
            }
          });
        }

        if (headerName === "Статус" && statusColIdx !== -1) {
          reportDataRows.forEach((dataRowArray, rowIndexZeroBased) => {
            const statusText = String(dataRowArray[statusColIdx]);
            const backgroundColorHex = STATUS_COLORS_GSHEETS[statusText] || OTHER_STATUS_BACKGROUND_COLOR;
            if (backgroundColorHex) {
               worksheet.getRange(rowIndexZeroBased + 2, statusColIdx + 1).setBackground(backgroundColorHex);
            }
          });
        }
        
        if (headerName === "Расчет зп") columnDataRange.setNumberFormat("0.00");
        else if (headerName === "Оценка времени" || headerName.startsWith("Затрекано (")) columnDataRange.setNumberFormat("0");
      });
    }
    
    if (reportHeadersArray.length > 0) {
        for (let colIndex = 0; colIndex < reportHeadersArray.length; colIndex++) {
            worksheet.autoResizeColumn(colIndex + 1); 
        }
        
        const minWidths = {
          "Приоритет": 110,
          "Исполнитель": 130,
          "Статус": 100,
          "Дата начала": 120,
          "Дата окончания": 120,
          "Комментарии": 200};
          
        const dateHeaderMinWidth = 150; 

        reportHeadersArray.forEach((headerName, colIndexZeroBased) => {
            const currentColumnWidth = worksheet.getColumnWidth(colIndexZeroBased + 1);
            let newWidth = -1;

            if (minWidths[headerName] && currentColumnWidth < minWidths[headerName]) {
                newWidth = minWidths[headerName];
            } else if (headerName.startsWith("Затрекано (") && currentColumnWidth < dateHeaderMinWidth) {
                newWidth = dateHeaderMinWidth;
            }

            if (newWidth > 0) {
                 worksheet.setColumnWidth(colIndexZeroBased + 1, newWidth);
            }
        });
    }
    Logger.log("Данные успешно записаны и отформатированы в Google Таблице.");

  } catch (error) { 
    Logger.log(`Ошибка при работе с Google Sheets: ${error}. Stack: ${error.stack}`); 
  }
}

function runReportGeneration() {
  let finalLogMessage = ""; 
  let operationSucceeded = false;

  try {
    Logger.log("Запуск генерации отчета (логика).");
    const weeekApiToken = getWeeekApiToken_(); 
    const googleSheetFileId = getGoogleSheetId_(); 
    
    const currentDate = new Date();
    const [reportingPeriod1, reportingPeriod2] = getReportingPeriods_(currentDate);
    const [p1StartDate, p1EndDate] = reportingPeriod1; 
    const [p2StartDate, p2EndDate] = reportingPeriod2; 
    
    Logger.log(`Период 1: ${Utilities.formatDate(p1StartDate, Session.getScriptTimeZone(), "dd-MM-yyyy")} - ${Utilities.formatDate(p1EndDate, Session.getScriptTimeZone(), "dd-MM-yyyy")}`);
    Logger.log(`Период 2: ${Utilities.formatDate(p2StartDate, Session.getScriptTimeZone(), "dd-MM-yyyy")} - ${Utilities.formatDate(p2EndDate, Session.getScriptTimeZone(), "dd-MM-yyyy")}`);
    
    const periodStringForSheetName = 
        `${Utilities.formatDate(p1StartDate, Session.getScriptTimeZone(), "dd.MM")}` +
        `-${Utilities.formatDate(p1EndDate, Session.getScriptTimeZone(), "dd.MM")}`;
    let dynamicTargetSheetName = `${SHEET_NAME} (${periodStringForSheetName})`;
    dynamicTargetSheetName = dynamicTargetSheetName.replace(/[\[\]\*\/\\\?\:]/g, "").substring(0, 95);
    Logger.log(`Целевой лист: '${dynamicTargetSheetName}'`);
    
    const { membersMap, tasksData, boardColumnNamesMap } = fetchWeeekData_(weeekApiToken);

    if (!tasksData || tasksData.length === 0) {
      finalLogMessage = "Задачи не загружены. Отчет не будет сформирован.";
      Logger.log(finalLogMessage);
    } else {
      const sheetRowsData = processTasksToSheetRows_(
        tasksData, membersMap, boardColumnNamesMap, reportingPeriod1, reportingPeriod2
      );
      const reportHeaders = [
        "Задача (Номер)", "Приоритет", "Исполнитель", "Статус",
        "Дата начала", "Дата окончания", "Оценка времени",
        `Затрекано (${Utilities.formatDate(p1StartDate, Session.getScriptTimeZone(), "dd.MM")}-${Utilities.formatDate(p1EndDate, Session.getScriptTimeZone(), "dd.MM")})`, 
        "Комментарии", 
        `Затрекано (${Utilities.formatDate(p2StartDate, Session.getScriptTimeZone(), "dd.MM")}-${Utilities.formatDate(p2EndDate, Session.getScriptTimeZone(), "dd.MM")})`, 
        "Комментарии", 
        "Стоимость в час", "Расчет зп"
      ];
      updateGoogleSheet_(googleSheetFileId, dynamicTargetSheetName, reportHeaders, sheetRowsData);
      finalLogMessage = `Отчет WEEEK ('${dynamicTargetSheetName}') успешно сгенерирован и обновлен!`;
      operationSucceeded = true;
      Logger.log(finalLogMessage);
    }

  } catch (error) {
    if (!finalLogMessage) { 
        finalLogMessage = `Произошла ошибка: ${error.message}.`;
    }
    Logger.log(`Критическая ошибка при генерации отчета: ${error.message}. Stack: ${error.stack ? error.stack : 'Стек недоступен'}`);
  } finally {
    if (operationSucceeded) {
      Logger.log(`Завершение операции: Успех. Сообщение: ${finalLogMessage}`);
    } else {
      Logger.log(`Завершение операции: Ошибка. Сообщение: ${finalLogMessage || 'Произошла неизвестная ошибка.'}`);
    }
  }
}
