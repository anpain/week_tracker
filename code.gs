// КОНФИГИ
const WEEEK_API_BASE_URL = "https://api.weeek.net/public/v1/";
const WEEEK_TOKEN_KEY = "WEEEK_API_TOKEN"; 
const GOOGLE_SHEET_ID_KEY = "GOOGLE_SHEET_ID"; 
const GENERAL_REPORT_SHEET_NAME = "Общий отчет";

const EXCLUDED_EXECUTORS = ["Андрей", "Анастасия"]; 

const PRIORITY_MAP = {0: "Низкий", 1: "Средний", 2: "Высокий", 3: "Замороженный"};
const PRIORITY_COLORS_GSHEETS = {
  "Замороженный": "#AED6F1", "Высокий": "#b10202", "Средний": "#ffe5a0", "Низкий": "#11734b"      
};
const HEADER_BACKGROUND_COLOR = "#34495E"; 
const HEADER_FONT_COLOR = "#FFFFFF";       
const FONT_FAMILY = "Roboto"; 
const STATUS_COLORS_GSHEETS = {
  "Бэклог": "#FFD780", "В работе": "#5a3286", "Тестирование": "#ffe5a0", 
  "Демо": "#d4edbc", "Тестирование пользователем": "#ffe5a0", "Тестирование пользователями": "#ffe5a0", 
  "Завершено": "#11734b", "Ждут выгрузки на прод": "#FBA0E3", "Заблокировано": "#0a53a8",
  "Баги": "#b10202", "Админ.задачи": "#215a6c"
};
const OTHER_STATUS_BACKGROUND_COLOR = "#F2F3F4"; 
const PROJECT_COLORS_GSHEETS = {
  "Система учета": "#0039aa", "Техническая поддержка": "#b83018",
  "Портал для сотрудников": "#8d5bc2", "Телеграм-боты": "#f7e0b5"
};
const OTHER_PROJECT_BACKGROUND_COLOR = null; 
const EXECUTOR_COLORS_GSHEETS = {
  "Вячеслав": "#c6dbe1", "Игорь": "#e6cff2", "Роман": "#d4edbc"
};
const OTHER_EXECUTOR_BACKGROUND_COLOR = null; 
const ROW_ERROR_BACKGROUND_COLOR = "#FADBD8"; 
const ESTIMATE_MISSING_BACKGROUND_COLOR = "#FADBD8"; 
const TEXT_COLOR_LIGHT = "#FFFFFF"; 
const TEXT_COLOR_DARK = "#000000";  
 
function getWeeekApiToken_() {
  const token = PropertiesService.getScriptProperties().getProperty(WEEEK_TOKEN_KEY);
  if (!token) {
    const errorMessage = `WEEEK API токен не найден в Свойствах Скрипта (ключ '${WEEEK_TOKEN_KEY}'). Установите его вручную.`;
    Logger.log(errorMessage);
    throw new Error(errorMessage); 
  }
  return token;
}

function getGoogleSheetId_() {
  const sheetId = PropertiesService.getScriptProperties().getProperty(GOOGLE_SHEET_ID_KEY);
  if (!sheetId) {
    const errorMessage = `Google Sheet ID не найден в Свойствах Скрипта (ключ '${GOOGLE_SHEET_ID_KEY}'). Установите его вручную.`;
    Logger.log(errorMessage);
    throw new Error(errorMessage); 
  }
  return sheetId;
}

// ПАРСИНГ WEEEK API
function makeApiRequest_(endpoint, apiToken, queryParameters) {
  const requestOptions = {'method': 'get', 'headers': {'Authorization': `Bearer ${apiToken}`}, 'muteHttpExceptions': true};
  let targetUrl = WEEEK_API_BASE_URL + endpoint;
  if (queryParameters) {
    const queryString = Object.keys(queryParameters).map(key => `${encodeURIComponent(key)}=${encodeURIComponent(queryParameters[key])}`).join('&');
    if (queryString) targetUrl += (targetUrl.includes('?') ? '&' : '?') + queryString;
  }
  try {
    const httpResponse = UrlFetchApp.fetch(targetUrl, requestOptions);
    const responseCode = httpResponse.getResponseCode();
    const responseText = httpResponse.getContentText();
    if (responseCode >= 200 && responseCode < 300) {
      try { return JSON.parse(responseText); } 
      catch (jsonError) { Logger.log(`Ошибка парсинга JSON от ${targetUrl}: ${jsonError}. Ответ: ${responseText.substring(0, 200)}`); return null; }
    } else { Logger.log(`Ошибка HTTP для ${targetUrl}: Код ${responseCode}. Ответ: ${responseText.substring(0, 200)}`); return null; }
  } catch (fetchError) { Logger.log(`Исключение при запросе к ${targetUrl}: ${fetchError}`); return null; }
}

function fetchWorkspaceMembers_(apiToken) {
  const apiResponse = makeApiRequest_("ws/members", apiToken);
  const membersMap = {};
  if (apiResponse && apiResponse.success && Array.isArray(apiResponse.members)) {
    apiResponse.members.forEach(member => { if (member.id) membersMap[member.id] = member.firstName || member.id; });
  } else { Logger.log("Не удалось загрузить участников рабочей области или ответ не содержит массив 'members'."); }
  return membersMap;
}

function fetchBoardColumnsForBoard_(boardId, apiToken) {
  const apiResponse = makeApiRequest_(`tm/board-columns?boardId=${boardId}`, apiToken);
  if (apiResponse && apiResponse.success && Array.isArray(apiResponse.boardColumns)) {
    return apiResponse.boardColumns;
  }
  Logger.log(`Не удалось загрузить колонки для доски ID: ${boardId}, или ответ не содержит массив 'boardColumns'.`);
  return []; 
}

function fetchAllTasks_(apiToken) {
  const requestParams = { "perPage": 1000000 }; 
  const apiResponse = makeApiRequest_("tm/tasks", apiToken, requestParams);
  
  if (apiResponse && apiResponse.success && Array.isArray(apiResponse.tasks)) { 
    Logger.log(`Загружено ${apiResponse.tasks.length} задач.`);
    if (apiResponse.hasMore === true) {
      Logger.log(`ВНИМАНИЕ: API WEEEK сообщил, что есть еще задачи (hasMore: true) даже при perPage=${requestParams.perPage}.`);
    }
    return apiResponse.tasks;
  } else {
    Logger.log("Ошибка при загрузке задач или ответ API некорректен/пуст. Возвращается пустой массив задач.");
    if (apiResponse) Logger.log(`Ответ API: ${JSON.stringify(apiResponse).substring(0, 500)}`);
    return []; 
  }
}

function fetchAllProjects_(apiToken) {
  Logger.log("Загрузка списка всех проектов...");
  const apiResponse = makeApiRequest_("tm/projects", apiToken); 
  const projectsMap = {};
  if (apiResponse && apiResponse.success && Array.isArray(apiResponse.projects)) { 
    apiResponse.projects.forEach(project => {
      if (project.id && (project.name || project.title)) projectsMap[project.id] = project.name || project.title;
    });
    Logger.log(`Загружено ${Object.keys(projectsMap).length} проектов.`);
  } else { Logger.log("Не удалось загрузить список проектов, или ответ не содержит массив 'projects'."); }
  return projectsMap;
}

function fetchWorkspaceDetails_(apiToken) {
  Logger.log("Загрузка деталей рабочего пространства...");
  const apiResponse = makeApiRequest_("ws", apiToken);
  if (apiResponse && apiResponse.success && apiResponse.workspace && apiResponse.workspace.id) {
    return apiResponse.workspace.id;
  } else { Logger.log("Не удалось загрузить детали рабочего пространства."); return null; }
}

function fetchWeeekData_(apiToken) {
  Logger.log("Загрузка всех данных из WEEEK...");
  const workspaceId = fetchWorkspaceDetails_(apiToken);
  const workspaceMembers = fetchWorkspaceMembers_(apiToken);
  const allTasksData = fetchAllTasks_(apiToken);
  const allProjectsMap = fetchAllProjects_(apiToken); 
  const boardColumnNamesMap = {}; 

  if (allTasksData.length > 0) {
    const uniqueBoardIdsRegistry = {}; 
    allTasksData.forEach(task => { if(task.boardId != null) uniqueBoardIdsRegistry[task.boardId] = true; });
    Object.keys(uniqueBoardIdsRegistry).forEach(boardIdString => {
      const boardIdNumber = parseInt(boardIdString, 10);
      if (!isNaN(boardIdNumber)) {
        fetchBoardColumnsForBoard_(boardIdNumber, apiToken).forEach(column => {
          if (column.id != null) {
            const columnIdNumber = parseInt(column.id, 10);
            if(!isNaN(columnIdNumber)) boardColumnNamesMap[`${boardIdNumber}_${columnIdNumber}`] = column.name || String(columnIdNumber);
          }
        });
      }
    });
  }
  return { workspaceId, membersMap: workspaceMembers, tasksData: allTasksData, boardColumnNamesMap, allProjectsMap };
}

// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
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
  const period1StartDay = 6, period1EndDay = 20, period2StartDay = 21, period2EndDay = 5;
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
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "dd.MM.yyyy");
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

function getContrastTextColor_(hexBackgroundColor) {
  if (!hexBackgroundColor || hexBackgroundColor.length < 7 || hexBackgroundColor[0] !== '#') return TEXT_COLOR_DARK;
  const r = parseInt(hexBackgroundColor.slice(1, 3), 16);
  const g = parseInt(hexBackgroundColor.slice(3, 5), 16);
  const b = parseInt(hexBackgroundColor.slice(5, 7), 16);
  const luma = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luma > 0.5 ? TEXT_COLOR_DARK : TEXT_COLOR_LIGHT;
}

function parseSheetDate_(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return null;
  const parts = dateStr.split('.');
  if (parts.length !== 3) return null;
  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; 
  const year = parseInt(parts[2], 10);
  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
  return new Date(year, month, day);
}

// ОБРАБОТКА ТАБЛИЦ
function processTasksToSheetRows_(allTasks, workspaceMembersMap, boardColumnNamesMap, allProjectsMap, currentWorkspaceId) {
  const sheetRows = []; 
  allTasks.forEach(taskItem => {
    let taskRepresentation; 
    if (taskItem.id && currentWorkspaceId) taskRepresentation = `https://app.weeek.net/ws/${currentWorkspaceId}/task/${taskItem.id}`;
    else if (taskItem.id) taskRepresentation = `(ID Задачи: ${taskItem.id})`; 
    else taskRepresentation = taskItem.title || ""; 

    const projectIdForTask = taskItem.projectId; 
    const projectName = (projectIdForTask && allProjectsMap[projectIdForTask]) ? allProjectsMap[projectIdForTask] : "";
    const priorityValue = taskItem.priority;
    const priorityString = priorityValue != null ? (PRIORITY_MAP[priorityValue] || String(priorityValue)) : "";
    
    let statusName = "";
    const taskBoardId = taskItem.boardId != null ? parseInt(taskItem.boardId, 10) : null;
    const taskColumnId = taskItem.boardColumnId != null ? parseInt(taskItem.boardColumnId, 10) : null;
    if (taskBoardId != null && !isNaN(taskBoardId) && taskColumnId != null && !isNaN(taskColumnId)) {
      statusName = boardColumnNamesMap[`${taskBoardId}_${taskColumnId}`] || `Неизвестный статус (ID: ${taskColumnId})`;
    } else if (taskItem.boardColumnId != null) {
      statusName = `Неизвестный статус (ID: ${taskItem.boardColumnId})`;
    }
    
    const creationDate = parseDateStringSafe_(taskItem.createdAt); 
    let completionDate = null;
    if (taskItem.isCompleted) {
      const timeEntries = (taskItem.timeEntries || taskItem.workloads || []).filter(e => e.date);
      if (timeEntries.length > 0) completionDate = parseDateStringSafe_(timeEntries.sort((a,b)=>new Date(b.date)-new Date(a.date))[0].date);
      else if (taskItem.updatedAt) completionDate = parseDateStringSafe_(taskItem.updatedAt);
    }
    
    let estimatedValue = (taskItem.duration!=null&&taskItem.duration>0)? (taskItem.duration/60) : null;

    const baseRowDataForTask = [taskRepresentation,projectName,priorityString, statusName,creationDate,completionDate,estimatedValue];
    
    const workloads = taskItem.workloads || [];
    if (workloads.length > 0) {
      workloads.forEach(workloadEntry => {
        const workloadUserId = workloadEntry.userId; 
        const workloadExecutorName = workloadUserId ? (workspaceMembersMap[workloadUserId] || workloadUserId) : "";
        const workloadDate = parseDateStringSafe_(workloadEntry.date);
        const durationMinutes = workloadEntry.duration || 0;
        const trackedHours = durationMinutes > 0 ? (durationMinutes / 60) : null;
        const commentText = (workloadEntry.comment || "").trim() || null;

        const fullRow = [
            baseRowDataForTask[0], baseRowDataForTask[1], baseRowDataForTask[2], 
            workloadExecutorName,
            baseRowDataForTask[3], baseRowDataForTask[4], baseRowDataForTask[5], baseRowDataForTask[6],
            workloadDate, trackedHours, commentText, null, null
        ];
        sheetRows.push(fullRow.map(formatValueForSheetCell_));
      });
    } else { 
      const executorIds = taskItem.executorIds || [];
      if (executorIds.length > 0) {
        executorIds.forEach(executorId => {
          const executorName = executorId ? (workspaceMembersMap[executorId] || executorId) : "";
          const fullRow = [
              baseRowDataForTask[0], baseRowDataForTask[1], baseRowDataForTask[2],
              executorName, 
              baseRowDataForTask[3], baseRowDataForTask[4], baseRowDataForTask[5], baseRowDataForTask[6],
              null, null, null, null, null
          ];
          sheetRows.push(fullRow.map(formatValueForSheetCell_));
        });
      } else {
        const taskOwnerId = taskItem.userId;
        const taskOwnerName = taskOwnerId ? (workspaceMembersMap[taskOwnerId] || taskOwnerId) : "";
        const fullRow = [
            baseRowDataForTask[0], baseRowDataForTask[1], baseRowDataForTask[2], 
            taskOwnerName, 
            baseRowDataForTask[3], baseRowDataForTask[4], baseRowDataForTask[5], baseRowDataForTask[6],
            null, null, null, null, null
        ];
        sheetRows.push(fullRow.map(formatValueForSheetCell_));
      }
    }
  });
  return sheetRows;
}

// GOOGLE SHEETS
function getOrCreateSheet_(spreadsheetObject, sheetNameForReport) {
  let targetSheet = spreadsheetObject.getSheetByName(sheetNameForReport);
  if (!targetSheet) {
    targetSheet = spreadsheetObject.insertSheet(sheetNameForReport);
    Logger.log(`Лист '${sheetNameForReport}' создан.`);
  } else {
    Logger.log(`Лист '${sheetNameForReport}' найден. Очистка...`);
  }
  return targetSheet;
}

function updateGoogleSheet_(
    googleSheetFileId, 
    targetSheetNameForReport, 
    reportHeadersArray, 
    reportDataRows,
    projectNamesForDropdown, 
    priorityValuesForDropdown, 
    executorNamesForDropdown, 
    statusNamesForDropdown  
  ) {
  try {
    const spreadsheet = SpreadsheetApp.openById(googleSheetFileId);
    const worksheet = getOrCreateSheet_(spreadsheet, targetSheetNameForReport);
    
    worksheet.clearConditionalFormatRules();
    const fullSheetRange = worksheet.getRange(1, 1, Math.max(1, worksheet.getMaxRows()), Math.max(1, worksheet.getMaxColumns()));
    fullSheetRange.clear({contentsOnly: true, formatOnly: true, validationsOnly: true});
    SpreadsheetApp.flush();

    const dataToWriteToSheet = [reportHeadersArray, ...reportDataRows];
    
    const headerMap = {};
    reportHeadersArray.forEach((header, index) => { headerMap[header] = index; });

    const trackedTimeColIdx = headerMap["Затреканное время"];
    const hourlyRateColIdx = headerMap["Стоимость в час"];
    const salaryCalcColIdx = headerMap["Расчет зп"];
    const priorityColIdx = headerMap["Приоритет"];
    const statusColIdx = headerMap["Статус"];
    const estimateColIdx = headerMap["Оценка"];
    const projectColIdx = headerMap["Проект"];
    const executorColIdx = headerMap["Исполнитель"];
    
    for (let rowIndex = 1; rowIndex < dataToWriteToSheet.length; rowIndex++) {
        const sheetRowNumber = rowIndex + 1; 
        if (trackedTimeColIdx > -1 && hourlyRateColIdx > -1 && salaryCalcColIdx > -1) {
            const trackedTimeLetter = convertColumnIndexToLetter_(trackedTimeColIdx + 1);
            const hourlyRateLetter = convertColumnIndexToLetter_(hourlyRateColIdx + 1);
            dataToWriteToSheet[rowIndex][salaryCalcColIdx] = 
                `=IFERROR(N(${trackedTimeLetter}${sheetRowNumber}) * N(${hourlyRateLetter}${sheetRowNumber}), "")`;
        }
    }
    
    if (dataToWriteToSheet.length > 0) {
      const targetRange = worksheet.getRange(1, 1, dataToWriteToSheet.length, reportHeadersArray.length);
      targetRange.setValues(dataToWriteToSheet); 
      targetRange.setFontFamily(FONT_FAMILY).setFontSize(10).setVerticalAlignment("middle")
               .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setBorder(true, true, true, true, true, true, "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID);
    }
    SpreadsheetApp.flush(); 

    worksheet.getRange(1, 1, 1, reportHeadersArray.length) 
      .setFontWeight("bold").setFontSize(11).setFontFamily(FONT_FAMILY) 
      .setHorizontalAlignment("center").setBackground(HEADER_BACKGROUND_COLOR).setFontColor(HEADER_FONT_COLOR);
    worksheet.setFrozenRows(1); 
    
    let conditionalFormatRules = []; 

    if (reportDataRows.length > 0) {
      const dataRowsRange = worksheet.getRange(2, 1, reportDataRows.length, reportHeadersArray.length);
      dataRowsRange.setFontWeight(null);

      conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=MOD(ROW(), 2) = 0').setBackground("#F0F8FF").setRanges([dataRowsRange]).build());
      
      if (priorityColIdx > -1) {
        const range = worksheet.getRange(2, priorityColIdx + 1, reportDataRows.length, 1);
        range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(priorityValuesForDropdown, true).setAllowInvalid(true).build());
        priorityValuesForDropdown.forEach(priorityText => {
          const bgColor = PRIORITY_COLORS_GSHEETS[priorityText];
          if (bgColor) {
            conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(priorityText).setBackground(bgColor).setFontColor(getContrastTextColor_(bgColor)).setRanges([range]).build());
          }
        });
      }
      if (statusColIdx > -1) {
        const range = worksheet.getRange(2, statusColIdx + 1, reportDataRows.length, 1);
        range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(statusNamesForDropdown, true).setAllowInvalid(true).build());
        statusNamesForDropdown.forEach(statusText => {
          const bgColor = STATUS_COLORS_GSHEETS[statusText] || OTHER_STATUS_BACKGROUND_COLOR;
          conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(statusText).setBackground(bgColor).setFontColor(getContrastTextColor_(bgColor)).setRanges([range]).build());
        });
      }
      if (projectColIdx > -1) {
        const range = worksheet.getRange(2, projectColIdx + 1, reportDataRows.length, 1);
        range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(projectNamesForDropdown, true).setAllowInvalid(true).build());
        projectNamesForDropdown.forEach(projectText => { 
          const bgColor = PROJECT_COLORS_GSHEETS[projectText]; 
          if (bgColor) { 
            conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(projectText).setBackground(bgColor).setFontColor(getContrastTextColor_(bgColor)).setRanges([range]).build());
          }
        });
      }
      if (executorColIdx > -1) {
        const range = worksheet.getRange(2, executorColIdx + 1, reportDataRows.length, 1);
        range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(executorNamesForDropdown, true).setAllowInvalid(true).build());
        executorNamesForDropdown.forEach(executorText => {
          const bgColor = EXECUTOR_COLORS_GSHEETS[executorText];
          if (bgColor) {
            conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(executorText).setBackground(bgColor).setFontColor(getContrastTextColor_(bgColor)).setRanges([range]).build());
          }
        });
      }
      
      const commentColIdx = headerMap["Комментарий к треку"];
      if (estimateColIdx > -1) {
        const range = worksheet.getRange(2, estimateColIdx + 1, reportDataRows.length, 1);
        conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground(ROW_ERROR_BACKGROUND_COLOR).setRanges([range]).build());
      }
      if (commentColIdx > -1) {
        const range = worksheet.getRange(2, commentColIdx + 1, reportDataRows.length, 1);
        conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground(ROW_ERROR_BACKGROUND_COLOR).setRanges([range]).build());
      }

      worksheet.setConditionalFormatRules(conditionalFormatRules);
      
      reportHeadersArray.forEach((headerName, columnIndexZeroBased) => {
        const columnRange = worksheet.getRange(2, columnIndexZeroBased + 1, reportDataRows.length, 1);
        switch(headerName) {
            case "Задача": case "Комментарий к треку": case "Проект":
                columnRange.setHorizontalAlignment("left"); break;
            case "Стоимость в час": case "Расчет зп":
                columnRange.setHorizontalAlignment("right"); break;
            default: columnRange.setHorizontalAlignment("center");
        }
        
        switch(headerName) {
            case "Расчет зп": case "Затреканное время": case "Оценка":
                columnRange.setNumberFormat("0.00"); break;
            case "Стоимость в час":
                columnRange.setNumberFormat("0"); break;
        }
      });
    }
    
    const columnWidthsSettings = { 
        "Задача": 300, "Проект": 150, "Приоритет": 100, "Исполнитель": 120, "Статус": 225,
        "Дата создания задачи": 100, "Дата закрытия задачи": 100, "Оценка": 70,
        "Дата трека": 100, "Затреканное время": 80, "Комментарий к треку": 400,
        "Стоимость в час": 80, "Расчет зп": 90
    };
    reportHeadersArray.forEach((headerName, colIndexZeroBased) => {
        const columnIndexOneBased = colIndexZeroBased + 1;
        try {
            if (columnWidthsSettings[headerName]) worksheet.setColumnWidth(columnIndexOneBased, columnWidthsSettings[headerName]);
            else worksheet.autoResizeColumn(columnIndexOneBased); 
        } catch (e) {
             Logger.log(`Не удалось изменить размер колонки '${headerName}': ${e.message}. Пропуск.`);
        }
    });
    Logger.log("Данные успешно записаны и отформатированы.");
  } catch (error) { 
    Logger.log(`Ошибка при работе с листом '${targetSheetNameForReport}': ${error.message}. Stack: ${error.stack}`); 
    throw error; 
  }
}

// ГЕНЕРАЦИЯ
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Отчет WEEEK')
    .addItem('Сгенерировать ВСЕ Отчеты', 'runFullReportGeneration')
    .addToUi();
}

function runFullReportGeneration() {
  let finalLogMessageOverall = "Задачи не были обработаны или произошла начальная ошибка.";
  let overallOperationSucceeded = false; 

  try {
    Logger.log("Запуск генерации отчетов WEEEK.");
    const weeekApiToken = getWeeekApiToken_();
    const googleSheetFileId = getGoogleSheetId_(); 
    const currentDate = new Date();
    
    const { workspaceId, membersMap, tasksData, boardColumnNamesMap, allProjectsMap } = fetchWeeekData_(weeekApiToken);

    if (workspaceId === null) {
        Logger.log("ПРЕДУПРЕЖДЕНИЕ: Не удалось получить ID рабочего пространства WEEEK. URL задач могут быть неполными.");
    }

    let allProcessedSheetRows = [];
    if (!tasksData || tasksData.length === 0) {
      Logger.log("Задачи не загружены.");
    } else {
      allProcessedSheetRows = processTasksToSheetRows_(
        tasksData, membersMap, boardColumnNamesMap, allProjectsMap, workspaceId 
      );
    }

    const reportHeaders = [ 
        "Задача", "Проект", "Приоритет", "Исполнитель", "Статус",
        "Дата создания задачи", "Дата закрытия задачи", "Оценка",
        "Дата трека", "Затреканное время", "Комментарий к треку",
        "Стоимость в час", "Расчет зп"
    ];
    
    const dateTrackColIdx = reportHeaders.indexOf("Дата трека");
    if (allProcessedSheetRows.length > 0 && dateTrackColIdx !== -1) {
      Logger.log("Сортировка задач по дате трека...");
      allProcessedSheetRows.sort((a, b) => {
        const dateA = parseSheetDate_(a[dateTrackColIdx]);
        const dateB = parseSheetDate_(b[dateTrackColIdx]);
        if (!dateA && !dateB) return 0;
        if (!dateA) return -1;
        if (!dateB) return 1;
        return dateA - dateB;
      });
    }

    const uniqueProjectNames = [...new Set(Object.values(allProjectsMap).filter(name => name))].sort();
    const priorityValues = Object.values(PRIORITY_MAP);
    const allUniqueExecutorNames = [...new Set(Object.values(membersMap).filter(name => name))].sort();
    
    const statusColIdx = reportHeaders.indexOf("Статус");
    const uniqueStatusNamesFromData = statusColIdx > -1 ? [...new Set(allProcessedSheetRows.map(row => row[statusColIdx]).filter(Boolean))] : [];
    const uniqueStatusNamesFromAPI = [...new Set(Object.values(boardColumnNamesMap).filter(Boolean))];
    const comprehensiveStatusNames = [...new Set([...uniqueStatusNamesFromAPI, ...uniqueStatusNamesFromData])].sort();

    Logger.log(`Создание '${GENERAL_REPORT_SHEET_NAME}'...`);
    updateGoogleSheet_(
        googleSheetFileId, GENERAL_REPORT_SHEET_NAME, reportHeaders, allProcessedSheetRows,
        uniqueProjectNames, priorityValues, allUniqueExecutorNames, comprehensiveStatusNames
    );
    Logger.log(`Лист '${GENERAL_REPORT_SHEET_NAME}' создан/обновлен.`);
    Utilities.sleep(1500); 

    Logger.log("Группировка задач для отчетов по исполнителям и периодам...");
    const reportsBySheetName = {}; 
    
    const [reportingPeriod1, reportingPeriod2] = getReportingPeriods_(currentDate);
    const periodsToProcess = [
      { name: `(${Utilities.formatDate(reportingPeriod1[0], Session.getScriptTimeZone(), "dd.MM")}-${Utilities.formatDate(reportingPeriod1[1], Session.getScriptTimeZone(), "dd.MM")})`, start: reportingPeriod1[0], end: reportingPeriod1[1] },
      { name: `(${Utilities.formatDate(reportingPeriod2[0], Session.getScriptTimeZone(), "dd.MM")}-${Utilities.formatDate(reportingPeriod2[1], Session.getScriptTimeZone(), "dd.MM")})`, start: reportingPeriod2[0], end: reportingPeriod2[1] }
    ];

    const executorColIdx = reportHeaders.indexOf("Исполнитель");
    const excludedExecutorsLower = EXCLUDED_EXECUTORS.map(name => name.toLowerCase());

    if (allProcessedSheetRows.length > 0 && executorColIdx !== -1 && dateTrackColIdx !== -1) {
      allProcessedSheetRows.forEach(row => {
        const executorName = row[executorColIdx];
        if (!executorName || typeof executorName !== 'string' || excludedExecutorsLower.includes(executorName.toLowerCase())) return;

        const trackDate = parseSheetDate_(row[dateTrackColIdx]);
        if (!trackDate) return;
        trackDate.setHours(0,0,0,0);

        for (const period of periodsToProcess) {
          const periodStartNorm = new Date(period.start); periodStartNorm.setHours(0,0,0,0);
          const periodEndNorm = new Date(period.end); periodEndNorm.setHours(0,0,0,0);

          if (trackDate >= periodStartNorm && trackDate <= periodEndNorm) {
            let sheetName = `${executorName} ${period.name}`.replace(/[\[\]\*\/\\\?\:]/g, "").substring(0, 95);
            if (!reportsBySheetName[sheetName]) reportsBySheetName[sheetName] = [];
            reportsBySheetName[sheetName].push(row);
            break; 
          }
        }
      });
    }
    
    Logger.log(`Начинаю создание ${Object.keys(reportsBySheetName).length} отчетов по исполнителям...`);
    const sortedSheetNames = Object.keys(reportsBySheetName).sort();
    for (const sheetName of sortedSheetNames) {
      const rowsForSheet = reportsBySheetName[sheetName];
      Logger.log(`Создание отчета '${sheetName}' с ${rowsForSheet.length} строками...`);
      updateGoogleSheet_(
        googleSheetFileId, sheetName, reportHeaders, rowsForSheet,
        uniqueProjectNames, priorityValues, allUniqueExecutorNames, comprehensiveStatusNames
      );
      Utilities.sleep(1500); 
    }

    Logger.log("Завершено создание отчетов по исполнителям и периодам.");
    finalLogMessageOverall = `Все отчеты WEEEK успешно сгенерированы!`;
    overallOperationSucceeded = true; 
    
  } catch (error) {
    finalLogMessageOverall = `Произошла глобальная ошибка: ${error.message}.`;
    Logger.log(`КРИТИЧЕСКАЯ ОШИБКА при генерации отчетов: ${error}. Stack: ${error.stack ? error.stack : 'недоступно'}`);
    overallOperationSucceeded = false; 
  } finally {
    const finalStatus = overallOperationSucceeded ? "УСПЕХ" : "ОШИБКА";
    Logger.log(`ЗАВЕРШЕНИЕ ВСЕХ ОПЕРАЦИЙ: ${finalStatus}. Итоговое сообщение: ${finalLogMessageOverall}`);
  }
}