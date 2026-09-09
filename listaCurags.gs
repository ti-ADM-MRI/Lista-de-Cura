function doGet() {
  var template = HtmlService.createTemplateFromFile('listaCura');
  return template.evaluate()
    .setTitle('Pedidos de Cura - Shabat')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function normalizeName(name) {
  if (!name || typeof name !== 'string') return '';
  var normalized = name.toLowerCase().trim();
  var accents = {
    'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a',
    'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
    'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
    'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
    'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
    'ç': 'c', 'ñ': 'n'
  };
  for (var accent in accents) {
    normalized = normalized.replace(new RegExp(accent, 'g'), accents[accent]);
  }
  normalized = normalized.replace(/\s+/g, ' ');
  normalized = normalized.replace(/\b(dos|das|de|do|da)\b/g, '').trim();
  normalized = normalized.replace(/\s+/g, ' ');
  return normalized;
}

function getExistingNames(sheet) {
  var existingNames = [];
  try {
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow > 1 && lastColumn > 0) {
      var range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
      var values = range.getValues();
      for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
          var cellValue = values[row][col];
          if (cellValue && cellValue.toString().trim() !== '') {
            existingNames.push(normalizeName(cellValue.toString()));
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Erro ao ler nomes: ' + e.toString());
  }
  return existingNames;
}

function getNextAvailableCell(sheet) {
  var namesPerColumn = 32;
  var totalNames = 0;
  try {
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow > 1 && lastColumn > 0) {
      var range = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), lastColumn);
      var values = range.getValues();
      for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
          if (values[row][col] && values[row][col].toString().trim() !== '') {
            totalNames++;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Erro ao contar nomes: ' + e.toString());
  }
  var columnIndex = Math.floor(totalNames / namesPerColumn) + 1;
  var rowIndex = (totalNames % namesPerColumn) + 2;
  return { row: rowIndex, column: columnIndex };
}

function saveNames(names, shabatDate) {
  var status = checkFormStatus();
  if (!status.isOpen) {
    throw new Error('Formulário está fechado. ' + status.message);
  }
  
  var spreadsheetName = 'Pedidos de Cura - Shabat ' + shabatDate;
  var ss;
  var spreadsheetId;
  var maxRetries = 3;
  
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      var files = DriveApp.getFilesByName(spreadsheetName);
      var found = false;
      
      while (files.hasNext()) {
        var file = files.next();
        if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
          spreadsheetId = file.getId();
          ss = SpreadsheetApp.openById(spreadsheetId);
          found = true;
          break;
        }
      }
      
      if (!found) {
        ss = SpreadsheetApp.create(spreadsheetName);
        spreadsheetId = ss.getId();
        var sheet = ss.getActiveSheet();
        sheet.setName('Pedidos');
        for (var col = 1; col <= 10; col++) {
          sheet.setColumnWidth(col, 250);
        }
        sheet.getRange(1, 1).setValue('Nome Completo');
        sheet.getRange(1, 1).setFontWeight('bold');
        sheet.getRange(1, 1).setBackground('#4285f4');
        sheet.getRange(1, 1).setFontColor('white');
        Utilities.sleep(1000);
      }
      
      var sheet = ss.getSheetByName('Pedidos');
      if (!sheet) {
        sheet = ss.getSheets()[0];
        sheet.setName('Pedidos');
      }
      
      var existingNames = getExistingNames(sheet);
      var newNames = [];
      var duplicates = [];
      
      for (var i = 0; i < names.length; i++) {
        var name = names[i];
        if (!name || typeof name !== 'string') continue;
        
        var sanitizedName = name.replace(/[<>{}[\]();'"]/g, '').trim();
        if (sanitizedName.length < 3 || sanitizedName.length > 100) continue;
        
        var normalizedName = normalizeName(sanitizedName);
        if (existingNames.indexOf(normalizedName) !== -1) {
          duplicates.push(sanitizedName);
        } else {
          newNames.push(sanitizedName);
          existingNames.push(normalizedName);
        }
      }
      
      if (newNames.length === 0) {
        var msg = 'Nenhum nome novo adicionado.' + (duplicates.length > 0 ? ' ' + duplicates.length + ' duplicado(s) ignorado(s).' : '');
        return { success: true, message: msg, duplicates: duplicates.length, newNames: 0, spreadsheetId: spreadsheetId, spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit' };
      }
      
      var nextCell = getNextAvailableCell(sheet);
      var currentRow = nextCell.row;
      var currentColumn = nextCell.column;
      
      for (var i = 0; i < newNames.length; i++) {
        sheet.getRange(currentRow, currentColumn).setValue(newNames[i]);
        currentRow++;
        if (currentRow > 33) {
          currentRow = 2;
          currentColumn++;
        }
      }
      
      var msg = newNames.length + ' nome(s) adicionado(s)!' + (duplicates.length > 0 ? ' ' + duplicates.length + ' duplicado(s) ignorado(s).' : '');
      return { success: true, message: msg, duplicates: duplicates.length, newNames: newNames.length, spreadsheetId: spreadsheetId, spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit' };
      
    } catch (e) {
      if (attempt === maxRetries) throw new Error('Falha após ' + maxRetries + ' tentativas: ' + e.toString());
      Utilities.sleep(2000 * attempt);
    }
  }
}

function createSlidesPresentation(names, shabatDate) {
  try {
    var presentationName = 'Pedidos de Cura - Shabat ' + shabatDate;
    var presentation = SlidesApp.create(presentationName);
    var presentationId = presentation.getId();
    
    var validNames = names.filter(function(name) {
      return name && typeof name === 'string' && name.trim().length > 0;
    });
    
    if (validNames.length === 0) {
      throw new Error('Nenhum nome válido');
    }
    
    var namesPerSlide = 32;
    var totalSlides = Math.ceil(validNames.length / namesPerSlide);
    var slides = presentation.getSlides();
    
    for (var slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
      var slide;
      
      if (slideIndex === 0 && slides.length > 0) {
        slide = slides[0];
        var pageElements = slide.getPageElements();
        for (var i = 0; i < pageElements.length; i++) {
          pageElements[i].remove();
        }
      } else {
        slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      }
      
      // Título simplificado
      var title = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 50, 20, 620, 60);
      title.getText().setText('Pedidos de Cura - Shabat ' + shabatDate);
      title.getText().getTextStyle().setFontSize(24).setBold(true);
      title.getFill().setTransparent();
      
      var startIndex = slideIndex * namesPerSlide;
      var endIndex = Math.min(startIndex + namesPerSlide, validNames.length);
      var slideNames = validNames.slice(startIndex, endIndex);
      
      var columnWidth = 150;
      var columnSpacing = 20;
      var startX = 50;
      var rowHeight = 28;
      var startY = 100;
      var columns = 4;
      
      for (var i = 0; i < slideNames.length; i++) {
        var col = i % columns;
        var row = Math.floor(i / columns);
        var xPos = startX + (col * (columnWidth + columnSpacing));
        var yPos = startY + (row * rowHeight);
        
        var nameShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, xPos, yPos, columnWidth, 25);
        nameShape.getText().setText(slideNames[i]);
        nameShape.getText().getTextStyle().setFontSize(11);
        nameShape.getFill().setTransparent();
      }
      
      // Footer simplificado - SEM setAlignment
      if (slideIndex === totalSlides - 1) {
        var footerText = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 50, 450, 620, 60);
        footerText.getText().setText('...e o Ponto de Estudos da Torah no Brasil e no Mundo.');
        footerText.getText().getTextStyle().setFontSize(14).setItalic(true);
        footerText.getFill().setTransparent();
      }
    }
    
    return { 
      success: true, 
      presentationId: presentationId, 
      presentationUrl: 'https://docs.google.com/presentation/d/' + presentationId + '/edit' 
    };
    
  } catch (e) {
    Logger.log('ERRO: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function checkFormStatus() {
  var now = new Date();
  var shabatInfo = getShabatDate();
  var shabatDate = shabatInfo.shabatDate;
  var isDeveloper = checkDeveloperAccess();
  
  var wednesdayDate = new Date(shabatDate);
  wednesdayDate.setDate(shabatDate.getDate() - 2);
  var mondayDate = new Date(shabatDate);
  mondayDate.setDate(shabatDate.getDate() + 2);
  
  var wednesdayOpen = new Date(wednesdayDate);
  wednesdayOpen.setHours(0, 1, 0, 0);
  var saturdayClose = new Date(shabatDate);
  saturdayClose.setHours(13, 0, 0, 0);
  var mondayClose = new Date(mondayDate);
  mondayClose.setHours(16, 0, 0, 0);
  
  var isOpen = false, message = '', showWarning = false;
  
  if (isDeveloper) {
    isOpen = true;
    message = 'MODO DESENVOLVEDOR - Formulário aberto para o Shabat ' + shabatInfo.dateString;
  } else if (now >= wednesdayOpen && now <= saturdayClose) {
    isOpen = true;
    message = 'Formulário aberto para o Shabat ' + shabatInfo.dateString;
  } else {
    isOpen = false;
    showWarning = true;
    message = 'Formulário fechado. Abre na Quarta-feira ' + formatDateString(getNextWednesday()) + ' às 00:01';
  }
  
  return { isOpen: isOpen, message: message, showWarning: showWarning, shabatDate: shabatInfo.dateString, currentTime: now.toLocaleString('pt-BR'), isDeveloper: isDeveloper };
}

function getShabatDate() {
  var today = new Date();
  var daysUntilSaturday = 6 - today.getDay();
  if (daysUntilSaturday < 0) daysUntilSaturday += 7;
  var shabatDate = new Date(today);
  shabatDate.setDate(today.getDate() + daysUntilSaturday);
  shabatDate.setHours(0, 0, 0, 0);
  return {
    dateString: String(shabatDate.getDate()).padStart(2, '0') + '/' + String(shabatDate.getMonth() + 1).padStart(2, '0') + '/' + shabatDate.getFullYear(),
    shabatDate: shabatDate,
    daysUntilSaturday: daysUntilSaturday
  };
}

function getNextWednesday() {
  var today = new Date();
  var daysUntilWednesday = 3 - today.getDay();
  if (daysUntilWednesday <= 0) daysUntilWednesday += 7;
  var nextWednesday = new Date(today);
  nextWednesday.setDate(today.getDate() + daysUntilWednesday);
  return nextWednesday;
}

function formatDateString(date) {
  return String(date.getDate()).padStart(2, '0') + '/' + String(date.getMonth() + 1).padStart(2, '0') + '/' + date.getFullYear();
}

function processForm(names, shabatDate) {
  try {
    if (!names || !Array.isArray(names)) throw new Error('Parâmetro names inválido');
    if (!shabatDate) throw new Error('Parâmetro shabatDate inválido');
    
    var spreadsheetResult = saveNames(names, shabatDate);
    var presentationResult = createSlidesPresentation(names, shabatDate);
    
    return { success: true, spreadsheet: spreadsheetResult, presentation: presentationResult, message: spreadsheetResult.message };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function checkDeveloperAccess() {
  try {
    var email = Session.getActiveUser().getEmail();
    var developerEmails = ['ysamrocha@gmail.com', 'administrativa@remanescentedeisrael.com'];
    return developerEmails.indexOf(email) !== -1;
  } catch (e) {
    return false;
  }
}

// =====================================================
// NOVA FUNÇÃO: Busca APENAS planilhas (ignora apresentações)
// =====================================================
function getValidSpreadsheet(shabatDate) {
  var spreadsheetName = 'Pedidos de Cura - Shabat ' + shabatDate;
  Logger.log('Buscando planilha (filtrando por tipo): ' + spreadsheetName);
  
  var files = DriveApp.getFilesByName(spreadsheetName);
  
  while (files.hasNext()) {
    var file = files.next();
    
    // FILTRO CRÍTICO: Só aceita Google Sheets, ignora Slides/PDFs/etc
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      try {
        var ss = SpreadsheetApp.openById(file.getId());
        Logger.log('✅ Planilha válida encontrada! ID: ' + file.getId());
        return { ss: ss, id: file.getId(), name: spreadsheetName };
      } catch (e) {
        Logger.log('❌ Arquivo ignorado (ID inválido): ' + file.getId());
      }
    } else {
      Logger.log('️ Arquivo ignorado (não é planilha): ' + file.getName() + ' - Tipo: ' + file.getMimeType());
    }
  }
  
  Logger.log('Nenhuma planilha válida encontrada');
  return null;
}

function getDeveloperStats(shabatDate) {
  try {
    var result = getValidSpreadsheet(shabatDate);
    
    if (!result) {
      return { 
        success: false, 
        error: 'Planilha não encontrada para o Shabat ' + shabatDate,
        notCreated: true
      };
    }
    
    var ss = result.ss;
    var sheet = ss.getSheetByName('Pedidos');
    
    if (!sheet) {
      return { success: false, error: 'Aba "Pedidos" não encontrada' };
    }
    
    var totalNames = 0;
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    
    if (lastRow > 1 && lastColumn > 0) {
      var range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
      var values = range.getValues();
      for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
          if (values[row][col] && values[row][col].toString().trim() !== '') {
            totalNames++;
          }
        }
      }
    }
    
    return {
      success: true,
      totalNames: totalNames,
      totalSlides: Math.ceil(totalNames / 32),
      spreadsheetId: result.id,
      spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + result.id + '/edit',
      spreadsheetName: result.name,
      lastModified: new Date().toLocaleString('pt-BR')
    };
    
  } catch (e) {
    Logger.log('ERRO getDeveloperStats: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function exportNamesAsJSON(shabatDate) {
  try {
    var result = getValidSpreadsheet(shabatDate);
    if (!result) throw new Error('Planilha não encontrada');
    
    var ss = result.ss;
    var sheet = ss.getSheetByName('Pedidos');
    if (!sheet) throw new Error('Aba Pedidos não encontrada');
    
    var allNames = [];
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    
    if (lastRow > 1 && lastColumn > 0) {
      var range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
      var values = range.getValues();
      for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
          if (values[row][col] && values[row][col].toString().trim() !== '') {
            allNames.push({
              nome: values[row][col].toString().trim(),
              coluna: String.fromCharCode(65 + col),
              linha: row + 2
            });
          }
        }
      }
    }
    
    return {
      spreadsheetName: result.name,
      spreadsheetId: result.id,
      totalNomes: allNames.length,
      exportadoEm: new Date().toLocaleString('pt-BR'),
      nomes: allNames
    };
    
  } catch (e) {
    Logger.log('ERRO exportNamesAsJSON: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function regenerateSlidesForDev(shabatDate) {
  try {
    var result = getValidSpreadsheet(shabatDate);
    if (!result) throw new Error('Planilha não encontrada');
    
    var ss = result.ss;
    var sheet = ss.getSheetByName('Pedidos');
    if (!sheet) throw new Error('Aba Pedidos não encontrada');
    
    var allNames = [];
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    
    if (lastRow > 1 && lastColumn > 0) {
      var range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
      var values = range.getValues();
      for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
          if (values[row][col] && values[row][col].toString().trim() !== '') {
            allNames.push(values[row][col].toString().trim());
          }
        }
      }
    }
    
    if (allNames.length === 0) {
      throw new Error('Nenhum nome encontrado na planilha');
    }
    
    return createSlidesPresentation(allNames, shabatDate);
    
  } catch (e) {
    Logger.log('ERRO regenerateSlidesForDev: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
