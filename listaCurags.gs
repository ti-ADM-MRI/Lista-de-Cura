function doGet() {
  var template = HtmlService.createTemplateFromFile('listaCura');
  return template.evaluate()
    .setTitle('Pedidos de Cura - Shabat')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1') // <-- LINHA ESSENCIAL
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
    Logger.log('=== INÍCIO (OTIMIZADO PARA POWERPOINT) ===');
    
    var presentationName = 'Pedidos de Cura - Shabat ' + shabatDate;
    var presentation = SlidesApp.create(presentationName);
    var presentationId = presentation.getId();
    
    var validNames = names.filter(function(name) {
      return name && typeof name === 'string' && name.trim().length > 0;
    });
    
    if (validNames.length === 0) throw new Error('Nenhum nome válido');
    
    var totalNomes = validNames.length;
    
    // === DIMENSÕES ===
    var SLIDE_WIDTH = 720, SLIDE_HEIGHT = 405;
    var CARD_X = 35, CARD_Y = 25, CARD_WIDTH = 650, CARD_HEIGHT = 355;
    var TITLE_X = 55, TITLE_Y = 32, TITLE_WIDTH = 610, TITLE_HEIGHT = 38;
    var NAMES_X = 55, NAMES_Y = 80, NAMES_WIDTH = 610, NAMES_HEIGHT = 280;
    var FOOTER_X = 55, FOOTER_Y = 368, FOOTER_WIDTH = 610, FOOTER_HEIGHT = 12;
    
    var COLUMNS = 4;
    var COL_PADDING = 12, COL_SPACING = 12;
    var COLS_AVAILABLE_WIDTH = NAMES_WIDTH - (COL_PADDING * 2);
    var COL_WIDTH = (COLS_AVAILABLE_WIDTH - (COL_SPACING * (COLUMNS - 1))) / COLUMNS;
    
    // === MENOS NOMES POR COLUNA (margem para PowerPoint) ===
    var maxNomesPorColuna = 25; // Reduzido para PowerPoint
    var maxNomesPorSlide = maxNomesPorColuna * COLUMNS; // 100
    var totalSlides = Math.ceil(totalNomes / maxNomesPorSlide);
    if (totalSlides < 1) totalSlides = 1;
    
    Logger.log('Slides: ' + totalSlides + ' (25 nomes/coluna)');
    
    var slides = presentation.getSlides();
    var nomeAtual = 0;
    
    for (var slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
      var nomesRestantes = totalNomes - nomeAtual;
      var slidesRestantes = totalSlides - slideIndex;
      var nomesNesteSlide = Math.ceil(nomesRestantes / slidesRestantes);
      var nomesPorColunaBase = Math.floor(nomesNesteSlide / COLUMNS);
      var colunasComExtra = nomesNesteSlide % COLUMNS;
      
      var slide;
      if (slideIndex === 0 && slides.length > 0) {
        slide = slides[0];
        var pageElements = slide.getPageElements();
        for (var i = 0; i < pageElements.length; i++) pageElements[i].remove();
      } else {
        slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      }
      
      // 1. FUNDO
      var bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT);
      bg.getFill().setSolidFill('#4A90E2');
      
      // 2. FAIXAS
      slide.insertShape(SlidesApp.ShapeType.RECTANGLE, SLIDE_WIDTH - 75, 0, 75, SLIDE_HEIGHT).getFill().setSolidFill('#357ABD');
      slide.insertShape(SlidesApp.ShapeType.RECTANGLE, SLIDE_WIDTH - 38, 28, 38, SLIDE_HEIGHT - 56).getFill().setSolidFill('#2C6AA8');
      
      // 3. CARD
      slide.insertShape(SlidesApp.ShapeType.RECTANGLE, CARD_X, CARD_Y, CARD_WIDTH, CARD_HEIGHT).getFill().setSolidFill('#FFFFFF');
      
      // 4. TÍTULO
      var title = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, TITLE_X, TITLE_Y, TITLE_WIDTH, TITLE_HEIGHT);
      title.getText().setText('Pedidos de Cura - Shabat ' + shabatDate);
      title.getText().getTextStyle().setFontSize(20).setBold(true).setForegroundColor('#1a73e8');
      title.getFill().setTransparent();
      
      // 5. SHAPE DUMMY (absorve estilo padrão ANTES das colunas)
      var dummy = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, -500, -500, 10, 10);
      dummy.getFill().setTransparent();
      var dummyText = dummy.getText();
      dummyText.setText('x');
      dummyText.getTextStyle().setFontSize(1).setForegroundColor('#FFFFFF');
      try {
        dummyText.getParagraphStyle().setSpaceAbove(0).setSpaceBelow(0).setLineSpacing(100);
      } catch(e) {}
      
      // 6. COLUNAS
      for (var col = 0; col < COLUMNS; col++) {
        var colX = NAMES_X + COL_PADDING + (col * (COL_WIDTH + COL_SPACING));
        var nomesNestaColuna = nomesPorColunaBase + (col < colunasComExtra ? 1 : 0);
        
        var columnNames = [];
        for (var row = 0; row < nomesNestaColuna; row++) {
          if (nomeAtual < totalNomes) {
            columnNames.push(validNames[nomeAtual]);
            nomeAtual++;
          }
        }
        
        if (columnNames.length > 0) {
          var colShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, colX, NAMES_Y, COL_WIDTH, NAMES_HEIGHT);
          colShape.getFill().setTransparent();
          
          var textRange = colShape.getText();
          textRange.setText(columnNames.join('\n'));
          textRange.getTextStyle().setFontSize(9).setFontFamily('Trebuchet MS').setForegroundColor('#555555');
          
          try {
            textRange.getParagraphStyle().setSpaceAbove(0).setSpaceBelow(0).setLineSpacing(100);
          } catch(e) {}
        }
      }
      
      // 7. FOOTER
      if (slideIndex === totalSlides - 1) {
        var footer = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, FOOTER_X, FOOTER_Y, FOOTER_WIDTH, FOOTER_HEIGHT);
        footer.getText().setText('...e o Ponto de Estudos da Torah no Brasil e no Mundo.');
        footer.getText().getTextStyle().setFontSize(8).setItalic(true).setForegroundColor('#666666');
        footer.getFill().setTransparent();
      }
    }
    
    return { success: true, presentationId: presentationId, presentationUrl: 'https://docs.google.com/presentation/d/' + presentationId + '/edit' };
    
  } catch (e) {
    Logger.log('ERRO: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function regenerateSlidesForDev(shabatDate) {
  try {
    Logger.log('=== INÍCIO regenerateSlidesForDev ===');
    
    var result = getValidSpreadsheet(shabatDate);
    if (!result) throw new Error('Planilha não encontrada');
    
    var ss = result.ss;
    var sheet = ss.getSheetByName('Pedidos');
    if (!sheet) throw new Error('Aba Pedidos não encontrada');
    
    var allNames = [];
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    
    Logger.log('Planilha: ' + lastRow + ' linhas x ' + lastColumn + ' colunas');
    
    // LEITURA COLUNA POR COLUNA (ordem correta)
    for (var col = 1; col <= lastColumn; col++) {
      for (var row = 2; row <= lastRow; row++) {
        var cellValue = sheet.getRange(row, col).getValue();
        if (cellValue && cellValue.toString().trim() !== '') {
          allNames.push(cellValue.toString().trim());
        }
      }
    }
    
    Logger.log('Total de nomes lidos: ' + allNames.length);
    
    if (allNames.length === 0) {
      throw new Error('Nenhum nome encontrado na planilha');
    }
    
    return createSlidesPresentation(allNames, shabatDate);
    
  } catch (e) {
    Logger.log('ERRO regenerateSlidesForDev: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function checkFormStatus() {
  var now = new Date();
  var shabatInfo = getShabatDate();
  var shabatDate = shabatInfo.shabatDate;
  var isDeveloper = checkDeveloperAccess();
  
  // Calcula quarta-feira (3 dias antes do sábado)
  var wednesdayDate = new Date(shabatDate);
  wednesdayDate.setDate(shabatDate.getDate() - 3); // CORRIGIDO: -3 dias
  
  // Calcula segunda-feira (2 dias depois do sábado)
  var mondayDate = new Date(shabatDate);
  mondayDate.setDate(shabatDate.getDate() + 2);
  
  // Define horários limites
  var wednesdayOpen = new Date(wednesdayDate);
  wednesdayOpen.setHours(0, 1, 0, 0); // Quarta-feira 00:01
  
  var saturdayClose = new Date(shabatDate);
  saturdayClose.setHours(10, 1, 0, 0); // Sábado 10:01
  
  var mondayClose = new Date(mondayDate);
  mondayClose.setHours(16, 0, 0, 0); // Segunda-feira 16:00
  
  var isOpen = false;
  var message = '';
  var showWarning = false;
  
  if (isDeveloper) {
    isOpen = true;
    message = 'MODO DESENVOLVEDOR - Formulário aberto para o Shabat ' + shabatInfo.dateString;
    showWarning = false;
  }
  else if (now >= wednesdayOpen && now <= saturdayClose) {
    isOpen = true;
    message = 'Formulário aberto para o Shabat ' + shabatInfo.dateString;
  } else if (now > saturdayClose && now <= mondayClose) {
    isOpen = false;
    showWarning = true;
    message = 'Formulário fechado. Próxima abertura: Quarta-feira ' + 
              formatDateString(getNextWednesday()) + ' às 00:01';
  } else {
    isOpen = false;
    showWarning = true;
    message = 'Formulário fechado. Abre na Quarta-feira ' + 
              formatDateString(getNextWednesday()) + ' às 00:01';
  }
  
  return {
    isOpen: isOpen,
    message: message,
    showWarning: showWarning,
    shabatDate: shabatInfo.dateString,
    currentTime: now.toLocaleString('pt-BR'),
    isDeveloper: isDeveloper
  };
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
    var developerEmails = ['patricia.fonseca.sf@gmail.com', 'ysamrocha@gmail.com', 'administrativa@remanescentedeisrael.com'];
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

function debugPlanilha() {
  try {
    var shabatDate = '12/09/2026';
    var result = getValidSpreadsheet(shabatDate);
    
    if (!result) {
      return { error: 'Planilha não encontrada' };
    }
    
    var ss = result.ss;
    var sheet = ss.getSheetByName('Pedidos');
    
    if (!sheet) {
      return { error: 'Aba Pedidos não encontrada' };
    }
    
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    
    Logger.log('=== DEBUG PLANILHA ===');
    Logger.log('ID da planilha: ' + result.id);
    Logger.log('Nome: ' + result.name);
    Logger.log('Última linha: ' + lastRow);
    Logger.log('Última coluna: ' + lastColumn + ' (letra: ' + String.fromCharCode(64 + lastColumn) + ')');
    
    var totalNomes = 0;
    var nomesPorColuna = [];
    
    for (var col = 1; col <= lastColumn; col++) {
      var nomesNaColuna = 0;
      for (var row = 2; row <= lastRow; row++) {
        var cellValue = sheet.getRange(row, col).getValue();
        if (cellValue && cellValue.toString().trim() !== '') {
          nomesNaColuna++;
          totalNomes++;
        }
      }
      nomesPorColuna.push({
        coluna: String.fromCharCode(64 + col),
        nomes: nomesNaColuna
      });
      Logger.log('Coluna ' + String.fromCharCode(64 + col) + ': ' + nomesNaColuna + ' nomes');
    }
    
    Logger.log('TOTAL DE NOMES: ' + totalNomes);
    
    return {
      spreadsheetId: result.id,
      totalNomes: totalNomes,
      totalColunas: lastColumn,
      nomesPorColuna: nomesPorColuna
    };
    
  } catch (e) {
    Logger.log('ERRO debugPlanilha: ' + e.toString());
    return { error: e.toString() };
  }
}

function checkDeveloperAccess() {
  var email = Session.getActiveUser().getEmail();
  
  Logger.log('=== CHECK DEV ACCESS ===');
  Logger.log('Email recebido: "' + email + '"');
  Logger.log('Email é vazio?: ' + (email === ''));
  Logger.log('Email é null?: ' + (email === null));
  
  var devEmails = [
    'ysamrocha@gmail.com',
    'administrativa@remanescentedeisrael.com'
  ];
  
  if (!email || email.trim() === '') {
    Logger.log('❌ USUÁRIO NÃO LOGADO - Retornando FALSE');
    return false;
  }
  
  var isDev = devEmails.indexOf(email.toLowerCase()) !== -1;
  Logger.log('Email está na lista?: ' + isDev);
  Logger.log('===================');
  
  return isDev;
}
