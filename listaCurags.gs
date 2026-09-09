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
    Logger.log('=== INÍCIO createSlidesPresentation ===');
    Logger.log('Total de nomes recebidos: ' + names.length);
    
    var presentationName = 'Pedidos de Cura - Shabat ' + shabatDate;
    var presentation = SlidesApp.create(presentationName);
    var presentationId = presentation.getId();
    
    var validNames = names.filter(function(name) {
      return name && typeof name === 'string' && name.trim().length > 0;
    });
    
    if (validNames.length === 0) {
      throw new Error('Nenhum nome válido');
    }
    
    var totalNomes = validNames.length;
    Logger.log('Nomes válidos: ' + totalNomes);
    
    // === DIMENSÕES FIXAS DO SLIDE ===
    var SLIDE_WIDTH = 720;
    var SLIDE_HEIGHT = 405;
    
    // === COORDENADAS ABSOLUTAS ===
    var CARD_X = 30, CARD_Y = 20, CARD_WIDTH = 660, CARD_HEIGHT = 365;
    var TITLE_X = 50, TITLE_Y = 30, TITLE_WIDTH = 620, TITLE_HEIGHT = 35;
    
    // Área de nomes
    var NAMES_X = 50;
    var NAMES_Y = 75;  // Mais abaixo para compensar PowerPoint
    var NAMES_WIDTH = 620;
    var NAMES_HEIGHT = 285;  // Reduzido para evitar vazamento no PPTX
    
    var FOOTER_X = 50, FOOTER_Y = 365, FOOTER_WIDTH = 620, FOOTER_HEIGHT = 15;
    
    // === CONFIGURAÇÃO DE COLUNAS ===
    var COLUMNS = 4;
    var COL_PADDING = 10;
    var COL_SPACING = 10;
    var COLS_AVAILABLE_WIDTH = NAMES_WIDTH - (COL_PADDING * 2);
    var COL_WIDTH = (COLS_AVAILABLE_WIDTH - (COL_SPACING * (COLUMNS - 1))) / COLUMNS;
    
    // === CÁLCULO DINÂMICO DE SLIDES (mais conservador para PPTX) ===
    var maxNomesPorColuna = 27; // Reduzido de 28 para 27 (mais margem para PPTX)
    var maxNomesPorSlide = maxNomesPorColuna * COLUMNS; // 108
    var totalSlides = Math.ceil(totalNomes / maxNomesPorSlide);
    if (totalSlides < 1) totalSlides = 1;
    
    Logger.log('Total de slides: ' + totalSlides + ' (max ' + maxNomesPorColuna + ' nomes/coluna)');
    
    var slides = presentation.getSlides();
    var nomeAtual = 0;
    
    for (var slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
      Logger.log('Criando slide ' + (slideIndex + 1) + ' de ' + totalSlides);
      
      var nomesRestantes = totalNomes - nomeAtual;
      var slidesRestantes = totalSlides - slideIndex;
      var nomesNesteSlide = Math.ceil(nomesRestantes / slidesRestantes);
      
      var nomesPorColunaBase = Math.floor(nomesNesteSlide / COLUMNS);
      var colunasComExtra = nomesNesteSlide % COLUMNS;
      
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
      
      // CAMADA 1: FUNDO AZUL
      var bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT);
      bg.getFill().setSolidFill('#4A90E2');
      
      // CAMADA 2: FAIXAS DECORATIVAS
      var deco1 = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, SLIDE_WIDTH - 80, 0, 80, SLIDE_HEIGHT);
      deco1.getFill().setSolidFill('#357ABD');
      var deco2 = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, SLIDE_WIDTH - 40, 30, 40, SLIDE_HEIGHT - 60);
      deco2.getFill().setSolidFill('#2C6AA8');
      
      // CAMADA 3: CARD BRANCO
      var card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, CARD_X, CARD_Y, CARD_WIDTH, CARD_HEIGHT);
      card.getFill().setSolidFill('#FFFFFF');
      
      // CAMADA 4: TÍTULO
      var title = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, TITLE_X, TITLE_Y, TITLE_WIDTH, TITLE_HEIGHT);
      title.getText().setText('Pedidos de Cura - Shabat ' + shabatDate);
      title.getText().getTextStyle().setFontSize(20).setBold(true).setForegroundColor('#1a73e8');
      title.getFill().setTransparent();
      
      // CAMADA 5: NOMES EM 4 COLUNAS
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
          // Shape com margem extra para PowerPoint
          var colShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, colX, NAMES_Y - 8, COL_WIDTH, NAMES_HEIGHT + 16);
          colShape.getFill().setTransparent();
          
          var textRange = colShape.getText();
          textRange.setText(columnNames.join('\n'));
          textRange.getTextStyle().setFontSize(9).setFontFamily('Trebuchet MS').setForegroundColor('#555555');
          
          // Estilo otimizado para PowerPoint
          try {
            var paraStyle = textRange.getParagraphStyle();
            paraStyle.setSpaceAbove(0);
            paraStyle.setSpaceBelow(0);
            paraStyle.setLineSpacing(95); // 95% para compensar PowerPoint
          } catch (e) {
            Logger.log('Aviso coluna ' + col + ': ' + e.toString());
          }
        }
      }
      
      // CAMADA 6: FOOTER (apenas no último slide)
      if (slideIndex === totalSlides - 1) {
        var footerText = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, FOOTER_X, FOOTER_Y, FOOTER_WIDTH, FOOTER_HEIGHT);
        footerText.getText().setText('...e o Ponto de Estudos da Torah no Brasil e no Mundo.');
        footerText.getText().getTextStyle().setFontSize(9).setItalic(true).setForegroundColor('#666666');
        footerText.getFill().setTransparent();
      }
    }
    
    Logger.log('Slides criados com sucesso! ID: ' + presentationId);
    
    return { 
      success: true, 
      presentationId: presentationId, 
      presentationUrl: 'https://docs.google.com/presentation/d/' + presentationId + '/edit' 
    };
    
  } catch (e) {
    Logger.log('ERRO em createSlidesPresentation: ' + e.toString());
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
