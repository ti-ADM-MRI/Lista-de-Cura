/*Funcionalidades:
   Cria planilha automaticamente com o nome "Pedidos de Cura - Shabat DD/MM/AAAA"
   Formulário dinâmico com botão (+) para adicionar mais campos
   Botão (-) para remover campos
   Botão "Enviar" salva na planilha e gera apresentação
   Formulário permanece na tela para ajustes
   Gera Google Slides com layout em 4 colunas
   Último slide contém o texto "...e o Ponto de Estudos da Torah no Brasil e no Mundo."
   Sem travamentos mesmo com muitos nomes
   Links da planilha e apresentação exibidos após envio
  
  Resumo das mudanças implementadas:
   Planilha em colunas: 32 nomes por coluna (A2-A33, B2-B33, etc.)
   Abertura: Quarta-feira 00:01 (2 dias antes do sábado)
   Fechamento: Sábado 13:00
   Aviso: Mostra mensagem quando fechado, informando próxima abertura
   Preservação: Dados não são apagados até segunda-feira 16:00
   Validação: Impede envio fora do horário permitido
*/
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
  
  // Remove caracteres perigosos e normaliza
  var normalized = name.toLowerCase().trim();
  
  // Remove acentos
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
  
  // Remove espaços múltiplos
  normalized = normalized.replace(/\s+/g, ' ');
  
  // Remove artigos/preposições
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
      Logger.log('Tentativa ' + attempt + ' de ' + maxRetries);
      
      // Busca planilha existente de forma mais robusta
      var files = DriveApp.getFilesByName(spreadsheetName);
      var found = false;
      
      while (files.hasNext()) {
        var file = files.next();
        // Verifica se é realmente uma planilha Google
        if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
          spreadsheetId = file.getId();
          ss = SpreadsheetApp.openById(spreadsheetId);
          found = true;
          Logger.log('Planilha encontrada: ' + spreadsheetId);
          break;
        }
      }
      
      if (!found) {
        // Cria nova planilha
        ss = SpreadsheetApp.create(spreadsheetName);
        spreadsheetId = ss.getId();
        Logger.log('Nova planilha criada: ' + spreadsheetId);
        
        var sheet = ss.getActiveSheet();
        sheet.setName('Pedidos');
        
        for (var col = 1; col <= 10; col++) {
          sheet.setColumnWidth(col, 250);
        }
        
        sheet.getRange(1, 1).setValue('Nome Completo');
        sheet.getRange(1, 1).setFontWeight('bold');
        sheet.getRange(1, 1).setBackground('#4285f4');
        sheet.getRange(1, 1).setFontColor('white');
        
        // Aguarda um momento para garantir que a planilha está pronta
        Utilities.sleep(1000);
      }
      
      // Garante que existe aba 'Pedidos'
      var sheet = ss.getSheetByName('Pedidos');
      if (!sheet) {
        var sheets = ss.getSheets();
        sheet = sheets[0];
        sheet.setName('Pedidos');
      }
      
      // Lê nomes existentes
      var existingNames = getExistingNames(sheet);
      Logger.log('Total de nomes existentes: ' + existingNames.length);
      
      // Filtra e valida nomes
      var newNames = [];
      var duplicates = [];
      var invalidNames = [];
      
      for (var i = 0; i < names.length; i++) {
        var name = names[i];
        
        if (!name || typeof name !== 'string') {
          invalidNames.push('Nome inválido');
          continue;
        }
        
        var sanitizedName = name.replace(/[<>{}[\]();'"]/g, '').trim();
        
        if (sanitizedName.length < 3) {
          invalidNames.push(name);
          continue;
        }
        
        if (sanitizedName.length > 100) {
          invalidNames.push(name.substring(0, 50) + '...');
          continue;
        }
        
        var normalizedName = normalizeName(sanitizedName);
        var isDuplicate = existingNames.indexOf(normalizedName) !== -1;
        
        if (isDuplicate) {
          duplicates.push(sanitizedName);
        } else {
          newNames.push(sanitizedName);
          existingNames.push(normalizedName);
        }
      }
      
      Logger.log('Novos: ' + newNames.length + ', Duplicados: ' + duplicates.length);
      
      if (newNames.length === 0) {
        var msg = 'Nenhum nome novo adicionado.';
        if (duplicates.length > 0) msg += ' ' + duplicates.length + ' duplicado(s) ignorado(s).';
        
        return {
          success: true,
          message: msg,
          duplicates: duplicates.length,
          newNames: 0,
          spreadsheetId: spreadsheetId,
          spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit'
        };
      }
      
      // Encontra próxima célula
      var nextCell = getNextAvailableCell(sheet);
      Logger.log('Próxima célula: ' + String.fromCharCode(64 + nextCell.column) + nextCell.row);
      
      // Insere nomes
      var namesPerColumn = 32;
      var currentRow = nextCell.row;
      var currentColumn = nextCell.column;
      
      for (var i = 0; i < newNames.length; i++) {
        sheet.getRange(currentRow, currentColumn).setValue(newNames[i]);
        
        currentRow++;
        if (currentRow > namesPerColumn + 1) {
          currentRow = 2;
          currentColumn++;
        }
      }
      
      var msg = newNames.length + ' nome(s) adicionado(s)!';
      if (duplicates.length > 0) msg += ' ' + duplicates.length + ' duplicado(s) ignorado(s).';
      
      return {
        success: true,
        message: msg,
        duplicates: duplicates.length,
        newNames: newNames.length,
        spreadsheetId: spreadsheetId,
        spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit'
      };
      
    } catch (e) {
      Logger.log('Erro na tentativa ' + attempt + ': ' + e.toString());
      
      if (attempt === maxRetries) {
        throw new Error('Falha após ' + maxRetries + ' tentativas: ' + e.toString());
      }
      
      // Aguarda antes de tentar novamente
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
      
      var title = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 50, 20, 620, 60);
      title.getText().setText('Pedidos de Cura - Shabat ' + shabatDate);
      title.getText().getTextStyle().setFontSize(24).setBold(true).setForegroundColor('#1a73e8');
      title.getFill().setTransparent();
      title.setBorder(null, null, null);
      
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
        nameShape.setBorder(null, null, null);
      }
      
      if (slideIndex === totalSlides - 1) {
        var footerText = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 50, 450, 620, 60);
        footerText.getText().setText('...e o Ponto de Estudos da Torah no Brasil e no Mundo.');
        footerText.getText().getTextStyle().setFontSize(14).setItalic(true).setForegroundColor('#666666');
        footerText.getFill().setTransparent();
        footerText.setBorder(null, null, null);
        footerText.getText().getParagraphStyle().setAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }
    }
    
    return {
      success: true,
      presentationId: presentationId,
      presentationUrl: 'https://docs.google.com/presentation/d/' + presentationId + '/edit'
    };
    
  } catch (e) {
    Logger.log('ERRO slides: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function checkFormStatus() {
  var now = new Date();
  var shabatInfo = getShabatDate();
  var shabatDate = shabatInfo.shabatDate;
  
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
  
  var isOpen = false;
  var message = '';
  var showWarning = false;
  
  if (now >= wednesdayOpen && now <= saturdayClose) {
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
    currentTime: now.toLocaleString('pt-BR')
  };
}

function getShabatDate() {
  var today = new Date();
  var dayOfWeek = today.getDay();
  var daysUntilSaturday = 6 - dayOfWeek;
  if (daysUntilSaturday < 0) daysUntilSaturday += 7;
  
  var shabatDate = new Date(today);
  shabatDate.setDate(today.getDate() + daysUntilSaturday);
  shabatDate.setHours(0, 0, 0, 0);
  
  var day = String(shabatDate.getDate()).padStart(2, '0');
  var month = String(shabatDate.getMonth() + 1).padStart(2, '0');
  var year = shabatDate.getFullYear();
  
  return {
    dateString: day + '/' + month + '/' + year,
    shabatDate: shabatDate,
    daysUntilSaturday: daysUntilSaturday
  };
}

function getNextWednesday() {
  var today = new Date();
  var currentDay = today.getDay();
  var daysUntilWednesday = 3 - currentDay;
  
  if (daysUntilWednesday <= 0) {
    daysUntilWednesday += 7;
  }
  
  var nextWednesday = new Date(today);
  nextWednesday.setDate(today.getDate() + daysUntilWednesday);
  
  return nextWednesday;
}

function formatDateString(date) {
  var day = String(date.getDate()).padStart(2, '0');
  var month = String(date.getMonth() + 1).padStart(2, '0');
  var year = date.getFullYear();
  return day + '/' + month + '/' + year;
}

function processForm(names, shabatDate) {
  try {
    if (!names || !Array.isArray(names)) {
      throw new Error('Parâmetro names inválido');
    }
    
    if (!shabatDate) {
      throw new Error('Parâmetro shabatDate inválido');
    }
    
    var spreadsheetResult = saveNames(names, shabatDate);
    var presentationResult = createSlidesPresentation(names, shabatDate);
    
    return {
      success: true,
      spreadsheet: spreadsheetResult,
      presentation: presentationResult,
      message: spreadsheetResult.message
    };
  } catch (e) {
    Logger.log('Erro processForm: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
