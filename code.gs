function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Lock System')
    .addItem('Thiết lập hệ thống', 'setupSystem')
    .addToUi();
}

function setupSystem() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Thêm cột người tạo và trạng thái nếu chưa có
  var lastCol = sheet.getLastColumn();
  sheet.getRange(9, lastCol + 1).setValue("Email tạo");
  sheet.getRange(9, lastCol + 2).setValue("Trạng thái");
  
  // Tạo data validation cho cột trạng thái
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Chờ duyệt', 'Đã duyệt'])
    .build();
  
  var statusRange = sheet.getRange(10, lastCol + 2, sheet.getMaxRows() - 1, 1);
  statusRange.setDataValidation(rule);
  
  // Thiết lập trigger để theo dõi thay đổi
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var statusCol = sheet.getLastColumn(); // Cột trạng thái
  var creatorCol = statusCol - 1; // Cột người tạo
  
  if (range.getColumn() < statusCol) {
    var row = range.getRow();
    var creatorCell = sheet.getRange(row, creatorCol);
    var statusCell = sheet.getRange(row, statusCol);
    
    // Ghi email người tạo nếu chưa có
    if (creatorCell.getValue() === '') {
      creatorCell.setValue(Session.getActiveUser().getEmail());
    }
  }
  
  // Xử lý khi trạng thái được chỉnh sửa
  if (range.getColumn() == statusCol) {
    var row = range.getRow();
    var rowRange = sheet.getRange(row, 1, 1, statusCol);
    var newValue = range.getValue();
    var creatorEmail = sheet.getRange(row, creatorCol).getValue();
    var ownerEmail = SpreadsheetApp.getActive().getOwner().getEmail();
    var currentUser = Session.getActiveUser().getEmail();
    
    // Xóa mọi protection hiện tại
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
      var protection = protections[i];
      if (protection.getRange().getRow() == row) {
        protection.remove();
      }
    }
    
    if (newValue === 'Chờ duyệt') {
      // Tạo protection cho người tạo và chủ sở hữu
      var protection = rowRange.protect();
      protection.removeEditors(protection.getEditors());
      protection.addEditor(creatorEmail);
      protection.addEditor(ownerEmail);
      protection.setDescription('Chỉ người tạo và chủ sở hữu có thể chỉnh sửa - Pending');
    } 
    else if (newValue === 'Đã duyệt') {
      // Kiểm tra quyền để chuyển sang Confirmed
      if (currentUser !== ownerEmail) {
        range.setValue('Pending');
        SpreadsheetApp.getUi().alert('Chỉ chủ sở hữu mới có thể xác nhận trạng thái.');
        return;
      }
      
      // Tạo protection chỉ cho chủ sở hữu
      var protection = rowRange.protect();
      protection.removeEditors(protection.getEditors());
      protection.addEditor(ownerEmail);
      protection.setDescription('Chỉ chủ sở hữu có thể chỉnh sửa - Confirmed');
    }
  }
}
