// =========================================================================
// 연봉표 저장 (Google Sheets)
// =========================================================================

function addAllowanceColumn(allowanceName) {
  requireAdmin_();
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('연봉표');
    if (!sheet) return { success: false, message: '연봉표 시트가 없습니다.' };
    
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 이미 존재하는 열인지 확인
    if (headerRow.includes(allowanceName)) {
      return { success: true, message: '이미 존재하는 수당입니다.' };
    }
    
    // 새 열 추가 (마지막 열 다음)
    const newColIndex = sheet.getLastColumn() + 1;
    sheet.getRange(1, newColIndex).setValue(allowanceName);
    
    return { success: true, message: `${allowanceName} 열이 추가되었습니다.` };
  } catch (error) {
    return { success: false, message: '열 추가 실패: ' + error.message };
  }
}

function saveSalaryRecord(salaryRecordJson) {
  requireAdmin_();
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('연봉표');
    if (!sheet) return { success: false, message: '연봉표 시트가 없습니다.' };
    
    const record = JSON.parse(salaryRecordJson);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 해당 직원의 행 찾기 또는 새 행 추가
    let targetRow = null;
    const lastRow = sheet.getLastRow();
    
    for (let row = 2; row <= lastRow; row++) {
      const yearVal = sheet.getRange(row, 1).getValue();
      const empIdVal = sheet.getRange(row, 2).getValue();
      
      if (yearVal == record.year && empIdVal == record.empId) {
        targetRow = row;
        break;
      }
    }
    
    // 새 행 추가
    if (!targetRow) {
      targetRow = lastRow + 1;
    }
    
    // 기본 정보 입력 (A~E: 연도, 직원ID, 직원명, 급수, 호봉)
    sheet.getRange(targetRow, 1).setValue(record.year);
    sheet.getRange(targetRow, 2).setValue(record.empId);
    sheet.getRange(targetRow, 3).setValue(record.empName);
    sheet.getRange(targetRow, 4).setValue(record.grade);
    sheet.getRange(targetRow, 5).setValue(record.hobon);
    
    // F: 기본급
    sheet.getRange(targetRow, 6).setValue(record.basicSalary);
    
    // G: 정액급식비 (금액)
    sheet.getRange(targetRow, 7).setValue(record.mealSubsidy || 0);
    
    // H: 관리자수당 (금액)
    sheet.getRange(targetRow, 8).setValue(record.managerAllowance || 0);
    
    // I열부터: 기타 수당 (JSON 형식)
    if (record.otherAllowances && Object.keys(record.otherAllowances).length > 0) {
      for (let colIdx = 0; colIdx < headerRow.length; colIdx++) {
        const columnName = headerRow[colIdx];
        
        // I열(인덱스 8) 이후의 기타 수당 열
        if (colIdx >= 8 && record.otherAllowances[columnName]) {
          // JSON 형식으로 저장: {"name":"수당명","months":[...],"amount":...}
          const allowanceAmount = record.otherAllowances[columnName];
          
          // settingsData에서 months 정보 찾기 (record에 포함되어 있어야 함)
          const allowanceData = record.allowanceDetails && record.allowanceDetails[columnName];
          const months = allowanceData?.months || [];
          
          const jsonData = {
            name: columnName,
            months: months,
            amount: allowanceAmount
          };
          
          sheet.getRange(targetRow, colIdx + 1).setValue(JSON.stringify(jsonData));
        }
      }
    }
    
    return { success: true, message: '월급 기록이 저장되었습니다.' };
  } catch (error) {
    return { success: false, message: '저장 실패: ' + error.message };
  }
}
