import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

/**
 * スプレッドシートのコピー操作するクラス
 */
export default class SpreadSheetCreator
{
    /**
     * シートのコピーを行う
     *
     * @param {GoogleAppsScript.Spreadsheet.Sheet} oldSheet
     * @param {string} targetSheetId
     * @param {string} newSheetName
     * @param {boolean} isMoveTop
     */
    private static copySheet(oldSheet: Sheet, targetSheetId: string, newSheetName: string, isMoveTop: boolean = true)
    {
        const spreadsheet = SpreadsheetApp.openById(targetSheetId);
        
        // すでに該当のシートがある場合は何もしない
        const sheet = spreadsheet.getSheetByName(newSheetName);
        if (sheet) {
            return;
        }
        
        const newSheet = oldSheet.copyTo(spreadsheet);
        spreadsheet.setActiveSheet(newSheet);
        newSheet.setName(newSheetName);
        
        if (isMoveTop) {
            spreadsheet.moveActiveSheet(1);
        }
    }
    
    /**
     * シートのコピーを行う
     *
     * @param {string} driveDirectoryId
     * @param {GoogleAppsScript.Spreadsheet.Sheet} templateSheet
     * @param {string} newSheetName
     */
    public copy(driveDirectoryId: string, templateSheet: Sheet, newSheetName: string)
    {
        const files = DriveApp.getFolderById(driveDirectoryId).getFiles();
        while (files.hasNext()) {
            SpreadSheetCreator.copySheet(templateSheet, files.next().getId(), newSheetName);
        }
    }
}

