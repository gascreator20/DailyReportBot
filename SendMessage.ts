import ChatWork from "./ChatWork";
import Common from "./Common";
import Config from "./Config";
import KeyValueSheet from "./KeyValueSheet";
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

/**
 * メッセージ送信に関する操作を行うためのクラス
 */
export default class SendMessage
{
    public _config;
    
    /**
     * constructor
     */
    constructor()
    {
        this._config = new Config();
    }
    
    /**
     * 記入数のカウントを増やす
     *
     * @param {string[]} member
     * @param {boolean} isEmptyMessage
     */
    private static changeWriteCountDetail(member: string[], isEmptyMessage: boolean)
    {
        const config = new Config();
        const sheet = new KeyValueSheet();
        const columnNumber: number = sheet.getColumnNumber("メンバー管理", "記入数");
        const rowNumber: number = sheet.getRowNumber("メンバー管理", member["名前"], "名前");
        
        if (config.countTypeByCellReport === "累計カウント") {
            if (!isEmptyMessage) {
                sheet.write("メンバー管理", rowNumber, columnNumber, String(Number(member["記入数"]) + 1));
            }
        }
        
        if (config.countTypeByCellReport === "リセットカウント") {
            if (isEmptyMessage) {
                sheet.write("メンバー管理", rowNumber, columnNumber, "0");
            } else {
                sheet.write("メンバー管理", rowNumber, columnNumber, String(Number(member["記入数"]) + 1));
            }
        }
    }
    
    /**
     * 記入回数のカウントを計測する
     *
     * @param {string[]} member
     * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sheet
     */
    public changeWriteCount(member: string[], sheet: Sheet)
    {
        const config = new Config();
        const invalidTextByCellReport = config.invalidTextByCellReport;
        const invalidTextByCellReportArray = invalidTextByCellReport.split(',');
        
        if (config.countTypeByCellReport === "カウントしない" || sheet === null) {
            return;
        }
        
        const cells = JSON.parse(config.cellNumberByCellReport);
        for (const cell of cells["output"]) {
            const cellValue = sheet.getRange(cell["cellNumber"]).getValue();
            if (cellValue === "" || invalidTextByCellReportArray.includes(String(cellValue))) {
                SendMessage.changeWriteCountDetail(member, true);
            } else {
                SendMessage.changeWriteCountDetail(member, false);
            }
        }
    }
    
    /**
     * 作業予定計画のエラーチェック（翌日分を参照）
     *
     * @param {string[][]} members
     * @param {boolean} isNextDay
     */
    public sendRequestNextDayPlanError(members: string[][], isNextDay: boolean)
    {
        this.sendMessage(members, "作業予定記入要求エラーテンプレート", isNextDay, this._config.workerRoomId);
    }
    
    /**
     * 作業予定計画の記入依頼（翌日分を参照）
     *
     * @param {string[][]} members
     */
    public sendRequestNextDayPlan(members: string[][])
    {
        this.sendMessage(members, "作業予定記入要求テンプレート", true, this._config.workerRoomId);
    }
    
    /**
     * 作業報告
     *
     * @param {string[][]} members
     */
    public sendDailyTimeReport(members: string[][])
    {
        this.sendMessage(members, "作業報告成功テンプレート", false, this._config.workerRoomId);
    }
    
    /**
     * 作業報告のエラー報告
     *
     * @param {string[][]} members
     */
    public sendDailyTimeReportError(members: string[][])
    {
        this.sendMessage(members, "作業報告のエラーテンプレート", false, this._config.workerRoomId);
    }
    
    /**
     * 作業報告のルール違反者の報告
     *
     * @param {string[][]} members
     */
    public sendDailyTimeReportNgRule(members: string[][])
    {
        this.sendMessage(members, "作業報告のルール違反テンプレート", false, this._config.workerRoomId);
    }
    
    /**
     * 作業報告の記入依頼
     *
     * @param {string[][]} members
     */
    public sendRequestDailyTimeReport(members: string[][])
    {
        this.sendMessage(members, "作業報告記入依頼テンプレート", false, this._config.workerRoomId);
    }
    
    /**
     * 終業報告
     *
     * @param {string[][]} members
     */
    public sendEndOfWorkReport(members: string[][])
    {
        this.sendMessage(members, "終了報告テンプレート", false, this._config.workerRoomId, false);
    }
    
    /**
     * 指定セル番号の内容を送る
     *
     * @param {string[][]} members
     * @param {string[]} calendar
     * @param {string} templateName
     * @returns {string}
     */
    public sendCellContent(members: string[][], calendar: string[], templateName: string)
    {
        const config = new Config();
        const sortKey = config.reportSortKey;
        const keyValueSheetReader = new KeyValueSheet();
        const configTemplate = keyValueSheetReader.find("設定", "キー名", templateName);
        const template = configTemplate["値"];
        let message = "";
        
        if (!config.cellNumberByCellReport || config.cellNumberByCellReport === "") {
            return "指定が正しくありません。json形式で入力して下さい";
        }
        
        // 報告の順番をソートするかどうか
        if (sortKey) {
            let loopKeyName = "";
            const membersGroup = Common.groupBy(members, sortKey);
            for (const key in membersGroup) {
                if (!membersGroup.hasOwnProperty(key)) {
                    continue;
                }
                
                for (const member of membersGroup[key]) {
                    // グループ名を最初に記載する
                    if (loopKeyName !== key) {
                        message += this._config.emoji + key + this._config.emoji + "\n";
                        loopKeyName = key;
                    }
                    
                    message += this.getTextBySendCellContent(member, calendar);
                }
            }
        } else {
            for (const member of members) {
                message += this.getTextBySendCellContent(member, calendar);
            }
        }
        
        if (members.length >= 1) {
            message = template.replace("@", message);
        } else {
            message = template.replace("@", "該当者なし");
        }
        
        // メッセージ送信
        const chatWork = new ChatWork();
        chatWork.sendMessage(message, this._config.roomIdByCellReport);
    }
    
    /**
     * シートのURLを取得する
     *
     * @param {string[]} member
     * @param {boolean} isNextDay
     * @returns {string}
     */
    private getSheetURL(member: string[], isNextDay: boolean): string
    {
        // 設定値の取得
        const keyValueSheetReader = new KeyValueSheet();
        const config = new Config();
        const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", config.calendarType);
        const calendar = keyValueSheetReader.find("カレンダー", this._config.calendarSheetKey, nowTime, isNextDay);
        const sheetByMember = SpreadsheetApp.openById(member["SpreadsheetID"]).getSheetByName(calendar[this._config.calendarSheetKey]);
        
        if (!sheetByMember) {
            return "シートが見つかりませんでした";
        }
        
        return member["SheetURL"] + "#gid=" + sheetByMember.getSheetId()
    }
    
    /**
     * セル内容の送信を行う際に使用する本文を取得する
     *
     * @param {string[]} member
     * @param {string} calendar
     * @returns {string}
     */
    private getTextBySendCellContent(member: string[], calendar: string[]): string
    {
        const sheetByMember = SpreadsheetApp.openById(member["SpreadsheetID"]).getSheetByName(calendar[this._config.calendarSheetKey]);
        
        if (sheetByMember) {
            return this.getCellMessage(member, sheetByMember);
        } else {
            return member["名前"] + "\n" + "シートが見つかりませんでした" + "\n\n";
        }
    }
    
    /**
     * 指定セルの内容を一括送信する用のメッセージを取得
     *
     * @param {string[]} member
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
     * @returns {string}
     */
    private getCellMessage(member: string[], sheet: Sheet)
    {
        let message = "";
        const config = new Config();
        const cells = JSON.parse(config.cellNumberByCellReport);
        const invalidTextByCellReport = config.invalidTextByCellReport;
        const invalidTextByCellReportArray = invalidTextByCellReport.split(',');
        
        for (const cell of cells["output"]) {
            let countMessage = "";
            if (config.countTypeByCellReport !== "カウントしない") {
                countMessage = "：記入数" + member["記入数"] + "回";
            }
            
            const cellValue = sheet.getRange(cell["cellNumber"]).getValue();
            message += member["名前"] + countMessage + "\n" + this.getSheetURL(member, false) + "\n";
            message += cell["title"] + "\n";
            if (cellValue === "" || invalidTextByCellReportArray.includes(String(cellValue))) {
                message += "特になし" + "\n\n";
            } else {
                message += cellValue + "\n\n";
            }
        }
        
        return message;
    }
    
    /**
     * メッセージを送る
     *
     * @param {string[][]} members
     * @param {string} templateName
     * @param {boolean} isNextDay
     * @param {string} targetRoomId
     * @param {boolean} withToChat
     */
    private sendMessage(members: string[][], templateName: string, isNextDay: boolean, targetRoomId: string, withToChat: boolean = true)
    {
        // 設定の取得
        const config = new Config();
        const sortKey = config.reportSortKey;
        
        // テンプレートの取得
        const keyValueSheetReader = new KeyValueSheet();
        const configTemplate = keyValueSheetReader.find("設定", "キー名", templateName);
        const template = configTemplate["値"];
        let message: string = "";
        
        // 報告の順番をソートするかどうか
        if (sortKey) {
            let loopKeyName = "";
            const membersGroup = Common.groupBy(members, sortKey);
            for (const key in membersGroup) {
                if (!membersGroup.hasOwnProperty(key)) {
                    continue;
                }
                
                for (const member of membersGroup[key]) {
                    // グループ名を最初に記載する
                    if (loopKeyName !== key) {
                        message += this._config.emoji + key + this._config.emoji + "\n";
                        loopKeyName = key;
                    }
                    
                    message += this.getSendMessageText(withToChat, member, isNextDay);
                }
            }
        } else {
            for (const member of members) {
                message += this.getSendMessageText(withToChat, member, isNextDay);
            }
        }
        
        if (members.length >= 1) {
            message = template.replace("@", message);
        } else {
            message = template.replace("@", "該当者なし");
        }
        
        // メッセージ送信
        const chatWork = new ChatWork();
        chatWork.sendMessage(message, targetRoomId);
    }
    
    /**
     * 送信本文を取得する
     *
     * @param {boolean} withToChat
     * @param {string[]} member
     * @param {boolean} isNextDay
     * @returns {string}
     */
    private getSendMessageText(withToChat: boolean, member: string[], isNextDay: boolean)
    {
        let message = "";
        
        if (withToChat) {
            message += member["ChatWorkID"] + "\n" + this.getSheetURL(member, isNextDay) + "\n\n";
        } else {
            message += member["名前"] + "\n" + this.getSheetURL(member, isNextDay) + "\n\n";
        }
        
        return message;
    }
}

