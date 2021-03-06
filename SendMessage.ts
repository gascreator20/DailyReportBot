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
            if (cellValue === "") {
                SendMessage.changeWriteCountDetail(member, true);
            } else if (invalidTextByCellReportArray.indexOf(String(cellValue)) !== -1) {
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
     * @param {number} nextDayCount
     */
    public sendRequestNextDayPlanError(members: string[][], nextDayCount: number)
    {
        this.sendMessage(members, "作業予定記入要求エラーテンプレート", nextDayCount, this._config.workerRoomId);
    }
    
    /**
     * 作業予定計画の記入依頼（翌日分を参照）
     *
     * @param {string[][]} members
     */
    public sendRequestNextDayPlan(members: string[][])
    {
        this.sendMessage(members, "作業予定記入要求テンプレート", 1, this._config.workerRoomId);
    }
    
    /**
     * 作業報告
     *
     * @param {string[][]} members
     */
    public sendDailyTimeReport(members: string[][])
    {
        this.sendMessage(members, "作業報告成功テンプレート", 0, this._config.workerRoomId);
    }
    
    /**
     * 作業報告のエラー報告
     *
     * @param {string[][]} members
     */
    public sendDailyTimeReportError(members: string[][])
    {
        this.sendMessage(members, "作業報告のエラーテンプレート", 0, this._config.workerRoomId);
    }
    
    /**
     * 作業報告のルール違反者の報告
     *
     * @param {string[][]} members
     */
    public sendDailyTimeReportNgRule(members: string[][])
    {
        this.sendMessage(members, "作業報告のルール違反テンプレート", 0, this._config.workerRoomId);
    }
    
    /**
     * 作業報告の記入依頼
     *
     * @param {string[][]} members
     */
    public sendRequestDailyTimeReport(members: string[][])
    {
        this.sendMessage(members, "作業報告記入依頼テンプレート", 0, this._config.workerRoomId);
    }
    
    /**
     * 終業報告
     *
     * @param {string[][]} members
     */
    public sendEndOfWorkReport(members: string[][])
    {
        this.sendMessage(members, "終了報告テンプレート", 0, this._config.endOfWorkReportRoomId, false);
    }
    
    /**
     * 朝会報告
     *
     * @param {string[][]} members
     */
    public sendMorning(members: string[][])
    {
        this.sendMessage(members, "朝会のテンプレート", 0, this._config.morningRoomId, true, false);
    }
    
    /**
     * 自由ワード報告
     *
     * @param {string[][]} members
     */
    public freedomWording(members: string[][])
    {
        this.sendMessage(members, "自由発言のテンプレート", 0, this._config.freedomWordingRoomId, true, false);
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
            console.log("指定が正しくありません。json形式で入力して下さい")
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
            console.log("対象者はいませんでした")
            return;
        }
        
        // メッセージ送信
        const chatWork = new ChatWork();
        chatWork.sendMessage(message, this._config.roomIdByCellReport);
    }
    
    /**
     * シートのURLを取得する
     *
     * @param {string[]} member
     * @param {boolean} nextDayCount
     * @returns {string}
     */
    private getSheetURL(member: string[], nextDayCount: number): string
    {
        // 設定値の取得
        const keyValueSheetReader = new KeyValueSheet();
        const config = new Config();
        const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", config.calendarType);
        const calendar = keyValueSheetReader.find("カレンダー", this._config.calendarSheetKey, nowTime, nextDayCount, config.spreadsheetIdMember);
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
            message += member["名前"] + countMessage + "\n" + this.getSheetURL(member, 0) + "\n";
            message += cell["title"] + "\n";
            
            // 無効な記入の場合は強制的に「特になし」とする
            if (cellValue === "") {
                message += "特になし" + "\n\n";
            } else if (invalidTextByCellReportArray.indexOf(String(cellValue)) !== -1) {
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
     * @param {number} nextDayCount
     * @param {string} targetRoomId
     * @param {boolean} withToChat
     * @param {boolean} sheetUrl
     */
    private sendMessage(members: string[][], templateName: string, nextDayCount: number, targetRoomId: string, withToChat: boolean = true, sheetUrl: boolean = true)
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
                        // URLなしの時はグループの区切り箇所にのみ改行後スペースを入れる
                        if (message !== "" && sheetUrl === false) {
                            message += "\n";
                        }
                        message += this._config.emoji + key + this._config.emoji + "\n";
                        loopKeyName = key;
                    }
                    
                    message += this.getSendMessageText(withToChat, member, nextDayCount, sheetUrl);
                }
            }
        } else {
            for (const member of members) {
                message += this.getSendMessageText(withToChat, member, nextDayCount, sheetUrl);
            }
        }
        
        if (members.length >= 1) {
            message = template.replace("@", message);
        } else {
            console.log("対象者はいませんでした")
            return;
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
     * @param {number} nextDayCount
     * @param {boolean} sheetUrl
     * @returns {string}
     */
    private getSendMessageText(withToChat: boolean, member: string[], nextDayCount: number, sheetUrl: boolean)
    {
        let message = "";
        
        if (withToChat) {
            message += member["ChatWorkID"] + "\n";
            if (sheetUrl) {
                message += this.getSheetURL(member, nextDayCount) + "\n\n";
            }
        } else {
            message += member["名前"] + "\n";
            if (sheetUrl) {
                message += this.getSheetURL(member, nextDayCount) + "\n\n";
            }
        }
        
        return message;
    }
}
