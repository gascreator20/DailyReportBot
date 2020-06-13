import Config from "./Config";
import DailyReportSheetReader from "./DailyReportSheetReader";
import KeyValueSheet from "./KeyValueSheet";

/**
 * 作業者に関する情報を操作するクラス
 */
export default class Worker
{
    /**
     * 現時刻の作業メンバーの情報を取得する
     *
     * @return {string[][]}
     */
    public getWorkMemberNowTime(): string[][]
    {
        const config = new Config();
        const keyValueSheetReader = new KeyValueSheet();
        const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", config.calendarType);
        const calendar = keyValueSheetReader.find("カレンダー", config.calendarSheetKey, nowTime);
        const members = this.getWorkMember();
        const joinMembers = [];
    
        for (const member of members) {
            // 勤務時間外の場合はnullが返る。nullの場合は何も作業していないので作業報告者対象から除外する
            const spreadSheetReader = new DailyReportSheetReader();
            const reportNowTime = spreadSheetReader.findWorkResultByNowTime(member, calendar);
            if (reportNowTime === null) {
                continue;
            }
        
            joinMembers.push(member);
        }
    
        return joinMembers;
    }
    
    /**
     * 作業メンバーの情報を取得する
     *
     * @param {boolean} isNextDay 次の営業日のデータを対象とするかどうか
     * @return {string[][]}
     */
    public getWorkMember(isNextDay: boolean = false): string[][]
    {
        const result = [];
        
        // 設定値の取得
        const keyValueSheetReader = new KeyValueSheet();
        const config = new Config()
        
        // 営業日を取得
        const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", config.calendarType);
        const calendar = keyValueSheetReader.find("カレンダー", config.calendarSheetKey, nowTime, isNextDay);
        
        // 対象外となるメンバーを除外する
        const rowsMemberList = keyValueSheetReader.getAll("メンバーリスト", config.spreadsheetIdMember);
        for (const rowsMember of rowsMemberList) {
            const rowsMemberManagement = keyValueSheetReader.find("メンバー管理", "名前", rowsMember["名前"], false, config.spreadsheetIdManagement);
            
            if (rowsMemberManagement === null) {
                console.log(rowsMember["名前"] + "の情報を管理メンバーシートで参照できませんでした");
                continue;
            }
            
            if (calendar[rowsMember["名前"]] === "休み" || calendar[rowsMember["名前"]] === "") {
                console.log(rowsMember["名前"] + "はお休みです");
                continue;
            }
            if (rowsMember["SpreadsheetID"] === "null" || rowsMember["SpreadsheetID"] === "") {
                console.log(rowsMember["名前"] + "のSpreadsheetIDを取得できませんでした");
                continue;
            }
            if (rowsMember["SheetURL"] === "null" || rowsMember["SheetURL"] === "") {
                console.log(rowsMember["名前"] + "のSheetURLを取得できませんでした");
                continue;
            }
            if (rowsMember["通知対象"] !== "TRUE" && rowsMember["通知対象"] !== true) {
                console.log(rowsMember["名前"] + "は通知対象外です");
                continue;
            }
            
            rowsMember["SpreadsheetID"] = rowsMemberManagement["SpreadsheetID"];
            rowsMember["SheetURL"] = rowsMemberManagement["SheetURL"];
            rowsMember["記入数"] = rowsMemberManagement["記入数"];
            
            result.push(rowsMember);
        }
        
        return result;
    }
}

