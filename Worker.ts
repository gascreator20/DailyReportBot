import Config from "./Config";
import KeyValueSheet from "./KeyValueSheet";

/**
 * 作業者に関する情報を操作するクラス
 */
export default class Worker
{
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
        const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
        const calendar = keyValueSheetReader.find("カレンダー", config.calendarSheetKey, nowTime, isNextDay);
        
        // 対象外となるメンバーを除外する
        const rowsMemberList = keyValueSheetReader.getAll("メンバーリスト");
        for (const rowsMember of rowsMemberList) {
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
            
            result.push(rowsMember);
        }
        
        return result;
    }
}

