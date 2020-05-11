import Common from "./Common";
import Config from "./Config";

/**
 * 日報シートを操作するクラス
 */
export default class DailyReportSheetReader
{
    /**
     * 該当者の指定offset時間帯の作業内容を取得
     *
     * @param {string[]} member
     * @param {string[]} calendar
     * @param {number} offsetRowCount
     * @returns {string[]}
     */
    private static findWorkReport(member: string[], calendar: string[], offsetRowCount: number): string[]
    {
        const config = new Config();
        
        // スケジュールの開始日時と終了日時のタイムスタンプを取得
        const [startTimeBySchedule, endTimeBySchedule] = Common.convertTime(calendar[config.calendarSheetKey], calendar[member["名前"]]);
        
        // シートの縦横すべてを取得
        const sheetByMember = SpreadsheetApp.openById(member["SpreadsheetID"]).getSheetByName(calendar[config.calendarSheetKey]);
        const sheetData = sheetByMember.getSheetValues(config.startRow, 1, config.rangeRow, 7);
        
        // 現在の時刻を取得
        const date = new Date();
        const timeStamp = new Date(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), 0).getTime();
        
        // スケジュールの開始日時の作業予定を取得（全時間帯をループ）
        for (let i = 0; i < config.rangeRow; i++) {
            // 該当時間帯の開始日時と終了日時のタイムスタンプを取得
            const [startTimeByEachTime, endTimeByEachTime] = Common.convertTime(calendar[config.calendarSheetKey], sheetData[i][0] + sheetData[i][1] + sheetData[i][2]);
            
            // 現在の時間帯の情報の判定
            if (timeStamp >= startTimeByEachTime && timeStamp <= endTimeByEachTime) {
                // 報告対象の時間帯の勤務があったかどうかを確認
                const [startTimeBeforeByEachTime, endTimeBeforeByEachTime] = Common.convertTime(
                    calendar[config.calendarSheetKey],
                    sheetData[i + offsetRowCount][0] + sheetData[i + offsetRowCount][1] + sheetData[i + offsetRowCount][2]
                );
                
                // 報告時刻の勤務がなければ報告対象には含めない
                if (startTimeBySchedule > endTimeBeforeByEachTime) {
                    console.log(member["名前"] + "はまだ勤務していません");
                    return null;
                }
                
                // 報告内容を返す
                return sheetData[i + offsetRowCount];
            }
        }
        
        console.log(member["名前"] + "の時間帯情報を取得できませんでした:" + timeStamp);
        
        return null;
    }
    
    /**
     * 該当者の作業開始時間帯の作業内容を取得
     *
     * @param {string[]} member
     * @param {string[]} calendar
     * @returns {string[]}
     */
    public findDailyReportByWorkStartTime(member: string[], calendar: string[]): string[]
    {
        const config = new Config();
        
        // スケジュールの開始日時と終了日時のタイムスタンプを取得
        const [startTimeBySchedule, endTimeBySchedule] = Common.convertTime(calendar[config.calendarSheetKey], calendar[member["名前"]]);
        
        // シートの縦横すべてを取得
        const sheetByMember = SpreadsheetApp.openById(member["SpreadsheetID"]).getSheetByName(calendar[config.calendarSheetKey]);
        const sheetData = sheetByMember.getSheetValues(config.startRow, 1, config.rangeRow, 7);
        
        // スケジュールの開始日時の作業予定を取得（全時間帯をループ）
        for (let i = 0; i < config.rangeRow; i++) {
            // 該当時間帯の開始日時と終了日時のタイムスタンプを取得
            const [startTimeByEachTime, endTimeByEachTime] = Common.convertTime(calendar[config.calendarSheetKey], sheetData[i][0] + sheetData[i][1] + sheetData[i][2]);
            
            // 一致する箇所の作業予定
            if (startTimeBySchedule <= endTimeByEachTime) {
                return sheetData[i];
            }
        }
        
        console.log(member["名前"] + "の時間帯情報を取得できませんでした:" + startTimeBySchedule);
        
        return null;
    }
    
    /**
     * 該当者の作業報告情報を取得する
     *
     * @param {string[]} member
     * @param {string[]} calendar
     * @returns {string[]}
     */
    public findWorkResultByNowTime(member: string[], calendar: string[]): string[]
    {
        return DailyReportSheetReader.findWorkReport(member, calendar, 0);
    }
    
    /**
     * 該当者の作業報告情報を取得する
     *
     * @param {string[]} member
     * @param {string[]} calendar
     * @returns {string[]}
     */
    public findWorkResultByBeforeTime(member: string[], calendar: string[]): string[]
    {
        return DailyReportSheetReader.findWorkReport(member, calendar, -1);
    }
}