/* tslint:disable:only-arrow-functions */
/**
 * 共通化の便利関数
 */
export default class Common
{
    /**
     * 2次元配列を連想配列に変換する（最初の行をキーとする）
     *
     * @param values 二次元配列
     * @return string[][] 連想配列化したオブジェクト
     */
    public static convertSheet(values: any[][])
    {
        const keys = values.splice(0, 1)[0];

        return values.map(function (row) {
            const object = [];
            row.map(function (column, index) {
                object[keys[index]] = column;
            });

            return object;
        });
    }
    
    /**
     * 10:00-18:00という文字列を開始と終了のタイムスタンプ（ミリ秒に変換する）
     *
     * @param date 年月日を示す文字列（例：20200504）
     * @param time 開始時刻、終了時刻を示す文字列（例：10:00-19:00）
     */
    public static convertTime(date: string, time: string)
    {
        date = String(date);
        time = String(time);
    
        const [startTime, endTime] = time.split('-');
        const [startHour, startMinute] = startTime.split(':');
        const [endHour, endMinute] = endTime.split(':');
        const year = date.substr(0, 4);
        const month = date.substr(4, 2);
        const day = date.substr(6, 2);
        
        // 月の値は0～11(1月～12月)
        const startDate = new Date(Number(year), Number(month) - 1, Number(day), Number(startHour), Number(startMinute), 0);
        const endDate = new Date(Number(year), Number(month) - 1, Number(day), Number(endHour), Number(endMinute), 0);
        
        return [startDate.getTime(), endDate.getTime()];
    }
    
    /**
     * 二次元配列をグループ化する
     *
     * @param {string[][]} values
     * @param {string} sortKey
     * @returns {string[][][]}
     */
    public static groupBy(values: string[][], sortKey: string)
    {
        const result = [];
        
        for (const value of values) {
            if (!result[value[sortKey]]) {
                result[value[sortKey]] = [];
            }
            result[value[sortKey]].push(value);
        }
        
        return result;
    }
}

