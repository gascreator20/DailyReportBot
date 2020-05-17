import Common from "./Common";
import Config from "./Config";

/**
 * 縦横の基本的なスプレッドシートの情報を取得するクラス
 */
export default class KeyValueSheet
{
    /**
     * @type {string[][][]}
     */
    protected static sheetDataCacheWithConvertBySheetName = [];
    
    /**
     * @type {string[][][]}
     */
    protected static sheetDataCacheBySheetName = [];
    
    /**
     * @type {number[]}
     */
    protected static rowcountCacheBySheetName = [];
    
    /**
     * @type {number[]}
     */
    protected static columnPositionNumberCache = [];
    
    /**
     * @type {string[][]}
     */
    protected static sheetColumnNameCacheBySheetName = [];
    
    /**
     * シートデータを読み込む
     *
     * @param {string} sheetName
     * @param {string} spreadsheetId
     */
    private static createLocalCache(sheetName: string, spreadsheetId: string = null)
    {
        let sheet;
        if (spreadsheetId === null) {
            sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        } else {
            sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
        }
        
        // 該当シートの縦横すべてを取得
        KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName] = Common.convertSheet(sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn()));
        KeyValueSheet.sheetDataCacheBySheetName[sheetName] = sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn());
        
        // 該当シートのヘッダー情報を取得
        KeyValueSheet.sheetColumnNameCacheBySheetName[sheetName] = sheet.getSheetValues(1, 1, 1, sheet.getLastColumn())[0];
        
        // 該当シートの行数を取得
        KeyValueSheet.rowcountCacheBySheetName[sheetName] = sheet.getLastRow() - 1;
    }
    
    /**
     * シートデータのキャッシュを破棄する
     */
    public flushLocalCache()
    {
        // 該当シートの縦横すべてを取得
        KeyValueSheet.sheetDataCacheWithConvertBySheetName = [];
        KeyValueSheet.sheetDataCacheBySheetName = [];
        
        // 該当シートのヘッダー情報を取得
        KeyValueSheet.sheetColumnNameCacheBySheetName = [];
        
        // 該当シートの行数を取得
        KeyValueSheet.rowcountCacheBySheetName = [];
    }
    
    /**
     * 対象のシートデータから指定の値を取得する
     *
     * @param {string} sheetName
     * @param {string} columnName
     * @param {string} targetName
     * @param {boolean} isNextRow
     * @param {boolean} spreadsheetId
     * @returns {string[]}
     */
    public find(sheetName: string, columnName: string, targetName: string, isNextRow: boolean = false, spreadsheetId: string = null): string[]
    {
        // 既にキャッシュ済みの場合は作らない（一度読み込んだものはキャッシュしておく）
        if (!KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName]) {
            KeyValueSheet.createLocalCache(sheetName, spreadsheetId);
        }
        
        for (let i = 0; i < KeyValueSheet.rowcountCacheBySheetName[sheetName]; i++) {
            if (String(KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName][i][columnName]) === targetName) {
                if (!isNextRow) {
                    return KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName][i];
                }
                
                if (KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName][i + 1]) {
                    return KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName][i + 1];
                } else {
                    console.log('参照データが見つかりませんでした');
                }
            }
        }
        
        console.log('参照データが見つかりませんでした');
        return null;
    }
    
    /**
     * 対象のシートデータから指定の値を取得する
     *
     * @param {string} sheetName
     * @param {string} columnName
     * @param {string} targetName
     * @returns {string[]}
     */
    public get(sheetName: string, columnName: string, targetName: string): string[][]
    {
        const result = [];
        
        // 既にキャッシュ済みの場合は作らない（一度読み込んだものはキャッシュしておく）
        if (!KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName]) {
            KeyValueSheet.createLocalCache(sheetName);
        }
        
        for (let i = 0; i < KeyValueSheet.rowcountCacheBySheetName[sheetName]; i++) {
            if (KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName][i][columnName] === targetName) {
                result.push(KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName][i]);
            }
        }
        
        return result;
    }
    
    /**
     * 該当シートの全情報を取得する（1行目は配列のキーとなる）
     *
     * @param {string} sheetName
     * @param {string} spreadsheetId
     * @returns {string[][]}
     */
    public getAll(sheetName: string, spreadsheetId: string): string[][]
    {
        // 既に作成済みの場合は作らない（一度読み込んだものはキャッシュしておく）
        if (!KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName]) {
            KeyValueSheet.createLocalCache(sheetName, spreadsheetId);
        }
        
        return KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName];
    }
    
    /**
     * 該当シートの特定の列のデータをすべて取得する
     *
     * @param {string} sheetName
     * @param {string} columnName
     * @param {string} spreadsheetId
     * @returns {string[]}
     */
    public getAllByColumnName(sheetName: string, columnName: string, spreadsheetId: string): string[]
    {
        // 既に作成済みの場合は作らない（一度読み込んだものはキャッシュしておく）
        if (!KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName]) {
            KeyValueSheet.createLocalCache(sheetName, spreadsheetId);
        }
        
        const results = [];
        for (const value of KeyValueSheet.sheetDataCacheWithConvertBySheetName[sheetName]) {
            results.push(value[columnName]);
        }
        
        return results;
    }
    
    /**
     * 該当シートの列番号を取得する
     *
     * @param {string} sheetName
     * @param {string} columnName
     * @returns {number}
     */
    public getColumnNumber(sheetName: string, columnName: string): number
    {
        const cacheKey = sheetName + "_CacheKeyColumnName_" + columnName;
        
        if (KeyValueSheet.columnPositionNumberCache[cacheKey]) {
            return KeyValueSheet.columnPositionNumberCache[cacheKey];
        }
        
        // 既に作成済みの場合は作らない（一度読み込んだものはキャッシュしておく）
        if (!KeyValueSheet.sheetColumnNameCacheBySheetName[sheetName]) {
            KeyValueSheet.createLocalCache(sheetName);
        }
        
        for (let i = 0; i < KeyValueSheet.sheetColumnNameCacheBySheetName[sheetName].length; i++) {
            if (columnName === KeyValueSheet.sheetColumnNameCacheBySheetName[sheetName][i]) {
                // カラム位置は1番はじまりのため1つ分ずらす
                KeyValueSheet.columnPositionNumberCache[cacheKey] = i + 1;
                return i + 1;
            }
        }
    }
    
    /**
     * 該当シートの行番号を取得する
     *
     * @param {string} sheetName
     * @param {string} targetName
     * @param {string} columnName
     * @returns {number}
     */
    public getRowNumber(sheetName: string, targetName: string, columnName: string): number
    {
        const columnNumber = this.getColumnNumber(sheetName, columnName);
        
        // 既に作成済みの場合は作らない（一度読み込んだものはキャッシュしておく）
        if (!KeyValueSheet.sheetDataCacheBySheetName[sheetName]) {
            KeyValueSheet.createLocalCache(sheetName);
        }
        
        for (const rowNumber in KeyValueSheet.sheetDataCacheBySheetName[sheetName]) {
            if (!KeyValueSheet.sheetDataCacheBySheetName[sheetName].hasOwnProperty(rowNumber)) {
                continue;
            }
            
            // 配列は0番はじまりのため-1する
            if (KeyValueSheet.sheetDataCacheBySheetName[sheetName][rowNumber][columnNumber - 1] === targetName) {
                // 行位置は1番はじまりのため1つ分ずらす
                return Number(rowNumber) + 1;
            }
        }
    }
    
    /**
     * 該当シートのセルに書き込む
     *
     * @param {string} sheetName
     * @param {number} rowNumber
     * @param {number} columnNumber
     * @param {string} message
     */
    public write(sheetName: string, rowNumber: number, columnNumber: number, message: string)
    {
        const config = new Config();
        const sheet = SpreadsheetApp.openById(config.spreadsheetIdManagement).getSheetByName(sheetName);
        
        sheet.getRange(rowNumber, columnNumber).setValue(message);
    }
}

