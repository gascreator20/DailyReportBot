import KeyValueSheet from "./KeyValueSheet";

/**
 * 設定用ファイル
 */
export default class Config
{
    private readonly _calendarSheetKey;
    private readonly _driveDirectoryId;
    private readonly _prefix;
    private readonly _apiToken;
    private readonly _startRow;
    private readonly _rangeRow;
    private readonly _testRoomId;
    private readonly _workerRoomId;
    private readonly _managementRoomId;
    private readonly _template;
    private readonly _isTest;
    private readonly _requestWriteTimes;
    private readonly _reportTime;
    private readonly _reportRetryTime;
    private readonly _isReportRetryTime;
    private readonly _endOfWorkReportTime;
    private readonly _endOfWorkReportRoomId;
    private readonly _reportSortKey;
    private readonly _roomIdByCellReport;
    private readonly _cellNumberByCellReport;
    private readonly _countTypeByCellReport;
    private readonly _invalidTextByCellReport;
    private readonly _templateCreateTime;
    private readonly _checkTodayPlanTime;
    private readonly _checkNextPlanTime;
    private readonly _requestNextPlanTime;
    private readonly _reportCellTime;
    
    constructor()
    {
        // 設定値の取得
        const keyValueSheetReader = new KeyValueSheet();
        this._driveDirectoryId = keyValueSheetReader.find("設定", "キー名", "ディレクトリのID")["値"];
        this._calendarSheetKey = keyValueSheetReader.find("設定", "キー名", "カレンダーシートのキー名")["値"];
        this._prefix = keyValueSheetReader.find("設定", "キー名", "スプレッドシートのファイル名の接頭語")["値"];
        this._startRow = keyValueSheetReader.find("設定", "キー名", "日報の時刻開始行番号")["値"];
        this._rangeRow = keyValueSheetReader.find("設定", "キー名", "日報の時刻表示行数")["値"];
        this._apiToken = keyValueSheetReader.find("設定", "キー名", "ApiToken")["値"];
        this._isReportRetryTime = keyValueSheetReader.find("設定", "キー名", "エラーの再検知を行うかどうか")["値"];
        this._reportSortKey = keyValueSheetReader.find("設定", "キー名", "報告のソートに使用するキー")["値"];
        
        // ルーム関連
        this._testRoomId = keyValueSheetReader.find("設定", "キー名", "テスト用ルームID")["値"];
        this._workerRoomId = keyValueSheetReader.find("設定", "キー名", "作業者ルームID")["値"];
        this._endOfWorkReportRoomId = keyValueSheetReader.find("設定", "キー名", "終業報告場所のルームID")["値"];
        this._isTest = keyValueSheetReader.find("設定", "キー名", "isTest")["値"];
        
        // 指定セルの報告用
        this._roomIdByCellReport = keyValueSheetReader.find("設定", "キー名", "指定セル報告の報告場所ルームID")["値"];
        this._cellNumberByCellReport = keyValueSheetReader.find("設定", "キー名", "指定セルの一斉報告の対象")["値"];
        this._countTypeByCellReport = keyValueSheetReader.find("設定", "キー名", "指定セルの発言内容をカウントする際の種別")["値"];
        this._invalidTextByCellReport = keyValueSheetReader.find("設定", "キー名", "指定セルの報告において、無効とするテキスト")["値"];
        
        // トリガー関連
        this._requestWriteTimes = keyValueSheetReader.find("設定", "キー名", "作業報告記入時間帯")["値"];
        this._reportTime = keyValueSheetReader.find("設定", "キー名", "作業報告記入から何分後に検知するか")["値"];
        this._reportRetryTime = keyValueSheetReader.find("設定", "キー名", "作業報告でエラーだった場合の再検知を何分後に行うか")["値"];
        this._endOfWorkReportTime = keyValueSheetReader.find("設定", "キー名", "終業報告トリガー発動時間")["値"];
        this._templateCreateTime = keyValueSheetReader.find("設定", "キー名", "次回分の日報を自動生成する時間")["値"];
        this._checkTodayPlanTime = keyValueSheetReader.find("設定", "キー名", "当日分の作業予定の記入が正しいかチェックする時間")["値"];
        this._checkNextPlanTime = keyValueSheetReader.find("設定", "キー名", "次回分の作業予定の記入が正しいかチェックする時間")["値"];
        this._requestNextPlanTime = keyValueSheetReader.find("設定", "キー名", "次回分の作業予定記入依頼の時間")["値"];
        this._reportCellTime = keyValueSheetReader.find("設定", "キー名", "指定セル報告を行う時間")["値"];
    }
    
    public get templateCreateTime()
    {
        return this._templateCreateTime;
    }
    
    public get checkTodayPlanTime()
    {
        return this._checkTodayPlanTime;
    }
    
    public get checkNextPlanTime()
    {
        return this._checkNextPlanTime;
    }
    
    public get requestNextPlanTime()
    {
        return this._requestNextPlanTime;
    }
    
    public get reportCellTime()
    {
        return this._reportCellTime;
    }
    
    public get roomIdByCellReport()
    {
        return this._roomIdByCellReport;
    }
    
    public get cellNumberByCellReport()
    {
        return this._cellNumberByCellReport;
    }
    
    public get countTypeByCellReport()
    {
        return this._countTypeByCellReport;
    }
    
    public get invalidTextByCellReport()
    {
        return this._invalidTextByCellReport;
    }
    
    public get reportSortKey()
    {
        return this._reportSortKey;
    }
    
    public get endOfWorkReportTime()
    {
        return this._endOfWorkReportTime;
    }
    
    public get endOfWorkReportRoomId()
    {
        return this._endOfWorkReportRoomId;
    }
    
    public get reportRetryTime()
    {
        return this._reportRetryTime;
    }
    
    public get isReportRetryTime()
    {
        return this._isReportRetryTime;
    }
    
    public get requestWriteTimes()
    {
        return this._requestWriteTimes;
    }
    
    public get reportTime()
    {
        return this._reportTime;
    }
    
    public get testRoomId()
    {
        return this._testRoomId;
    }
    
    public get workerRoomId()
    {
        return this._workerRoomId;
    }
    
    public get managementRoomId()
    {
        return this._managementRoomId;
    }
    
    public get template()
    {
        return this._template;
    }
    
    public get isTest()
    {
        return this._isTest;
    }
    
    public get startRow()
    {
        return this._startRow;
    }
    
    public get rangeRow()
    {
        return this._rangeRow;
    }
    
    public get apiToken()
    {
        return this._apiToken;
    }
    
    public get prefix()
    {
        return this._prefix;
    }
    
    public get driveDirectoryId()
    {
        return this._driveDirectoryId;
    }
    
    public get calendarSheetKey()
    {
        return this._calendarSheetKey;
    }
}

