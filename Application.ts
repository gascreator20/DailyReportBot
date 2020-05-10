import Config from "./Config";
import DailyReportSheetReader from "./DailyReportSheetReader";
import GoogleDrive from "./GoogleDrive";
import KeyValueSheet from "./KeyValueSheet";
import SendMessage from "./SendMessage";
import SpreadSheetCreator from "./SpreadSheetCreator";
import Worker from "./Worker";

/**
 * カスタムメニュー
 */
function onOpen()
{
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("日報報告機能")
        .addItem("次回分の日報を用意", "createNextDayTemplate")
        .addItem("作業予定記入チェック（当日分を確認）", "requestTodayPlanError")
        .addItem("作業報告記入依頼", "requestReportWrite")
        .addItem("作業報告", "reportCheck")
        .addItem("作業報告違反のチェック", "reportErrorCheckOnly")
        .addItem("次回の作業予定記入依頼", "requestNextDayPlan")
        .addItem("作業予定記入チェック（次回分を確認）", "requestNextDayPlanError")
        .addItem("終業報告", "endOfWorkReport")
        .addItem("特定のセルの内容を一斉報告", "targetCellReport")
        .addItem("本日の報告トリガーをセット", "setReportTrigger")
        .addItem("トリガーを初期化", "initializationTrigger")
        .addItem("実行トリガーをすべて削除", "deleteTrigger")
        .addItem("必須以外の実行トリガーを削除", "deleteTriggerWithoutRequired")
        .addItem("スプレッドシートIDの読み込み", "readSpreadsheetId")
        .addToUi();
}

/**
 * 特定セルの内容を送信する
 */
function targetCellReport()
{
    if (!isWorkDayToday_()) {
        return;
    }
    const config = new Config();
    const keyValueSheetReader = new KeyValueSheet();
    const worker = new Worker();
    const members = worker.getWorkMember();
    const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
    const calendar = keyValueSheetReader.find("カレンダー", config.calendarSheetKey, nowTime, false);
    const chatWork = new SendMessage();
    
    // カウントを進める
    for (const member of members) {
        const sheetByMember = SpreadsheetApp.openById(member["SpreadsheetID"]).getSheetByName(calendar[config.calendarSheetKey]);
        chatWork.changeWriteCount(member, sheetByMember);
    }
    
    // シートに更新を入れたので新しいデータを再読み込みするためにキャッシュをクリアする
    keyValueSheetReader.flushLocalCache();
    
    // 送信
    chatWork.sendCellContent(worker.getWorkMember(), calendar, "指定セルの内容をまとめて送信のエラーテンプレート");
}

/**
 * 終業報告
 */
function endOfWorkReport()
{
    if (!isWorkDayToday_()) {
        return;
    }
    
    const worker = new Worker();
    const members = worker.getWorkMember();
    
    // 報告
    const chatWork = new SendMessage();
    chatWork.sendEndOfWorkReport(members);
}

/**
 * 作業報告記入依頼
 */
function requestReportWrite()
{
    if (!isWorkDayToday_()) {
        return;
    }
    
    // 作業報告が必要なメンバーリスト
    const reportWriteMembers = [];
    
    // 営業日を取得
    const config = new Config();
    const keyValueSheetReader = new KeyValueSheet();
    const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
    const calendar = keyValueSheetReader.find("カレンダー", config.calendarSheetKey, nowTime);
    
    // 予定が正確に入力されているかどうかの検証
    const worker = new Worker();
    const members = worker.getWorkMember();
    
    for (const member of members) {
        const spreadSheetReader = new DailyReportSheetReader();
        const resultBeforeTime = spreadSheetReader.findWorkResultByBeforeTime(member, calendar);
        const planBeforeTime = resultBeforeTime[3];
        
        // 報告が未記入の場合はエラー
        if (planBeforeTime !== "") {
            reportWriteMembers.push(member);
        }
    }
    
    // 報告
    const chatWork = new SendMessage();
    if (reportWriteMembers.length >= 1) {
        chatWork.sendRequestDailyTimeReport(reportWriteMembers);
    }
}

/**
 * 作業報告エラーのみを行う
 */
function reportErrorCheckOnly()
{
    report_(false);
}

/**
 * 作業報告を行う（エラー検知および成功者の報告）
 */
function reportCheck()
{
    report_(true);
}

/**
 * 作業予定エラー検知（翌日）
 */
function requestNextDayPlanError()
{
    planError_(true);
}

/**
 * 作業予定エラー検知（当日）
 */
function requestTodayPlanError()
{
    planError_(false);
}

/**
 * 翌日の作業予定記入を要求
 */
function requestNextDayPlan()
{
    if (!isWorkDayToday_()) {
        return;
    }
    
    const worker = new Worker();
    const members = worker.getWorkMember(true);
    
    // 報告
    const chatWork = new SendMessage();
    chatWork.sendRequestNextDayPlan(members);
}

/**
 * 本日の作業報告用のトリガーをセットする
 */
function setReportTrigger()
{
    const config = new Config();
    const timesString = config.requestWriteTimes;
    const reportTime = config.reportTime;
    const timesArray = timesString.split(',');
    const nowTime = new Date();
    
    for (const time of timesArray) {
        const [hour, minute] = time.split(':');
        const triggerDate = new Date();
        triggerDate.setHours(Number(hour));
        triggerDate.setMinutes(Number(minute));
        
        // トリガーセットする日時が現在時刻よりも前の場合は正しく時間がセットされないため、トリガーはセットしないものとする
        if (triggerDate < nowTime) {
            continue;
        }
        
        // 作業報告記入依頼トリガー
        ScriptApp.newTrigger('requestReportWrite').timeBased().at(triggerDate).create();
        
        // 作業報告トリガー
        triggerDate.setMinutes(Number(minute) + reportTime);
        ScriptApp.newTrigger('reportCheck').timeBased().at(triggerDate).create();
    }
}

/**
 * トリガーを初期化する
 */
function initializationTrigger()
{
    const config = new Config();
    
    // まず最初にすべてのトリガーを削除
    deleteTrigger();
    
    // 毎日トリガーの初期化処理を行う（自分自身を再セットする）
    ScriptApp.newTrigger('initializationTrigger').timeBased().atHour(0).everyDays(1).create();
    
    // 本日の報告トリガーをセット
    setReportTrigger();
    
    // 終業報告
    const triggerTimeByEndOfWorkReportTime = getTriggerSetTime_(config.endOfWorkReportTime);
    if (triggerTimeByEndOfWorkReportTime) {
        ScriptApp.newTrigger('endOfWorkReport').timeBased().at(triggerTimeByEndOfWorkReportTime).create();
    }
    
    // 日報自動生成
    const triggerTimeByTemplateCreateTime = getTriggerSetTime_(config.templateCreateTime);
    if (triggerTimeByTemplateCreateTime) {
        ScriptApp.newTrigger('createNextDayTemplate').timeBased().at(triggerTimeByTemplateCreateTime).create();
    }
    
    // 当日分の作業予定の検証
    const triggerTimeByCheckTodayPlanTime = getTriggerSetTime_(config.checkTodayPlanTime);
    if (triggerTimeByCheckTodayPlanTime) {
        ScriptApp.newTrigger('requestTodayPlanError').timeBased().at(triggerTimeByCheckTodayPlanTime).create();
    }
    
    // 次回分の作業予定の検証
    const triggerTimeByCheckNextPlanTime = getTriggerSetTime_(config.checkNextPlanTime);
    if (triggerTimeByCheckNextPlanTime) {
        ScriptApp.newTrigger('requestNextDayPlanError').timeBased().at(triggerTimeByCheckNextPlanTime).create();
    }
    
    // 次回分の作業予定の記入依頼
    const triggerTimeByRequestNextPlanTime = getTriggerSetTime_(config.requestNextPlanTime);
    if (triggerTimeByRequestNextPlanTime) {
        ScriptApp.newTrigger('requestNextDayPlan').timeBased().at(triggerTimeByRequestNextPlanTime).create();
    }
    
    // 指定セルの報告
    const triggerTimeByReportCellTime = getTriggerSetTime_(config.reportCellTime);
    if (triggerTimeByReportCellTime) {
        ScriptApp.newTrigger('targetCellReport').timeBased().at(triggerTimeByReportCellTime).create();
    }
}

/**
 * セットされているトリガーを削除する（消す必要がないものは維持）
 */
function deleteTriggerWithoutRequired()
{
    // 消す対象としないトリガー一覧
    const ignore = [
        "createNextDayTemplate",
        "requestTodayPlanError",
        "requestNextDayPlan",
        "requestNextDayPlanError",
        "endOfWorkReport",
        "targetCellReport",
        "setReportTrigger"
    ];
    
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        if (!ignore.includes(trigger.getHandlerFunction())) {
            ScriptApp.deleteTrigger(trigger);
        }
    }
}

/**
 * セットされているトリガーを削除する
 */
function deleteTrigger()
{
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        ScriptApp.deleteTrigger(trigger);
    }
}

/**
 * テンプレートシートを用意する
 */
function createNextDayTemplate()
{
    if (!isWorkDayToday_()) {
        return;
    }
    
    // 次の営業日を取得
    const config = new Config();
    const keyValueSheetReader = new KeyValueSheet();
    const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
    const calendar = keyValueSheetReader.find("カレンダー", config.calendarSheetKey, nowTime, true);
    
    // 指定ディレクトリ内のすべてにテンプレートファイルをコピー（ファイル名は年月日）
    const spreadSheetCreator = new SpreadSheetCreator();
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTemplate = spreadSheet.getSheetByName("テンプレート");
    
    spreadSheetCreator.copy(config.driveDirectoryId, sheetTemplate, calendar[config.calendarSheetKey]);
}

/**
 * スプレッドシートID、URLを読み込む
 */
function readSpreadsheetId()
{
    const config = new Config();
    const keyValueSheetReader = new KeyValueSheet();
    const memberNames = keyValueSheetReader.getAllByColumnName("メンバーリスト", "名前");
    const spreadsheetIdColumnNumber = keyValueSheetReader.getColumnNumber("メンバーリスト", "SpreadsheetID");
    const sheetUrlColumnNumber = keyValueSheetReader.getColumnNumber("メンバーリスト", "SheetURL");
    
    // ドライブ内のファイルID、URLを取得する
    const googleDrive = new GoogleDrive();
    let loopCount = 2;
    for (const name of memberNames) {
        const sheetId = googleDrive.getFileId(config.driveDirectoryId, config.prefix + name);
        const sheetUrl = googleDrive.getSheetURL(config.driveDirectoryId, config.prefix + name);
        
        keyValueSheetReader.write("メンバーリスト", loopCount, spreadsheetIdColumnNumber, sheetId);
        keyValueSheetReader.write("メンバーリスト", loopCount, sheetUrlColumnNumber, sheetUrl);
        loopCount++;
    }
}

/**
 * トリガーセット用の時刻を取得（nullであればトリガーはセットしない）（非公開扱い）
 */
function getTriggerSetTime_(time: string)
{
    // 無効な値はtriggerセット
    if (time === null || time === "null" || time === "") {
        return null;
    }
    
    const nowTime = new Date();
    const [hour, minute] = time.split(':');
    const triggerDate = new Date();
    triggerDate.setHours(Number(hour));
    triggerDate.setMinutes(Number(minute));
    
    
    // トリガーセットする日時が現在時刻よりも前の場合は正しく時間がセットされないため、トリガーはセットしないものとする
    if (triggerDate < nowTime) {
        return null;
    }
    
    return triggerDate;
}

/**
 * 作業予定エラー検知（非公開扱い）
 */
function planError_(isNextDay: boolean)
{
    if (!isWorkDayToday_()) {
        return;
    }
    
    const errorMember = [];
    
    // 設定値の取得
    const config = new Config();
    const keyValueSheetReader = new KeyValueSheet();
    
    // 次の営業日を取得
    const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
    const calendar = keyValueSheetReader.find("カレンダー", config.calendarSheetKey, nowTime, isNextDay);
    
    // 予定が正確に入力されているかどうかの検証
    const worker = new Worker();
    const members = worker.getWorkMember(isNextDay);
    for (const member of members) {
        const spreadSheetReader = new DailyReportSheetReader();
        const result = spreadSheetReader.findDailyReportByWorkStartTime(member, calendar);
        const plan = result[3];
        
        if (plan === null) {
            console.log("参照できませんでした:" + member["名前"])
            errorMember.push(member["名前"]);
        } else if (plan === "") {
            console.log("予定が空白です:" + member["名前"])
            errorMember.push(member["名前"]);
        }
    }
    
    // 報告
    const chatWork = new SendMessage();
    chatWork.sendRequestNextDayPlanError(errorMember, isNextDay);
}

/**
 * 作業報告（非公開扱い）
 */
function report_(needSuccessReport: boolean = true)
{
    if (!isWorkDayToday_()) {
        return;
    }
    
    const errorMember = [];
    const ngRuleMember = [];
    const successMember = [];
    let error = false;
    
    // 営業日を取得
    const config = new Config();
    const keyValueSheetReader = new KeyValueSheet();
    const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
    const calendar = keyValueSheetReader.find("カレンダー", config.calendarSheetKey, nowTime);
    
    // 予定が正確に入力されているかどうかの検証
    const worker = new Worker();
    const members = worker.getWorkMember();
    
    for (const member of members) {
        // 勤務時間外の場合はnullが返る。nullの場合は何も処理していないので必然的に作業報告者対象から除外される
        const spreadSheetReader = new DailyReportSheetReader();
        const resultBeforeTime = spreadSheetReader.findWorkResultByBeforeTime(member, calendar);
        const resultReportBeforeTime = resultBeforeTime[5];
        const resultNowTime = spreadSheetReader.findWorkResultByNowTime(member, calendar);
        const resultReportNowTime = resultNowTime[5];
        
        if (resultReportBeforeTime === "") {
            // 報告が未記入の場合はエラー
            errorMember.push(member);
            error = true;
        } else if (resultReportNowTime !== "") {
            // 報告の先記入は禁止
            ngRuleMember.push(member);
            error = true;
        } else if (resultReportBeforeTime !== "" && needSuccessReport) {
            // 報告が正しい
            successMember.push(member);
        }
    }
    
    // 報告（対象者がいるもののみ送信する）
    const chatWork = new SendMessage();
    if (ngRuleMember.length >= 1) {
        chatWork.sendDailyTimeReportNgRule(errorMember);
    }
    if (errorMember.length >= 1) {
        chatWork.sendDailyTimeReportError(errorMember);
    }
    if (successMember.length >= 1 && needSuccessReport) {
        chatWork.sendDailyTimeReport(errorMember);
    }
    
    // 再通知トリガー（エラー通知のループを避けるために成功時の報告の時でしか再通知トリガーはセットしない）
    if (error && needSuccessReport && config.isReportRetryTime) {
        const triggerDate = new Date();
        triggerDate.setHours(triggerDate.getHours());
        triggerDate.setMinutes(triggerDate.getMinutes() + Number(config.reportRetryTime));
        
        // 再報告用のトリガーを設置
        ScriptApp.newTrigger('reportErrorCheckOnly').timeBased().at(triggerDate).create();
    }
}

/**
 * 営業日チェック（非公開扱い）
 */
function isWorkDayToday_()
{
    // 設定値の取得
    const config = new Config();
    const keyValueSheetReader = new KeyValueSheet();
    const nowTime = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
    const calendar = keyValueSheetReader.find("カレンダー", config.calendarSheetKey, nowTime);
    
    if (!calendar) {
        console.log("本日は営業日ではありません");
        return false;
    }
    
    return true;
}
