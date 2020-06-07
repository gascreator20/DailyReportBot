/**
 * ChatWorkに関する操作を行うためのクラス
 */
import Config from "./Config";

export default class ChatWork
{
    /**
     * ChatWorkへメッセージを送信する
     *
     * @param {string} message
     * @param {string} roomId
     * @param {string} apiToken
     */
    private static sendMessageToChatWork(message: string, roomId: string, apiToken: string)
    {
        // 設定値の取得
        const parameter = {
            headers: {
                "X-ChatWorkToken": apiToken
            },
            method: "post",
            payload: {
                body: message
            }
        };
        
        const url = "https://api.chatwork.com/v2/rooms/" + roomId + "/messages";
        
        // メッセージ送信
        // @ts-ignore
        UrlFetchApp.fetch(url, parameter);
    }
    
    /**
     * メッセージを送信する
     *
     * @param {string} message
     * @param {string} roomId
     */
    public sendMessage(message: string, roomId: string)
    {
        const config = new Config();
        
        // テスト時は必ずテストルームへ送信する
        if (config.isTest === "TRUE" || config.isTest === "true") {
            ChatWork.sendMessageToChatWork(message, config.testRoomId, config.apiToken);
        } else {
            ChatWork.sendMessageToChatWork(message, roomId, config.apiToken);
        }
    }
}

