/**
 * GoogleDriveを操作するクラス
 */
export default class GoogleDrive
{
    /**
     * @type {{}}
     */
    protected fileCache = {};
    
    /**
     * ファイルのIDを取得
     *
     * @param {string} configDataDirectoryId
     * @param {string} targetFileName
     * @returns {string | null}
     */
    public getFileId(configDataDirectoryId: string, targetFileName: string)
    {
        // 既にキャッシュ済みの場合は作らない（一度読み込んだものはキャッシュしておく）
        if (!this.fileCache[configDataDirectoryId] || !this.fileCache[configDataDirectoryId][targetFileName]) {
            this.createLocalCache(configDataDirectoryId, targetFileName);
        }
        
        if (this.fileCache[configDataDirectoryId][targetFileName] === "null") {
            return "null";
        }
        
        return this.fileCache[configDataDirectoryId][targetFileName].getId();
    }
    
    /**
     * ファイルのURLを取得
     *
     * @param {string} configDataDirectoryId
     * @param {string} targetFileName
     * @returns {string | null}
     */
    public getSheetURL(configDataDirectoryId: string, targetFileName: string)
    {
        // 既にキャッシュ済みの場合は作らない（一度読み込んだものはキャッシュしておく）
        if (!this.fileCache[configDataDirectoryId] || !this.fileCache[configDataDirectoryId][targetFileName]) {
            this.createLocalCache(configDataDirectoryId, targetFileName);
        }
        
        if (this.fileCache[configDataDirectoryId][targetFileName] === "null") {
            return "null";
        }
        
        return this.fileCache[configDataDirectoryId][targetFileName].getUrl();
    }
    
    /**
     * ドライブ内のファイル情報を取得（キャツシュ化）
     *
     * @param {string} configDataDirectoryId
     * @param {string} targetFileName
     */
    private createLocalCache(configDataDirectoryId: string, targetFileName: string)
    {
        // 二次元の連想配列を実現するための処理
        // 参考:http://nakawake.net/?p=831
        if (!this.fileCache[configDataDirectoryId]) {
            this.fileCache[configDataDirectoryId] = {};
        }
        
        const files = DriveApp.getFolderById(configDataDirectoryId).getFiles();
        
        while (files.hasNext()) {
            const file = files.next();
            const fileName = file.getName();
            
            if (targetFileName === fileName) {
                this.fileCache[configDataDirectoryId][targetFileName] = file;
                return;
            }
        }
        
        this.fileCache[configDataDirectoryId][targetFileName] = "null";
    }
}
