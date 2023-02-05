/**
 * docRotater for Google Apps Script
 * ドキュメントを自動バックアップするためのスクリプト
 */

// バックアップ先のフォルダ名。ドキュメントの存在するフォルダと同じ場所にフォルダが作成されます。${yyyy} ${mm} ${dd} ${filename} が指定可能
const BACKUP_FOLDER = '過去の議事録.${yyyy}年';

// バックアップ後のファイル名。ドキュメントの存在するフォルダと同じ場所にフォルダが作成されます。${yyyy} ${mm} ${dd} ${filename} が指定可能
// ファイル名自体に日付文字列が含まれる場合、この指定は無視されます(ファイル名に含まれる日付文字列が更新されます)
const BACKUP_FILENAME = '${filename}.${yyyy}.${mm}';

// トリガー起動時間、配列で指定 (毎日 nn 時ごろに起動) ※起動失敗する場合があるので複数回設定する
const TRIGGER_HOURS = [1, 3, 5];

// 設定保存用のキー名
const PKEY_SETTING = 'DAYS_OF_WEEK';
const PKEY_LASTBACKUP = 'LAST_BACKUP';

// エラーリトライ回数
const MAX_ERROR_RETRY = 4;



/** @classdesc 処理対象の曜日を記録する為のクラス **/
class BitDayOfWeek
{
	static getInstance()
	{
		if(BitDayOfWeek._instance === undefined) {
			BitDayOfWeek._instance = new BitDayOfWeek();
		}
		return BitDayOfWeek._instance;
	}

	constructor()
	{
		let store = getPropertiesStore();
		let value = store.getProperty(PKEY_SETTING);
		this.bitDayOfWeek = parseInt(value == undefined ? 0 : parseInt(value));
	}

	/** 該当日かどうかのテスト 
	 * @return {boolean} true=該当日
	 * @param {Date|number} date_or_dayOfWeek DateかDate.getDay() [曜日]を指定
	 */
	test(date_or_dayOfWeek)
	{
		if(date_or_dayOfWeek instanceof Date) date_or_dayOfWeek = date_or_dayOfWeek.getDay();
		return (this.bitDayOfWeek & (1 << date_or_dayOfWeek)) !== 0 ? true : false;
	}

	/** 該当日(のビット)を設定する
	 * @param {Date|number} date_or_dayOfWeek DateかDate.getDay() [曜日]を指定
	 */
	setBit(date_or_dayOfWeek)
	{
		if(date_or_dayOfWeek instanceof Date) date_or_dayOfWeek = date_or_dayOfWeek.getDay();
		this.bitDayOfWeek |= (1 << date_or_dayOfWeek);
		let store = getPropertiesStore();
		store.setProperty(PKEY_SETTING, '' + this.bitDayOfWeek);
	}

	/** ビット列を設定する (日～土までを一括で指定する)
	 * @param {number} bitValue 7ビット値、LSB→MSBで 日、月、火、水、木、金、土
	 */
	loadValue(bitValue)
	{
		this.bitDayOfWeek = bitValue == null ? 0 : bitValue;
		let store = getPropertiesStore();
		if(this.bitDayOfWeek != 0) store.setProperty(PKEY_SETTING, '' + this.bitDayOfWeek);
		else store.deleteProperty(PKEY_SETTING);
	}

	/** ビット列を取得する
	 * @return {number} 7ビット値、LSB→MSBで 日、月、火、水、木、金、土
	 */
	bitValue()
	{
		return this.bitDayOfWeek;
	}
}

/** プロパティストアの取得
 * @return {Properties} GoogleApps.Properties -- https://developers.google.com/apps-script/reference/properties/properties
 */
function getPropertiesStore()
{ // make as a dedicated function, in case of changing properties stores
	if(getPropertiesStore.cache === undefined) {
		getPropertiesStore.cache = PropertiesService.getScriptProperties(); // everyone should use the same prop storage.
	}
	return getPropertiesStore.cache;
}

/** ActiveDocumentに対応するFile Objectを取得する
 * @return {File} GoogleApps.File -- https://developers.google.com/apps-script/reference/drive/file
 */
function getActiveFile()
{ // make as a dedicated function, in case for supporting other doc types ... like Spreadsheet or Slide
	if(getActiveFile.cahce === undefined) {
		let id = DocumentApp.getActiveDocument().getId();
		getActiveFile.cahce = DriveApp.getFileById(id);
	}
	return getActiveFile.cahce;
}

/** ActiveDocument の Ui を取得
 * @return {Ui GoogleApps.Ui -- https://developers.google.com/apps-script/reference/base/ui
 */
function getUI()
{ // make as a dedicated function, in case for supporting other doc types ... like Spreadsheet or Slide
	if(getUI.cache === undefined) {
		getUI.cache = DocumentApp.getUi();
	}
	return getUI.cache;
}


/** ActiveDocument に表示するメニューの設定 (あるいは再設定)
 */
function createMenu()
{
	const menus = [
		{func: onMenu_Sunday, label: 'Sunday(日)', },
		{func: onMenu_Monday, label: 'Monday(月)'},
		{func: onMenu_Tuesday, label: 'Tuesday(火)'},
		{func: onMenu_Wednesday, label: 'Wednesday(水)', },
		{func: onMenu_Thurseday, label: 'Thurseday(木)', },
		{func: onMenu_Friday, label: 'Firday(金)', },
		{func: onMenu_Sataurday, label: 'Sataurday(土)', },
		null,
		{func: onMenu_DeleteAll, label: 'Delete all (設定削除)', },
		{func: onMenu_Help, label: 'Help (使い方)', },
	];

	let bitDays = BitDayOfWeek.getInstance();

	let menu = getUI().createMenu('[docRotater/議事録保存]');
	let i = 0;
	for(let v of menus) {
		if(v == null) {
			menu.addSeparator();
		} else {
			let check = (i < 7 && bitDays.test(i)) ? '✔ ' : '  ';
			menu.addItem(check + v.label, v.func.name);
		}
		i++;
	}
	menu.addSeparator();
	menu.addToUi();
}

/** ハンドラ -- ファイル表示時
 */
function onOpen()
{
	createMenu();
}

/** ハンドラ -- メニュー選択時
 */
function onMenu_Sunday() {setupEverydayTrigger(0);}
function onMenu_Monday() {setupEverydayTrigger(1);}
function onMenu_Tuesday() {setupEverydayTrigger(2);}
function onMenu_Wednesday() {setupEverydayTrigger(3);}
function onMenu_Thurseday() {setupEverydayTrigger(4);}
function onMenu_Friday() {setupEverydayTrigger(5);}
function onMenu_Sataurday() {setupEverydayTrigger(6);}

/** ハンドラ -- メニュー選択時 > 設定削除
 */
function onMenu_DeleteAll()
{
	let bitDays = BitDayOfWeek.getInstance();
	bitDays.loadValue(0);
	removeTrigger();
	createMenu();
	showResult(null);
}

/** ハンドラ -- メニュー選択時 > ヘルプ
 */
function onMenu_Help()
{
	dialog([
		'Select a day of week. The document would be backup at midnight of the selected day of week.',
		'',
		'ドキュメント保存する曜日を選択してください。指定した曜日の夜中に自動保存が実行されます。'
	].join('\n'));
	return;
}

/** ハンドラヘルパ -- 自動起動の設定
 */
function setupEverydayTrigger(dayOfWeek)
{
	let bitDays = BitDayOfWeek.getInstance();
	setTriggerEveryNight();
	bitDays.setBit(dayOfWeek);
	createMenu();
	showResult(bitDays.bitValue(), dayOfWeek);
}

/** ハンドラヘルパ -- 自動起動の設定結果をダイアログ通知
 */
function showResult(bitDayOfWeek, dayOfWeek = undefined)
{
	if(bitDayOfWeek == null || bitDayOfWeek === 0) {
		dialog('Document audo-backup setting has been deleted.\n'
			+ 'ドキュメントの自動保存設定は削除されました');
	} else {
		let nameDaysEN = ['Sun', 'Mon', 'Tue', 'Wed', 'Thr', 'Fri', 'Sat'];
		let nameDaysJP = ['日', '月', '火', '水', '木', '金', '土'];
		for(let i = 0; i < 7; i++) {
			if((bitDayOfWeek & (1 << i)) === 0) {
				nameDaysEN[i] = '***';
				nameDaysJP[i] = '✖';
			} else if(i === dayOfWeek) {
				nameDaysEN[i] = '+' + nameDaysEN[i] + '+';
				nameDaysJP[i] = '+' + nameDaysJP[i] + '+';
			}
		}
		dialog([
			'Document audo-backup has been set for following days of week.',
			'---> [' + nameDaysEN.join(', ') + ']',
			'',
			'ドキュメントの自動保存が以下の曜日で設定されました(されています)',
			'---> [' + nameDaysJP.join(', ') + ']',
		].join('\n'));
	}
	return;
}

/** ハンドラ -- 時限トリガーによる起動
 */
function onTrigger_AtEveryNight()
{
	let dateNow = new Date();
	while(true) {
		let lock = getDocumentLock();
		if(lock == undefined) break;
		try {
			let success = lock.tryLock(1000 * 10);
			if(!success) break;
			rotateDocument(dateNow);
		} finally {
			if(lock != undefined) lock.releaseLock();
			lock = undefined;
		}
		return;
	}
	console.log('Cound not get lock. give up!');
	return;
}

/** 時限トリガーの設定
 * @param {Function?} funcTrigger 自動起動設定される関数 (名前付き関数であるコト)
 */
function setTriggerEveryNight(funcTrigger = onTrigger_AtEveryNight)
{
	removeTrigger(funcTrigger); // remove trigger in advence, to suppress multiple triggers

	for(let hour of TRIGGER_HOURS) {
		let trigger = undefined;
		for(let i = MAX_ERROR_RETRY; i >= 0; i--) {
			try {
				trigger = ScriptApp.newTrigger(funcTrigger.name)
					.timeBased()
					.atHour(hour)
					.everyDays(1)
					.create();
			} catch(e) {
				console.log(e);
			}
			if(trigger !== undefined) break;
			Utilities.sleep(1000 * (i + 1));
		}
		if((trigger == null)
			|| (trigger.getEventType() != ScriptApp.EventType.CLOCK)
			|| (trigger.getHandlerFunction() !== funcTrigger.name)) {
			throw new Error('Failed to setup a trigger / 自動実行の設定に失敗しました');
		}
	}
	return;
}

/** 時限トリガーの解除
 * @param {Function?} funcTrigger 自動起動設定される関数 (名前付き関数であるコト)
 */
function removeTrigger(funcTrigger = onTrigger_AtEveryNight)
{
	let result = false;
	for(let i = MAX_ERROR_RETRY; i >= 0; i--) {
		try {
			let triggers = ScriptApp.getProjectTriggers();
			for(let r = 0; r < triggers.length; r++) {
				var func = triggers[r].getHandlerFunction(); // ここで得られるのは String 型
				if(func === funcTrigger.name) {
					ScriptApp.deleteTrigger(triggers[r]);
				}
			}
			result = true;
			break;
		} catch(e) {
			console.log(e);
		}
		Utilities.sleep(1000 * (i + 1));
	}
	if(!result) throw new Error('Failed to delete triggers / 自動実行の解除に失敗しました');
	return;
}

/** ダイアログ表示
 * @param {string} title タイトル文字列 ※msg省略時はtitleがメッセージ本文扱いになる
 * @param {string?} msg メッセージ分
 */
function dialog(title, msg = undefined)
{
	var ui = getUI();
	if(msg != null) result = ui.alert(title, msg, ui.ButtonSet.OK);
	else result = ui.alert(title, ui.ButtonSet.OK);
	return;
}

/** 文字列中のマクロ展開
 * @return {string} 結果文字列
 * @param {string} strSource マクロ展開対象となる文字列
 * @param {Date} date 日付→マクロ置換される ${yyyy}${mm}${m} ${dd} ${d}
 * @param {string} filename ファイル名→マクロ置換される ${filename}
*/
function replaceStringMacros(strSource, date, filename)
{
	let year = date.getFullYear();
	let month = date.getMonth() + 1;
	let dayOfMonth = date.getDate();
	let result = strSource;
	if(date) {
		result = result
			.replace('${yyyy}', (year + '').padStart(4, '0'))
			.replace('${y}', year + '')
			.replace('${mm}', (month + '').padStart(2, '0'))
			.replace('${m}', month + '')
			.replace('${dd}', (dayOfMonth + '').padStart(2, '0'))
			.replace('${d}', dayOfMonth + '');
	}
	if(filename) {
		result = result.replace('${filename}', filename);
	}
	return result;
}

/** ファイル名に含まれる日付文字列を更新
 * @return {string} filename変換結果
 * @param {string} filename 
 * @param {Date} date 
*/
function replaceDateInFilename(filename, date)
{
	const regexDate = /([0-9][0-9][0-9][0-9])([年\/\.\-])([0-9]?[0-9])([月\/\.\-])([0-9]?[0-9])([日]?)/;
	let match = filename.match(regexDate);
	if(match) {
		let [all, year, sep1, month, sep2, dayOfMonth, sep3] = match;
		year = (date.getFullYear() + '').padStart(year.length, '0');
		month = (date.getMonth() + 1 + '').padStart(month.length, '0');
		dayOfMonth = (date.getDate() + '').padStart(dayOfMonth.length, '0');
		filename = filename.replace(all, year + sep1 + month + sep2 + dayOfMonth + sep3);
	}
	return filename;
}


/** 排他制御 (ドキュメントロックを使用)
 * @return {Lock} https://developers.google.com/apps-script/reference/lock/lock-service#getDocumentLock()
 */
function getDocumentLock()
{
	let lock = LockService.getDocumentLock();
	let success = false;
	for(let i = 0; i < 4; i++) {
		try {success = lock.tryLock(1000 * (i + 1));}
		catch(e) { /* nothing todo */}
		if(success) break;
	}
	return success ? lock : undefined;
}


/** 指定日の作業が実行済みかどうか検査
 * @return {boolean} true=実行済
 * @param {Date} 検査対象日
 */
function hasUpdateAlready(dateNow)
{
	let store = getPropertiesStore();
	let value = store.getProperty(PKEY_LASTBACKUP);
	let lastUpdate = new Date(parseInt(value));
	if(lastUpdate && lastUpdate.toString() !== 'Invalid Date') {
		if(dateNow < lastUpdate) return true;
	}
	return false;
}

/** 作業実行済みマークを付与
 * @param {Date} 実行日時
 */
function markUpdateAlready(dateNow)
{
	let store = getPropertiesStore();
	store.setProperty(PKEY_LASTBACKUP, '' + dateNow.getTime());
	return;
}

/** ドキュメントのバックアップ処理
 * @param {Date} 実行日時
 */
function rotateDocument(dateNow)
{
	let today = new Date(dateNow.getFullYear(), dateNow.getMonth(), dateNow.getDate());
	let bitDays = BitDayOfWeek.getInstance();
	if(false == bitDays.test(today.getDay())) return;
	if(hasUpdateAlready(today)) return;

	let file = getActiveFile();
	let filename = file.getName();
	let parents = file.getParents();
	let folder = parents.hasNext() ? parents.next() : undefined;

	let destFolder = createFolder(folder, replaceStringMacros(BACKUP_FOLDER, today, null));
	let filenameNew = replaceDateInFilename(filename, dateNow);
	if(filenameNew === filename) {
		filenameNew = replaceStringMacros(BACKUP_FILENAME, today, filename);
		copyFile(file, destFolder, filenameNew);
	} else {
		copyFile(file, destFolder, filename);
		file.setName(filenameNew);
	}

	markUpdateAlready(dateNow);
	return;
}

/** フォルダ作成 ※同名フォルダがある場合は作成せず、既存フォルダを再利用する
 * @param {Folder} folder 親フォルダ
 * @param {string} name 作成するフォルダの名前
 */
function createFolder(folder, name)
{
	let founds = folder.getFoldersByName(name);
	if(founds.hasNext()) return founds.next();

	let folderNew = undefined;
	for(let i = MAX_ERROR_RETRY; i >= 0; i--) {
		try {folderNew = folder.createFolder(name);}
		catch(e) {console.log(e);}
		if(folderNew !== undefined) break;
		Utilities.sleep(1000 * (i + 1));
	}
	if(folderNew == undefined) {
		throw new Error('Could not create a folder -- ' + name);
	}
	return folderNew;
}

/** ファイルコピー ※同名ファイルが存在する場合は 何もしない
 * @return{File} コピー先のファイル (あるいは既存の同名ファイル)
 * @param {File} fileSrc コピー元ファイル
 * @param {Folder} folderDest コピー先フォルダ
 * @param {string} name 作成するファイルの名前
 */
function copyFile(fileSrc, folderDest, filename)
{
	let founds = folderDest.getFilesByName(filename);
	if(founds.hasNext()) return founds.next();

	let fileNew = undefined;
	for(let i = MAX_ERROR_RETRY; i >= 0; i--) {
		try {fileNew = fileSrc.makeCopy(filename, folderDest);}
		catch(e) {console.log(e);}
		if(fileNew !== undefined) break;
		Utilities.sleep(1000 * (i + 1));
	}
	if(fileNew == undefined) {
		throw new Error('Could not copy a file to ' + filename);
	}
	return fileNew;
}

// END of FILE //
