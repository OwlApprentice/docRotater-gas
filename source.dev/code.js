// バックアップ先のフォルダ名。ドキュメントの存在するフォルダと同じ場所にフォルダが作成されます。${yyyy} ${mm} ${dd} ${filename} が指定可能
const BACKUP_FOLDER = '過去の議事録.${yyyy}年';

// バックアップ先のフォルダ名。ドキュメントの存在するフォルダと同じ場所にフォルダが作成されます。${yyyy} ${mm} ${dd} ${filename} が指定可能
const BACKUP_FILENAME = '${filename}.${yyyy}.${mm}';

// トリガー起動時間 (毎日 nn 時ごろに起動)
const TRIGGER_AT_HOUR = 3;

// 設定保存用のキー名
const PKEY_SETTING = 'DAYS_OF_WEEK';
const PKEY_LASTBACKUP = 'LAST_BACKUP';

// エラーリトライ回数
const MAX_ERROR_RETRY = 4;



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
		this.store = PropertiesService.getScriptProperties(); // everyone shoud use the same prop storage.
		// this.store = PropertiesService.getUserProperties();  so we do not use UserProperties Store.
		let value = this.store.getProperty(PKEY_SETTING);
		this.bitDayOfWeek = parseInt(value == undefined ? 0 : parseInt(value));
	}

	test(date_or_dayOfWeek)
	{
		if(date_or_dayOfWeek instanceof Date) date_or_dayOfWeek = date_or_dayOfWeek.getDay();
		return (this.bitDayOfWeek & (1 << date_or_dayOfWeek)) !== 0 ? true : false;
	}

	setBit(date_or_dayOfWeek)
	{
		if(date_or_dayOfWeek instanceof Date) date_or_dayOfWeek = date_or_dayOfWeek.getDay();
		this.bitDayOfWeek |= (1 << date_or_dayOfWeek);
		this.store.setProperty(PKEY_SETTING, '' + this.bitDayOfWeek);
	}

	loadValue(bitValue)
	{
		this.bitDayOfWeek = bitValue == null ? 0 : bitValue;
		if(this.bitDayOfWeek != 0) this.store.setProperty(PKEY_SETTING, '' + this.bitDayOfWeek);
		else this.store.deleteProperty(PKEY_SETTING);
	}

	bitValue()
	{
		return this.bitDayOfWeek;
	}
}

function getActiveFile()
{ // make as a dedicated function, in case for supporting other doc types ... like Spreadsheet or Slide
	if(getActiveFile.cahce === undefined) {
		let id = DocumentApp.getActiveDocument().getId();
		getActiveFile.cahce = DriveApp.getFileById(id);
	}
	return getActiveFile.cahce;
}

function getUI()
{ // make as a dedicated function, in case for supporting other doc types ... like Spreadsheet or Slide
	if(getUI.cache === undefined) {
		getUI.cache = DocumentApp.getUi();
	}
	return getUI.cache;
}

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

function onOpen()
{
	createMenu();
}

function onMenu_Sunday() {setupEverydayTrigger(0);}
function onMenu_Monday() {setupEverydayTrigger(1);}
function onMenu_Tuesday() {setupEverydayTrigger(2);}
function onMenu_Wednesday() {setupEverydayTrigger(3);}
function onMenu_Thurseday() {setupEverydayTrigger(4);}
function onMenu_Friday() {setupEverydayTrigger(5);}
function onMenu_Sataurday() {setupEverydayTrigger(6);}

function onMenu_DeleteAll()
{
	let bitDays = BitDayOfWeek.getInstance();
	bitDays.loadValue(0);
	removeTrigger();
	createMenu();
	showResult(null);
}

function setupEverydayTrigger(dayOfWeek)
{
	let bitDays = BitDayOfWeek.getInstance();
	bitDays.setBit(dayOfWeek);
	setTriggerEveryNight();
	createMenu();
	showResult(bitDays.bitValue(), dayOfWeek);
}

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

function onMenu_Help()
{
	dialog([
		'Select a day of week. The document would be backup at approx 3:00am at selected day of week.',
		'',
		'ドキュメント保存する曜日を選択してください。指定した曜日の夜中03:00頃に自動保存が実行されます。'
	].join('\n'));
	return;
}

function onTigger_AtEveryNight()
{
	let bitDays = BitDayOfWeek.getInstance();
	let date = Date.new();
	if(bitDays.test(date)) {
		let lock = getDocumentLock();
		if(lock === undefined) {
			console.log('Cound not get lock. give up!');
		}
		try {
			rotateDocument(date);
		} finally {
			lock.releaseLock();
		}
	}
	return;
}

function setTriggerEveryNight(funcTrigger = onTigger_AtEveryNight)
{
	removeTrigger(funcTrigger); // remove trigger in advence, to suppress multiple triggers

	let trigger = undefined;
	for(let i = MAX_ERROR_RETRY; i >= 0; i--) {
		try {
			trigger = ScriptApp.newTrigger(funcTrigger.name)
				.timeBased()
				.atHour(TRIGGER_AT_HOUR)
				.everyDays(1)
				.create();
		} catch(e) {
			console.log(e);
		}
		if(trigger !== undefined) break;
		Utilities.Speep(1000 * (i + 1));
	}

	if((trigger == null)
		|| (trigger.getEventType() != ScriptApp.EventType.CLOCK)
		|| (trigger.getHandlerFunction() !== funcTrigger.name)) {
		showError('Failed to set a trigger / 自動実行の設定に失敗しました');
	}
	return;
}

function removeTrigger(funcTrigger = onTigger_AtEveryNight)
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
		} catch(e) {
			console.log(e);
		}
		Utilities.Speep(1000 * (i + 1));
	}
	if (!result) throw new Error('Failed to delete triggers / 自動実行の解除に失敗しました');
	return;
}

function showError(msg)
{
	dialog(msg);
	return;
}

function dialog(title, msg = undefined)
{
	var ui = getUI();
	if(msg != null) result = ui.alert(title, msg, ui.ButtonSet.OK);
	else result = ui.alert(title, ui.ButtonSet.OK);
	return;
}

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

function rotateDocument(dateNow)
{
	let today = new Date(dateNow.getFullYear(), dateNow.getMonth(), dateNow.getDate());

	let store = PropertiesService.getScriptProperties(); // everyone shoud use the same prop storage.
	let value = store.getProperty(PKEY_LASTBACKUP);
	let lastUpdate = new Date(parseInt(value));
	if(lastUpdate && lastUpdate.toString() !== 'Invalid Date') {
		if(lastUpdate < today) return;
	}

	let file = getActiveFile();
	let filename = file.getName();
	let parents = file.getParents();
	let folder = parents.hasNext() ? parents.next() : undefined;

	let destFolder = createFolder(folder, replaceStringMacros(BACKUP_FOLDER, today, null));
	copyFile(replaceStringMacros(BACKUP_FILENAME, today, filename), destFolder);

	store.setProperty(PKEY_LASTBACKUP, '' + dateNow.getTime());
	return;
}

function createFolder(folder, name)
{
	let founds = folder.getFoldersByName(name);
	if(founds.hasNext()) return founds.next();

	let folderNew = undefined;
	for(let i = MAX_ERROR_RETRY; i >= 0; i--) {
		try {folderNew = folder.createFolder(name);}
		catch(e) {console.log(e);}
		if(folderNew !== undefined) break;
		Utilities.Speep(1000 * (i + 1));
	}
	return folderNew; 
}

function copyFile( fileSrc, folderDest, filename )
{
	let fileNew = undefined;
	for(let i = MAX_ERROR_RETRY; i >= 0; i--) {
		try {fileNew = fileSrc.makeCopy(filename, destFolder);}
		catch(e) {console.log(e);}
		if(fileNew !== undefined) break;
		Utilities.Speep(1000 * (i + 1));
	}
	return fileNew;
}

// END of FILE //
