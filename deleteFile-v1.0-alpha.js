fso = new ActiveXObject("Scripting.FileSystemObject");
//test
var today = new Date();  		//獲取當前日期
//參數設定
var timeUnit = "isDay";			//設定時間單位，值為isSecond,isMinute,isHour,isDay,isMonth,isYear
var delTime = 1;			//設定刪除幾個單位時間前的文件
var startFolder = "c:\\storage";		//刪除資料夾路徑
var isBak = false;				//是否備份
var bakFolder = "C:\\storageBak" + "-" + today.toLocaleDateString();	//備份資料夾路徑,按日期命名

//初始化,勿修改
var fileCounter = 0;			//初始化文件計數器
var folderCounter = 0;			//初始化資料夾計數器
var errorCounter = 0;			//初始化錯誤計數器
var result = "";			//初始化結果
var output = new ResultWriter();	//初始化日誌記錄

//執行過程
//獲取刪除設定時間
var delDate = getDelDate();

//若需要備份且備份資料夾不存在，則創建備份資料夾,若存在備份資料夾，則在原備份資料夾名字後面加上當前時間最為新資料夾名稱
if(isBak){				
	if(fso.FolderExists(bakFolder) == false){
		fso.CreateFolder(bakFolder);
	}else{
		bakFolder += today.getHours() + "時" + today.getMinutes() + "分" + today.getSeconds() + "秒";
		fso.CreateFolder(bakFolder);
	}
}

output.line("刪除" + delDate.toLocaleString()  + "之前的文件\r\n");
output.line("執行時間為：" + today.toLocaleString() + "\r\n");
output.line("刪除資料夾路徑為：" + startFolder + "\r\n");

if(isBak){
	output.line("備份資料夾路徑為：" + bakFolder + "\r\n");
}

output.line("刪除文件及資料夾列表詳細信息如下：\r\n");

DeleteOldFiles(startFolder,delDate);			//執行刪除操作
DeleteEmptyFolders(startFolder);			//刪除目標路徑內的空資料夾

//匯總結果
//備份成功消息提示
//WScript.Echo("刪除舊的文件及目錄的操作已經完成！\n,共刪除文件" + fileCounter 
//	+ "個，目錄" + folderCounter + "個\n詳細結果請查看日誌文件.");
output.line("\r\n共處理文件" + fileCounter + "個，目錄" + folderCounter + "個\r\n");
output.line("其中刪除成功" + (fileCounter - errorCounter) + "個，刪除失敗" + errorCounter + "個\r\n");

function DeleteOldFiles(folderName,BeforeDate){
	var folder,selFile,fileCollection;
	try{
		folder = fso.GetFolder(folderName);
	}catch(e1){
		output.line("出現異常：" + e1.description + "\r\n");
		output.line("提示：目標文件夾異常，請檢查目標文件夾路徑\r\n");
		return;
	}
	fileCollection = folder.Files;
	var e = new Enumerator(fileCollection);	
	for(;!e.atEnd();e.moveNext()){
		var selFile = e.item();
		
		if(selFile.DateCreated <= BeforeDate){
			//操作記錄
			fileCounter++;
			output.line(fileCounter + ":" + selFile.Name + " in " + selFile.ParentFolder + "\r\n");
			output.line("創建時間為：" + selFile.DateCreated+ "\r\n");
			output.line("最後一次訪問時間為：" +selFile.DateLastAccessed + "\r\n")
			output.line("最後一次修改時間為：" + selFile.DateLastModified + "\r\n")

			if(isBak == true){
				//按原文件路徑生成新文件路徑
				var flPath = selFile.Path.substring(startFolder.length,selFile.Path.length - selFile.Name.length);
				var newPath = bakFolder + flPath;
				if(!fso.FolderExists(newPath)){
					fso.CreateFolder(newPath);
				}
				fso.CopyFile(selFile.Path,newPath,true);
			}
			//刪除原文件
			try{
				fso.deleteFile(selFile.path,true);
			}catch(e2){
				output.line("出現異常：" + e2.description + "\r\n");
				output.line("提示：可能是該文件正在使用中,該文件刪除失敗，跳過此文件繼續執行\r\n");
				errorCounter++;
			}finally{
				output.line(result);
				continue;
			}
		}
	}
	var enumSubFolder = new Enumerator(folder.SubFolders);
	//迴圈遍歷子資料夾
	for(;!enumSubFolder.atEnd();enumSubFolder.moveNext()){
		DeleteOldFiles(enumSubFolder.item().Path,BeforeDate);
	}
}
function DeleteEmptyFolders(folderName){
	var folder = fso.GetFolder(folderName);
	if(folder.Files.Count == 0 && folder.SubFolders.Count == 0){
		output.line(folder.Name + " in " + folder.ParentFolder + "\r\n");
		output.line("創建時間為：" + folder.DateCreated + "\r\n");
		output.line("最後一次訪問時間為：" + folder.DateLastAccessed + "\r\n");
		output.line("最後一次修改時間為：" + folder.DateLastModified + "\r\n")
		fso.DeleteFolder(folder.Path);
		folderCounter++;
	}else if(folder.SubFolders.Count != 0){
		var enumSubFolder = new Enumerator(folder.SubFolders);
		for(;!enumSubFolder.atEnd();enumSubFolder.moveNext()){
			DeleteEmptyFolders(enumSubFolder.item());
		}
	}
}
//獲取設定日期函數
function getDelDate(){
	
	var OlderThanDate = new Date();
	var time = today.getTime();		//獲取當前時間
	var MinMilli = 1000 * 60;		//一分鐘有多少微秒
	var HrMilli = MinMilli * 60;		//一小時有多少微秒
	var DyMilli = HrMilli * 24;		//一天有多少微秒

	switch(timeUnit){
		case "isSecond" :
			OlderThanDate.setTime(time - (delTime * 1000));break;
		case "isMinute" :
			OlderThanDate.setTime(time - (delTime * MinMilli));break;
		case "isHour"   :
			OlderThanDate.setTime(time - (delTime * HrMilli));break;
		case "isDay"    :
			OlderThanDate.setTime(time - (delTime * DyMilli));break;
		case "isMonth" :
			OlderThanDate.setMonth(today.getMonth() - delTime);break;
		case "isYear"   :
			OlderThanDate.setYear(today.getYear() - delTime);break;
	}
	return OlderThanDate;
}
//日誌記錄器
function ResultWriter(){
	var savepath = WScript.ScriptFullName.substr(0,(WScript.ScriptFullName.length-WScript.ScriptName.length));
	var rflPath = savepath+"deleteFiles-" + today.toLocaleDateString();
	if(fso.FileExists(rflPath + ".log")){
		ResultFile = fso.CreateTextFile(rflPath + today.getHours() + "時" + today.getMinutes() + "分" + today.getSeconds() + "秒.log",false,true);
	}else{
		ResultFile = fso.CreateTextFile(rflPath + ".log",false,true);
	}
	this.file=ResultFile;
	this.line=ResultWriter_Line;
}
function ResultWriter_Line(strings){
	this.file.Write(strings);
}
