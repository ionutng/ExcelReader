using ExcelReader;

var file = FileHelper.GetFile();

var info = await FileHelper.GetInfoFromFile(file);

var db = new DatabaseManager();

db.ShowData(info);