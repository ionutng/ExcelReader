using ExcelReader;

var fileHelper = new FileHelper();

var info = await fileHelper.GetInfoFromFile(fileHelper.GetFile());

var db = new DatabaseManager();

db.ShowData(info);