﻿using System.Collections.Immutable;
using System.Text;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Utf8StringInterpolation;
using ZLogger;
using ZLogger.Providers;

//==
var builder = ConsoleApp.CreateBuilder(args);
builder.ConfigureServices((ctx,services) =>
{
    // Register appconfig.json to IOption<MyConfig>
    services.Configure<MyConfig>(ctx.Configuration);

    // Using Cysharp/ZLogger for logging to file
    services.AddLogging(logging =>
    {
        logging.ClearProviders();
        logging.SetMinimumLevel(LogLevel.Trace);
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        var utcTimeZoneInfo = TimeZoneInfo.Utc;
        logging.AddZLoggerConsole(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });
        });
        logging.AddZLoggerRollingFile(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });

            // File name determined by parameters to be rotated
            options.FilePathSelector = (timestamp, sequenceNumber) => $"logs/{timestamp.ToLocalTime():yyyy-MM-dd}_{sequenceNumber:00}.log";
            
            // The period of time for which you want to rotate files at time intervals.
            options.RollingInterval = RollingInterval.Day;
            
            // Limit of size if you want to rotate by file size. (KB)
            options.RollingSizeKB = 1024;        
        });
    });
});

var app = builder.Build();
app.AddCommands<PortAssignApp>();
app.Run();


public class PortAssignApp : ConsoleAppBase
{
    bool isAllPass = true;

    readonly ILogger<PortAssignApp> logger;
    readonly IOptions<MyConfig> config;

    Dictionary<string, MyAssignDevice> dicMyAssignDevice = new Dictionary<string, MyAssignDevice>();
    Dictionary<string, MyDevice> dicMyProperty = new Dictionary<string, MyDevice>();
    Dictionary<string, List<MySw>> dicMySw = new Dictionary<string, List<MySw>>();

    public PortAssignApp(ILogger<PortAssignApp> logger,IOptions<MyConfig> config)
    {
        this.logger = logger;
        this.config = config;
    }

//    [Command("")]
    public void Assign(string definition, string props, string save)
    {
//== start
        logger.ZLogInformation($"==== tool {getMyFileVersion()} ====");
        if (!File.Exists(definition))
        {
            logger.ZLogError($"[NG] エクセルファイルが見つかりません{definition}");
            return;
        }
        if (!File.Exists(props))
        {
            logger.ZLogError($"[NG] エクセルファイルが見つかりません{props}");
            return;
        }

        int definitionDataRow = config.Value.DefinitionDataRow;
        string definitionSheetName = config.Value.DefinitionSheetName;
        string definitionWordKeyToColum = config.Value.DefinitionWordKeyToColum;
        string definitionWordNothing = config.Value.DefinitionWordNothing;
        int propertyDataRow = config.Value.PropertyDataRow;
        string propertySheetName = config.Value.PropertySheetName;
        string propertyWordKeyToColum = config.Value.PropertyWordKeyToColum;
        string saveWordFromColum = config.Value.SaveWordFromColum;
        string saveWordToColum = config.Value.SaveWordToColum;

        readDefinitionExcel(definition, definitionSheetName, definitionDataRow, definitionWordKeyToColum, dicMyAssignDevice);
        printMyAssigeDevice(dicMyAssignDevice);
        checkDuplicateAssignDevice(dicMyAssignDevice);

        readPropertyExcel(props, propertySheetName, propertyDataRow, propertyWordKeyToColum, dicMyProperty);
        printMyDevice(dicMyProperty);
        checkDuplicateDevice(dicMyProperty);

        updateDevicePropery(dicMyAssignDevice, dicMyProperty);

        assignDevice(dicMyAssignDevice, dicMySw);
        printMySw(dicMySw);

        saveMyAssignDevice(save, saveWordFromColum, saveWordToColum, dicMySw);

//== finish
        if (isAllPass)
        {
            logger.ZLogInformation($"== [Congratulations!] すべての処理をパスしました ==");
        }
        logger.ZLogInformation($"==== tool finish ====");
    }

    private void readDefinitionExcel(string excel, string sheetName, int firstDataRow, string wordKeyToColum, Dictionary<string, MyAssignDevice> dic)
    {
        logger.ZLogInformation($"== start Definitionファイルの読み込み ==");
        bool isError = false;
        Dictionary<string, int> dicKeyToColumn = new Dictionary<string, int>();
        foreach (var keyAndValue in wordKeyToColum.Split(','))
        {
            string[] item = keyAndValue.Split('/');
            dicKeyToColumn.Add(item[0], int.Parse(item[1]));
        }
        using FileStream fsExcel = new FileStream(excel, FileMode.Open, FileAccess.Read, FileShare.Read);
        using XLWorkbook xlWorkbookExcel = new XLWorkbook(fsExcel);
        IXLWorksheets sheetsExcel = xlWorkbookExcel.Worksheets;
        foreach (IXLWorksheet? sheet in sheetsExcel)
        {
            if (sheetName.Equals(sheet.Name))
            {
                int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
                logger.ZLogInformation($"excel:{excel},シート名:{sheet.Name}, 最後の行:{lastUsedRowNumber}, {wordKeyToColum}");

                for (int r = firstDataRow; r < lastUsedRowNumber + 1; r++)
                {
                    MyAssignDevice ad = new MyAssignDevice();
                    foreach (var key in dicKeyToColumn.Keys)
                    {
                        var property = typeof(MyAssignDevice).GetProperty(key);
                        if (property == null)
                        {
                            isError = true;
                            logger.ZLogError($"property is NULL  at sheet:{sheet.Name} row:{r} key:{key}");
                            continue;
                        }

                        IXLCell cellColumn = sheet.Cell(r, dicKeyToColumn[key]);
                        switch (cellColumn.DataType)
                        {
                            case XLDataType.Text:
                                var pt = property.PropertyType;
                                if (pt.IsGenericType && pt.GetGenericTypeDefinition() == typeof(List<>))
                                {
                                    var et = pt.GetGenericArguments()[0];
                                    if (et == typeof(MyDevice))
                                    {
                                        var word = cellColumn.GetValue<string>();
                                        foreach (var device in word.Split('|'))
                                        {
                                            MyDevice d = new MyDevice();
                                            d.deviceNumber = device;
                                            List<MyDevice> list = (List<MyDevice>)property.GetValue(ad);
                                            list.Add(d);
                                        }
                                    }
                                }
                                else
                                {
                                    property.SetValue(ad, cellColumn.GetValue<string>());
                                }
                                break;
                            case XLDataType.Number:
                                property.SetValue(ad, cellColumn.GetValue<int>().ToString());
                                break;
                            case XLDataType.Blank:
                                logger.ZLogTrace($"cell is Blank type at sheet:{sheet.Name} row:{r}");
                                break;
                            default:
                                logger.ZLogError($"cell is NOT type ( Text | Number | Blank) at sheet:{sheet.Name} row:{r}");
                                continue;
                        }
                    }
                    dic.Add(ad.groupKey, ad);
                }
            }
            else
            {
                logger.ZLogTrace($"Miss {sheet.Name}");
            }
        }

        if (!isError)
        {
            logger.ZLogInformation($"[OK] readDefinitionExcel()は正常に処理できました");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] readDefinitionExcel()でエラーが発生しました");
        }
        logger.ZLogInformation($"== end Definitionファイルの読み込み ==");
    }

    private void readPropertyExcel(string excel, string sheetName, int firstDataRow, string wordKeyToColum, Dictionary<string, MyDevice> dic)
    {
        logger.ZLogInformation($"== start Propertyファイルの読み込み ==");
        bool isError = false;
        Dictionary<string, int> dicKeyToColumn = new Dictionary<string, int>();
        foreach (var keyAndValue in wordKeyToColum.Split(','))
        {
            string[] item = keyAndValue.Split('/');
            dicKeyToColumn.Add(item[0], int.Parse(item[1]));
        }
        using FileStream fsExcel = new FileStream(excel, FileMode.Open, FileAccess.Read, FileShare.Read);
        using XLWorkbook xlWorkbookExcel = new XLWorkbook(fsExcel);
        IXLWorksheets sheetsExcel = xlWorkbookExcel.Worksheets;
        foreach (IXLWorksheet? sheet in sheetsExcel)
        {
            if (sheetName.Equals(sheet.Name))
            {
                int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
                logger.ZLogInformation($"excel:{excel},シート名:{sheet.Name}, 最後の行:{lastUsedRowNumber}, {wordKeyToColum}");

                for (int r = firstDataRow; r < lastUsedRowNumber + 1; r++)
                {
                    MyDevice de = new MyDevice();
                    foreach (var key in dicKeyToColumn.Keys)
                    {
                        var property = typeof(MyDevice).GetProperty(key);
                        if (property == null)
                        {
                            isError = true;
                            logger.ZLogError($"property is NULL  at sheet:{sheet.Name} row:{r} key:{key}");
                            continue;
                        }

                        IXLCell cellColumn = sheet.Cell(r, dicKeyToColumn[key]);
                        switch (cellColumn.DataType)
                        {
                            case XLDataType.Text:
                                property.SetValue(de, cellColumn.GetValue<string>());
                                break;
                            case XLDataType.Number:
                                property.SetValue(de, cellColumn.GetValue<int>().ToString());
                                break;
                            case XLDataType.Blank:
                                logger.ZLogTrace($"cell is Blank type at sheet:{sheet.Name} row:{r}");
                                break;
                            default:
                                logger.ZLogError($"cell is NOT type ( Text | Number | Blank) at sheet:{sheet.Name} row:{r}");
                                continue;
                        }
                    }
                    dic.Add(de.deviceNumber, de);
                }
            }
            else
            {
                logger.ZLogTrace($"Miss {sheet.Name}");
            }
        }

        if (!isError)
        {
            logger.ZLogInformation($"[OK] readPropertyExcel()は正常に処理できました");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] readPropertyExcel()でエラーが発生しました");
        }
        logger.ZLogInformation($"== end Propertyファイルの読み込み ==");
    }

    private void updateDevicePropery(Dictionary<string, MyAssignDevice> dicMyAssignDevice, Dictionary<string, MyDevice> dicMyProperty)
    {
        logger.ZLogInformation($"== start デバイス情報のアップデート ==");
        bool isError = false;

        Dictionary<string, List<string>> dicRosette = new Dictionary<string, List<string>>();
        int count = 1;
        foreach (var propKey in dicMyProperty.Keys)
        {
            string rosette = dicMyProperty[propKey].rosette;
            if (!rosette.Equals(""))
            {
                if (dicRosette.Keys.Contains(rosette))
                {
                    dicRosette[rosette].Add(dicMyProperty[propKey].deviceNumber);
                }
                else
                {
                    List<string> listRosette = new List<string>();
                    listRosette.Add(dicMyProperty[propKey].deviceNumber);
                    dicRosette.Add(rosette, listRosette);
                }
            }
            count++;
        }

        foreach (var key in dicMyAssignDevice.Keys)
        {
            var assignDevice = dicMyAssignDevice[key];
            List<MyDevice> listDevice = new List<MyDevice>();
            convertAssignDeviceToList(assignDevice.sw, listDevice);
            convertAssignDeviceToList(assignDevice.ap, listDevice);
            convertAssignDeviceToList(assignDevice.printer, listDevice);
            convertAssignDeviceToList(assignDevice.mfp, listDevice);
            convertAssignDeviceToList(assignDevice.ocr, listDevice);
            convertAssignDeviceToList(assignDevice.other, listDevice);

            foreach (var de in listDevice)
            {
                if (dicMyProperty.ContainsKey(de.deviceNumber))
                {
                    var device = dicMyProperty[de.deviceNumber];
                    de.floor = device.floor;
                    de.rackName = device.rackName;
                    de.roomName = device.roomName;
                    de.deviceName = device.deviceName;
                    de.modelName = device.modelName;
                    de.portName = device.portName;
                    de.cableName = device.cableName;
                    de.connectorName = device.connectorName;
                    if (dicRosette.Keys.Contains(device.rosette))
                    {
                        var rosetteIndex = dicRosette[device.rosette].IndexOf(device.deviceNumber) + 1;
                        de.hostName = string.Format("{0:G}{1:00}", device.rosette, rosetteIndex);
                        logger.ZLogTrace($"[rosette] deviceNumber:{device.deviceNumber},rosette:{de.hostName}");
                    }
                    else
                    {
                        de.hostName = device.hostName;
                    }
                }
                else
                {
                    isError = true;
                    logger.ZLogError($"[NG] アサイン情報に{de.deviceNumber}が見つかりませんでした");
                }
            }
        }

        if (!isError)
        {
            logger.ZLogInformation($"[OK] デバイス情報のアップデートは正常に処理できました");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] デバイス情報のアップデートでエラーが発生しました");
        }
        logger.ZLogInformation($"== end デバイス情報のアップデート ==");
    }

    private void assignDevice(Dictionary<string, MyAssignDevice> dicMyAssignDevice, Dictionary<string, List<MySw>> dicMySw)
    {
        logger.ZLogInformation($"== start ポートのアサイン ==");
        bool isError = false;

        foreach (var key in dicMyAssignDevice.Keys)
        {
            var assignDevice = dicMyAssignDevice[key];
            List<MyDevice> listAp = assignDevice.ap.ToList<MyDevice>();
            List<MyDevice> listDevice = new List<MyDevice>();
            convertAssignDeviceToList(assignDevice.printer, listDevice);
            convertAssignDeviceToList(assignDevice.mfp, listDevice);
            convertAssignDeviceToList(assignDevice.ocr, listDevice);
            convertAssignDeviceToList(assignDevice.other, listDevice);
            var countSW = convertAssignDeviceToCount(assignDevice.sw);
            var countAp = listAp.Count;
            var countDevice = listDevice.Count;
            var countAssign = countAp + countDevice;
            var countMaxPort = countSW * 12;
//            logger.ZLogDebug($"assignDevice() key:{key}, countSW:{countSW}, countDevice:{countAssign}");

            if (countMaxPort >= countAssign)
            {
                List<MySw> listMySw = new List<MySw>();
                for (int i = 0; i < countSW; i++)
                {
                    MySw mysw = new MySw();
                    mysw.id = i + 1;
                    mysw.sw = assignDevice.sw[i];
                    listMySw.Add(mysw);
                }
                for (int i = 0; i < countAp; i++)
                {
                    calcAssign(countSW, i + 1, out int targetPort, out int targetSwNumber);
                    listMySw[targetSwNumber - 1].ports[targetPort - 1] = listAp[i];
                }

                listDevice.Reverse();
                for (int i = 0; i < countDevice; i++)
                {
                    calcAssignDescending(countSW, i + 1, out int targetPort, out int targetSwNumber);
                    listMySw[targetSwNumber - 1].ports[targetPort - 1] = listDevice[i];
                }
                dicMySw.Add(assignDevice.rackName, listMySw);
            }
            else
            {
                isError = true;
                logger.ZLogError($"[NG] アサインしたい台数{countAssign}が収容台数{countMaxPort}を超えました");
            }
        }

        if (!isError)
        {
            logger.ZLogInformation($"[OK] ポートのアサインは正常に処理できました");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] ポートのアサインでエラーが発生しました");
        }
        logger.ZLogInformation($"== end ポートのアサイン ==");
    }

    private void checkDuplicateAssignDevice(Dictionary<string, MyAssignDevice> dic)
    {
        logger.ZLogInformation($"== start アサインデバイスの重複の確認 ==");
        bool isError = false;
        Dictionary<string, string> dicCheck = new Dictionary<string, string>();
        foreach (var key in dic.Keys.ToList())
        {
            MyAssignDevice ad = dic[key];
            try
            {
                foreach (var item in ad.sw)
                {
                    if (!item.deviceNumber.Equals("zero"))
                    {
                        dicCheck.Add(item.deviceNumber, ad.floor);
                    }
                }
                foreach (var item in ad.ap)
                {
                    if (!item.deviceNumber.Equals("zero"))
                    {
                        dicCheck.Add(item.deviceNumber, ad.floor);
                    }
                }
                foreach (var item in ad.printer)
                {
                    if (!item.deviceNumber.Equals("zero"))
                    {
                        dicCheck.Add(item.deviceNumber, ad.floor);
                    }
                }
                foreach (var item in ad.mfp)
                {
                    if (!item.deviceNumber.Equals("zero"))
                    {
                        dicCheck.Add(item.deviceNumber, ad.floor);
                    }
                }
                foreach (var item in ad.ocr)
                {
                    if (!item.deviceNumber.Equals("zero"))
                    {
                        dicCheck.Add(item.deviceNumber, ad.floor);
                    }
                }
                foreach (var item in ad.other)
                {
                    if (!item.deviceNumber.Equals("zero"))
                    {
                        dicCheck.Add(item.deviceNumber, ad.floor);
                    }
                }
            }
            catch (System.ArgumentException)
            {
                isError = true;
                logger.ZLogError($"重複エラー 階数:{ad.floor},ラック:{ad.rackName}");
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        if (!isError)
        {
            logger.ZLogInformation($"[OK] アサインデバイスの重複はありませんでした");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] アサインデバイスの重複が発見されました");
        }
        logger.ZLogInformation($"== end アサインデバイスの重複の確認 ==");
    }

    private void checkDuplicateDevice(Dictionary<string, MyDevice> dic)
    {
        logger.ZLogInformation($"== start デバイスの重複の確認 ==");
        bool isError = false;
        Dictionary<string, string> dicCheck = new Dictionary<string, string>();
        foreach (var key in dic.Keys)
        {
            MyDevice de = dic[key];
            try
            {
                dicCheck.Add(de.deviceNumber, "");
            }
            catch (System.ArgumentException)
            {
                isError = true;
                logger.ZLogError($"重複エラー 識別子:{de.deviceNumber}");
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        if (!isError)
        {
            logger.ZLogInformation($"[OK] デバイスの重複はありませんでした");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] デバイスの重複が発見されました");
        }
        logger.ZLogInformation($"== end デバイスの重複の確認 ==");
    }

    private void saveMyAssignDevice(string save, string saveWordFromColum, string saveWordToColum, Dictionary<string, List<MySw>> dic)
    {
        logger.ZLogInformation($"== start ファイルの新規作成 ==");
        bool isError = false;
        Dictionary<int, MySaveFile> dicWordFromColum = new Dictionary<int, MySaveFile>();
        foreach (var keyAndValue in saveWordFromColum.Split(','))
        {
            string[] item = keyAndValue.Split('/');
            MySaveFile file = new MySaveFile();
            file.column = int.Parse(item[0]);
            file.property = item[1];
            file.name = item[2];
            dicWordFromColum.Add(file.column, file);
        }
        Dictionary<int, MySaveFile> dicWordToColum = new Dictionary<int, MySaveFile>();
        foreach (var keyAndValue in saveWordToColum.Split(','))
        {
            string[] item = keyAndValue.Split('/');
            MySaveFile file = new MySaveFile();
            file.column = int.Parse(item[0]);
            file.property = item[1];
            file.name = item[2];
            dicWordToColum.Add(file.column, file);
        }

        const int SAVE_COLUMN_DATETIME = 1;
        const int SAVE_FIRST_ROW_DATETIME = 1;
        const int SAVE_FIRST_ROW_HEADER = 3;
        const int SAVE_FIRST_ROW_DATA = SAVE_FIRST_ROW_HEADER + 1;
        using var workbook = new XLWorkbook();
        var keys = dic.Keys.ToImmutableList();
        foreach (var key in keys)
        {
            foreach (var sw in dic[key])
            {
                var worksheet = workbook.AddWorksheet(sw.sw.deviceNumber);

                worksheet.Cell(SAVE_FIRST_ROW_DATETIME, SAVE_COLUMN_DATETIME).SetValue(convertDateTimeToJst(DateTime.Now));

                // FROM
                foreach (var excelkey in dicWordFromColum.Keys)
                {
                    worksheet.Cell(SAVE_FIRST_ROW_HEADER, dicWordFromColum[excelkey].column).SetValue(dicWordFromColum[excelkey].name);
                }
                // TO
                foreach (var excelkey in dicWordToColum.Keys)
                {
                    worksheet.Cell(SAVE_FIRST_ROW_HEADER, dicWordToColum[excelkey].column).SetValue(dicWordToColum[excelkey].name);
                }

                int row = SAVE_FIRST_ROW_DATA;
                for (int i = 0; i < sw.ports.Length; i++)
                {
                    // FROM
                    MyDevice swde = sw.sw;
                    foreach (var excelkey in dicWordFromColum.Keys)
                    {
//                        worksheet.Cell(row, dicWordFromColum[excelkey].column).SetValue(dicWordFromColum[excelkey].property);
                        var property = typeof(MyDevice).GetProperty(dicWordFromColum[excelkey].property);
                        worksheet.Cell(row, dicWordFromColum[excelkey].column).SetValue(property.GetValue(swde).ToString());
                    }

                    // TO
                    MyDevice de = sw.ports[i];
                    foreach (var excelkey in dicWordToColum.Keys)
                    {
                        var property = typeof(MyDevice).GetProperty(dicWordToColum[excelkey].property);
                        worksheet.Cell(row, dicWordToColum[excelkey].column).SetValue(property.GetValue(de).ToString());
                    }
                    row++;
                }

                for (int column = 0; column < worksheet.LastColumnUsed().ColumnNumber(); column++)
                {
                    worksheet.Column(column + 1).Style.Font.SetFontSize(10);
                    worksheet.Column(column + 1).Style.Font.SetFontName("Meiryo UI");
                    worksheet.Column(column + 1).Width = 6;
//                    worksheet.Column(column + 1).AdjustToContents().Style.Font.SetFontName("Meiryo UI");
                }
            }
            workbook.SaveAs(save);
        }

        if (!isError)
        {
            logger.ZLogInformation($"[OK] saveMySiteStatus()は正常に処理できました");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] saveMySiteStatus()でエラーが発生しました");
        }
        logger.ZLogInformation($"== end ファイルの新規作成 ==");
    }

    private void calcAssign(int countSw, int target, out int targetPort, out int targetSwNumber)
    {
        targetPort = -1;
        targetSwNumber = -1;
        targetPort = target / countSw;
        targetSwNumber = ( target % countSw );
        if (targetSwNumber == 0)
        {
            targetSwNumber = countSw;
        }
        if (countSw > targetSwNumber)
        {
            targetPort++;
        }

        logger.ZLogTrace($"calcAssign() countSw:{countSw}, target:{target}, swNumber:{targetSwNumber}, port:{targetPort}");
    }

    private void calcAssignDescending(int countSw, int target, out int targetPort, out int targetSwNumber)
    {
        targetPort = -1;
        targetSwNumber = -1;
        targetPort = target / countSw;
        targetSwNumber = target % countSw;

        if (countSw == 1)
        {
            targetSwNumber = targetSwNumber + countSw;
        }
        else
        {
            targetSwNumber = countSw - targetSwNumber + 1;
            if (targetSwNumber > countSw)
            {
                targetSwNumber = targetSwNumber - countSw;
            }
        }
        if (targetSwNumber == 1)
        {
            targetPort = targetPort - 1;
        }
        targetPort = 12 - targetPort;
        
        logger.ZLogTrace($"calcAssign() countSw:{countSw}, target:{target}, swNumber:{targetSwNumber}, port:{targetPort}");
    }

    private void convertAssignDeviceToList(List<MyDevice> assignDevice, List<MyDevice> device)
    {
        if (assignDevice.Count == 1)
        {
            if (assignDevice[0].deviceNumber == "zero")
            {
                return;
            }
        }
        device.AddRange(assignDevice);
    }

    private int convertAssignDeviceToCount(List<MyDevice> assignDevice)
    {
        if (assignDevice.Count == 1)
        {
            if (assignDevice[0].deviceNumber == "zero")
            {
                return 0;
            }
        }
        return assignDevice.Count;
    }

    private string convertMySw(MyDevice[] devices)
    {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < devices.Length; i++)
        {
            sb.Append(devices[i].deviceNumber);
            if (i < devices.Length - 1)
            {
                sb.Append("|");
            }
        }
        return sb.ToString();
    }
    private string convertMyDevice(List<MyDevice> device)
    {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < device.Count; i++)
        {
            sb.Append(device[i].deviceNumber);
            if (i < device.Count - 1)
            {
                sb.Append("|");
            }
        }
        return sb.ToString();
    }

    private void printMyAssigeDevice(Dictionary<string, MyAssignDevice> dic)
    {
        logger.ZLogTrace($"== start print ==");
        foreach (var ad in dic.Values.ToList())
        {
            logger.ZLogTrace($"キー:{ad.groupKey},階数:{ad.floor},ラック:{ad.rackName},SW:{convertMyDevice(ad.sw)},AP:{convertMyDevice(ad.ap)},PR:{convertMyDevice(ad.printer)},MFP:{convertMyDevice(ad.mfp)},OCR:{convertMyDevice(ad.ocr)},Other:{convertMyDevice(ad.other)}");
        }
        logger.ZLogTrace($"== end print ==");
    }

    private void printMyDevice(Dictionary<string, MyDevice> dic)
    {
        logger.ZLogTrace($"== start print ==");
        foreach (var de in dic.Values)
        {
            logger.ZLogTrace($"識別子:{de.deviceNumber},階数:{de.floor},ラック:{de.rackName},部屋名:{de.roomName},デバイス名:{de.deviceName},モデル名:{de.modelName},port:{de.portName},種別:{de.cableName},コネクタ:{de.connectorName},ローゼット:{de.rosette}");
        }
        logger.ZLogTrace($"== end print ==");
    }

    private void printMySw(Dictionary<string, List<MySw>> dic)
    {
        logger.ZLogTrace($"== start print ==");
        foreach (var list in dic.Values.ToList())
        {
            foreach (var sw in list)
            {
                logger.ZLogTrace($"キー:{sw.id},deviceNumber:{sw.sw.deviceNumber},port:{convertMySw(sw.ports)}");
            }
        }
        logger.ZLogTrace($"== end print ==");
    }

    private string convertDateTimeToJst(DateTime day)
    {
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, jstTimeZoneInfo).ToString("yyyy/MM/dd HH:mm");
    }

    private string getMyFileVersion()
    {
        System.Diagnostics.FileVersionInfo ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
        return ver.InternalName + "(" + ver.FileVersion + ")";
    }
}

//==
public class MyConfig
{
    public int DefinitionDataRow {get; set;} = -1;
    public string DefinitionSheetName {get; set;} = "";
    public string DefinitionWordKeyToColum {get; set;} = "";
    public string DefinitionWordNothing {get; set;} = "";
    public int PropertyDataRow {get; set;} = -1;
    public string PropertySheetName {get; set;} = "";
    public string PropertyWordKeyToColum {get; set;} = "";
    public string SaveWordFromColum {get; set;} = "";
    public string SaveWordToColum {get; set;} = "";
}

public class MySaveFile
{
    public string property { set; get; } = "";
    public int column { set; get; } = -1;
    public string name { set; get; } = "";
}

public class MySw
{
    public int id { set; get; } = -1;
    public MyDevice sw { set; get; } = new MyDevice();
    public MyDevice[] ports { set; get; } = {new MyDevice(),new MyDevice(),new MyDevice(),new MyDevice(),new MyDevice(),new MyDevice()
                                            ,new MyDevice(),new MyDevice(),new MyDevice(),new MyDevice(),new MyDevice(),new MyDevice() };
}
public class MyDevice
{
    public string deviceNumber { set; get; } = "";
    public string floor { set; get; } = "";
    public string rackName { set; get; } = "";
    public string deviceName { set; get; } = "";
    public string hostName { set; get; } = "";
    public string roomName { set; get; } = "";
    public string modelName { set; get; } = "";
    public string portName { set; get; } = "";
    public string cableName { set; get; } = "";
    public string connectorName { set; get; } = "";
    public string rosette { set; get; } = "";
}
public class MyAssignDevice
{
    public string groupKey { set; get; } = "";
    public string floor { set; get; } = "";
    public string rackName { set; get; } = "";
    public List<MyDevice> sw { set; get; } = new List<MyDevice>();
    public List<MyDevice> ap { set; get; } = new List<MyDevice>();
    public List<MyDevice> printer { set; get; } = new List<MyDevice>();
    public List<MyDevice> mfp { set; get; } = new List<MyDevice>();
    public List<MyDevice> ocr { set; get; } = new List<MyDevice>();
    public List<MyDevice> other { set; get; } = new List<MyDevice>();
}