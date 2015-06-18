#define USE_MOONSHARP
//#define USE_KOPILUA

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections;
using System.Net;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
#if USE_KOPILUA
using KopiLua;
#endif

namespace convSheet
{

    class Program
    {
        static void Main(string[] args)
        {
#if DEBUG
            Console.WriteLine("Start");
#endif
            try
            {
                (new Program()).MyMain(args);
            }
            catch (ApplicationException ex)
            {
                Console.WriteLine(ex.Message);
            }
            catch (Exception ex) {
                Console.WriteLine(ex.ToString());
            }
#if DEBUG
            Console.WriteLine("End");
            Console.Read();
#endif
        }

        public Program()
        {
        }

#if USE_MOONSHARP

        MoonSharp.Interpreter.Script lua;
        MoonSharp.Interpreter.Table checkSheet, applySheet, checkWorkbook, applyWorkbook, evalPath, replaceTable, infoTable;
        MoonSharp.Interpreter.Closure caller, replaceTags;


        void loadLua()
        {
            lua = new MoonSharp.Interpreter.Script();
            checkSheet = MoonSharp.Interpreter.DynValue.NewTable(lua).Table;
            applySheet = MoonSharp.Interpreter.DynValue.NewTable(lua).Table;
            checkWorkbook = MoonSharp.Interpreter.DynValue.NewTable(lua).Table;
            applyWorkbook = MoonSharp.Interpreter.DynValue.NewTable(lua).Table;
            evalPath = MoonSharp.Interpreter.DynValue.NewTable(lua).Table;
            replaceTable = MoonSharp.Interpreter.DynValue.NewTable(lua).Table;
            infoTable = MoonSharp.Interpreter.DynValue.NewTable(lua).Table;
            lua.Globals["checkSheet"] = checkSheet;
            lua.Globals["applySheet"] = applySheet;
            lua.Globals["checkWorkbook"] = checkSheet;
            lua.Globals["applyWorkbook"] = applySheet;
            lua.Globals["evalPath"] = evalPath;
            lua.Globals["info"] = infoTable;
            lua.Globals["crc32hash"] = (Func<string, int>)CRC32.ToHash;
            lua.Globals["fileget"] = (Func<string, BinaryData>)LoadAllBytesForLua;
            lua.Globals["httpget"] = (Func<string, byte[]>)LoadWebBytes;
            lua.Globals["tohash"] = (Func<string, int>)CRC32.ToHash;
            caller = lua.DoString(@"
local tracer = function(err)
    print(err)
end
return function(f, ...)
    local b, v = xpcall(f, tracer, ...)
    return v
end
").Function;
            replaceTags = lua.DoString(@"
return function(replaceTable)
    return function(str)
        local r = (str or """"):gsub(""{(.-)}"", replaceTable)
        return r
    end
end
").Function.Call(replaceTable).Function;
            var myAssembly = System.Reflection.Assembly.GetEntryAssembly();
            string assemblyDir = System.IO.Path.GetDirectoryName(myAssembly.Location);
            var libDirPath = System.IO.Path.GetFullPath(assemblyDir + "\\lua");
            ((MoonSharp.Interpreter.Loaders.ScriptLoaderBase)lua.Options.ScriptLoader).ModulePaths = new string[] { libDirPath.Replace('\\','/')+"/?.lua" };
            foreach (var file in System.IO.Directory.GetFiles(assemblyDir + "\\lua", "*.lua", SearchOption.AllDirectories))
            {
                var fullfile = System.IO.Path.GetFullPath(file);
                var filename = fullfile.Substring(libDirPath.Length + 1);
                var modulename = filename.Substring(0, filename.Length-4);
                lua.DoString("require(\""+modulename+"\");");
#if DEBUG
                Console.WriteLine("require "+modulename);
#endif
            }
        }
#endif // use moonsharp
#if USE_KOPILUA
        void loadLua() {
            var lua = Lua.LuaLNewState();
            Lua.LuaNetLoadBuffer(lua, "", 0, "chunk");
            Lua.LuaCall(lua, 0, 0);
        }
#endif // USE_NLUA


        public class Params
        {
            public Dictionary<string, string> dictionary = new Dictionary<string, string>();
            public List<string> list = new List<string>();

            public static Params Parse(string[] args)
            {
                var p = new Params();
                foreach (var s in args)
                {
                    if (s.StartsWith("-"))
                    {
                        var key = s.Substring(1);
                        var value = "";
                        var eqidx = key.IndexOf('=');
                        if (eqidx != -1)
                        {
                            var v = key;
                            key = v.Substring(0, eqidx);
                            value = v.Substring(eqidx+1);
                        }
                        if (p.dictionary.ContainsKey(key))
                            throw new ApplicationException("Parameter overrapping: "+key);
                        p.dictionary[key] = value;
                    }
                    else
                    {
                        p.list.Add(s);
                    }
                }

                return p;
            } 
        }

        void MyMain(string[] args_)
        {
            var args = Params.Parse(args_);
            loadLua();
            if (args.list.Count == 0)
            {
                Console.WriteLine(
                    "convsheet (XLSXFile) (SheetID) (Options)\n" +
                    "\n" +
                    "Options List:\n" +
                    "     -csv: Output CSV files\n" +
                    "-dir=PATH: Set output directory\n");
                return;
            }
            string spreadsheetPath;
            string sheetID = null;
            spreadsheetPath = args.list[0];
            var outputdir = args.dictionary.ContainsKey ("dir") ? args.dictionary ["dir"] + "\\" : "";
            if (args.list.Count > 1)
                sheetID = args.list[1];
            var ext = System.IO.Path.GetExtension(args.list[0]);
            Dictionary<string, List<List<string>>> workbookData;

            var bytes = EvalPath(spreadsheetPath);
            if (bytes == null) throw new ApplicationException("path Evaluate() error");
            Console.WriteLine("Load/Parse spreadsheet");
            workbookData = ParseWorkbookFromXSLX(bytes);
            if (args.dictionary.ContainsKey("csv"))
            {
                var dic = new Dictionary<string, string>();
                foreach (var kv in workbookData)
                {
                    var sheetname = kv.Key;
                    var filebase = spreadsheetPath.Substring(0, spreadsheetPath.Length-ext.Length) + outputdir;
                    var filename = filebase + "_" + sheetname + ".csv";
                    var csv = CSV.ToCSV(kv.Value);
                    dic[filename] = csv;
                }
                foreach (var kv in dic) {
                    Console.WriteLine("Output: "+kv.Key);
                    System.IO.Directory.CreateDirectory (System.IO.Path.GetDirectoryName(kv.Key));
                    if (!System.IO.File.Exists(kv.Key) || System.IO.File.ReadAllText (kv.Key) != kv.Value) {
                        System.IO.File.WriteAllText (kv.Key, kv.Value);
                    }
                }
                return;
            }

            //var workbookLuaTable = WorkbookToLuaTable(workbookData);
            //lua.Globals["workbook"] = workbookLuaTable;
            //lua.Globals["sheet"] = MoonSharp.Interpreter.DynValue.Nil;

            var replaceDic = new Dictionary<string, string>();
            replaceDic["filefull"] = spreadsheetPath;
            replaceDic["file"] = System.IO.Path.GetFileName(spreadsheetPath.Substring(0, spreadsheetPath.Length - System.IO.Path.GetExtension(spreadsheetPath).Length));
            replaceDic["ext"] = System.IO.Path.GetExtension(spreadsheetPath).Substring(1);
            replaceDic["dir"] = System.IO.Path.GetDirectoryName(spreadsheetPath);
            replaceDic["sheet"] = "";

            Console.WriteLine("Apply");
            // apply workbook
            foreach (var k in checkWorkbook.Pairs)
            {
                lua.Globals["workbook"] = WorkbookToLuaTable(workbookData);
                infoTable ["filefull"] = replaceDic ["filefull"];
                infoTable ["file"] = replaceDic ["file"];
                infoTable ["ext"] = replaceDic ["ext"];
                infoTable ["dir"] = replaceDic ["dir"];
                infoTable ["sheet"] = replaceDic ["sheet"];
                var fn = k.Value.Function;
                var res = caller.Call(fn);
                if (res.CastToBool())
                {
                    Console.WriteLine("ApplyWorkbook " + k.Key.ToPrintString());
                    var outputTable = caller.Call(applyWorkbook.Get(k.Key).Function).Table;
                    var newpath = spreadsheetPath + "." + k.Key.String;
                    WriteOutResultTable(outputTable, replaceDic, outputdir);
                }
            }

            // apply sheet
            var workbookLuaTable = WorkbookToLuaTable(workbookData);
            foreach (var p in workbookLuaTable.Pairs)
            {
                var sheetVal = p.Value;

                foreach (var k in checkSheet.Pairs)
                {
                    lua.Globals["workbook"] = workbookLuaTable;
                    lua.Globals["sheet"] = sheetVal;
                    replaceDic["sheet"] = p.Key.String;
                    infoTable ["filefull"] = replaceDic ["filefull"];
                    infoTable ["file"] = replaceDic ["file"];
                    infoTable ["ext"] = replaceDic ["ext"];
                    infoTable ["dir"] = replaceDic ["dir"];
                    infoTable ["sheet"] = replaceDic ["sheet"];
                    var fn = k.Value.Function;
                    var res = caller.Call(fn);
                    if (res.CastToBool())
                    {
                        Console.WriteLine("ApplySheet: " + k.Key.ToPrintString());
                        var outputTable = caller.Call(applySheet.Get(k.Key).Function).Table;
                        var newpath = spreadsheetPath + "." + k.Key.String;
                        WriteOutResultTable(outputTable, replaceDic, outputdir);
                    }
                }
            }
        }

        void WriteOutResultTable(MoonSharp.Interpreter.Table t, Dictionary<string, string> replaceDic, string outputdir)
        {
            foreach(var kv in replaceDic) {
#if DEBUG
                Console.WriteLine("DDic: " + kv.Key + ":" + kv.Value);
#endif
                replaceTable[kv.Key] = kv.Value;
            }
            foreach (var kv in t.Pairs)
            {
                var filename = replaceTags.Call(kv.Key).String;
#if DEBUG
                Console.WriteLine("D1: " + kv.Key);
                Console.WriteLine("D2: " + filename);
#endif
                var value = kv.Value.String;
                if (!string.IsNullOrEmpty(filename))
                {
                    Console.WriteLine("Output: " + outputdir + filename);
                    System.IO.Directory.CreateDirectory (System.IO.Path.GetDirectoryName(outputdir + filename));
                    if (!System.IO.File.Exists(outputdir + filename) || System.IO.File.ReadAllText (outputdir + filename) != value) {
                        System.IO.File.WriteAllText (outputdir + filename, value);
                    }
                }
            }
        }

        byte[] EvalPath(string path)
        {
            foreach (var kv in evalPath.Pairs)
            {
                try
                {
                    var name = kv.Key.String;
                    var f = kv.Value.Function;
                    var ret = caller.Call(f, path);
                    if (!ret.IsNil())
                    {
                        Console.WriteLine("EvalPath: " + name + " (" + path + ")");
                        var obj = ret.ToObject();
                        if (obj is byte[])
                        {
                            var bytes = (byte[])obj;
                            return bytes;
                        }
                        if (obj is BinaryData)
                        {
                            var bd = (BinaryData)obj;
                            var bytes = bd.bytes;
                            return bytes;
                        }
                    }
                }
                catch
                {

                }
            }
            return null;
        }

        MoonSharp.Interpreter.Table WorkbookToLuaTable(Dictionary<string, List<List<string>>> workbookData)
        {
            var sheet = new MoonSharp.Interpreter.Table(lua);
            foreach (var kv in workbookData)
            {
                var sh = kv.Value;
                var sheetForLua = new MoonSharp.Interpreter.Table(lua);
                int Y = 1;
                foreach (var row in sh)
                {
                    var line = new MoonSharp.Interpreter.Table(lua);
                    int X = 1;
                    foreach (var cell in row)
                    {
                        line[X] = cell;
                        X++;
                    }
                    sheetForLua[Y] = line;
                    Y++;
                }
                sheet[kv.Key] = sheetForLua;
            }
            return sheet;
        }

        Dictionary<string, List<List<string>>> ParseWorkbookFromGsheetFile(string path)
        {
            var data = LoadAllText(path);
            var d = parseGSheet(data);
            var wb = LoadGoogleXSLX(d["doc_id"]);
            return ParseWorkbookFromXSLX(wb);
        }
        Dictionary<string, List<List<string>>> ParseWorkbookFromXSLXFile(string path)
        {
            return ParseWorkbookFromXSLX(LoadAllBytes(path));
        }
        Dictionary<string, List<List<string>>> ParseWorkbookFromXSLX(byte[] data)
        {
            var wbResult = new Dictionary<string, List<List<string>>>();

            var pack = NPOI.OpenXml4Net.OPC.OPCPackage.Open(new MemoryStream(data));
            var workbook = new XSSFWorkbook(pack);
            foreach (ISheet sheet in workbook)
            {
                var aa = SheetToCells(sheet);
                Rectize(aa, "");
                
                //wbResult[workbook.GetSheetIndex(sheet).ToString()] = aa;
                wbResult[sheet.SheetName] = aa;
            }
            return wbResult;
        }
        Dictionary<string, List<List<string>>> ParseWorkbookFromCSV(string csv)
        {
            var wbResult = new Dictionary<string, List<List<string>>>();

            var sheet = CSV.ParseCSV(csv);
            Rectize(sheet, "");
            wbResult["0"] = sheet;
            wbResult["main"] = sheet;

            return wbResult;
        }
        Dictionary<string, List<List<string>>> ParseWorkbookFromCSVFile(string path)
        {
            return ParseWorkbookFromCSV(LoadAllText(path));
        }
        void Rectize<T>(List<List<T>> array2, T fillValue)
        {
            var maxcount = array2.Aggregate(0, (a, b) => Math.Max(a, b.Count));
            foreach (var l in array2)
            {
                while (l.Count < maxcount)
                {
                    l.Add(fillValue);
                }
            }
        }

        public class BinaryData {
            public byte[] bytes;
        }
        BinaryData LoadAllBytesForLua(string path)
        {
            MoonSharp.Interpreter.UserData.RegisterType(typeof(BinaryData), MoonSharp.Interpreter.InteropAccessMode.HideMembers);
            return new BinaryData() { bytes = LoadAllBytes(path) };
        }
        byte[] LoadAllBytes(string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var sr = new BinaryReader(fs))
                {
                    var bytes = sr.ReadBytes((int)fs.Length);
                    return bytes;
                }
            }
        }
        string LoadAllText(string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var sr = new StreamReader(fs))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        List<List<string>> SheetToCells(ISheet sheet)
        {
            var eval = sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
            var lines = new List<List<string>>();
            for (int i = 0; i <= sheet.LastRowNum; i++) {
                lines.Add(new List<string>());
            }
            foreach (XSSFRow r in sheet)
            {
                if (r.Cells.Count == 0) continue;
                var line = lines[r.Cells[0].RowIndex];
                for (int i = line.Count; i <= r.LastCellNum; i++)
                    line.Add("");
                var cells = r.Cells;
                foreach (var cel in cells)
                {
                    string output="";
                    if (cel.CellType == CellType.Formula)
                    {
                        var cellValue = eval.Evaluate(cel);
                        switch (cellValue.CellType)
                        {
                            case CellType.Blank:
                                break;
                            case CellType.Boolean:
                                output = cellValue.BooleanValue ? "true" : "false";
                                break;
                            case CellType.Numeric:
                                output = cellValue.NumberValue.ToString();
                                break;
                            case CellType.String:
                                output = cellValue.StringValue;
                                break;
                            case CellType.Formula:
                            case CellType.Error:
                            case CellType.Unknown:
                                throw new Exception("error type " + cellValue.CellType);
                        }
                    }
                    else
                    {
                        output = cel.ToString();
                    }
                    line[cel.ColumnIndex] = output;
                }
            }

            var maxCol = lines.Aggregate(0, (e, f) => Math.Max(e, f.Count));
            foreach (var l in lines)
            {
                while (maxCol > l.Count)
                {
                    l.Add("");
                }
            }

            return lines;
        }
        


        Dictionary<string, string> parseGSheet(string gsheet)
        {
            var dic = new Dictionary<string, string>();
            foreach (Match re in Regex.Matches(gsheet, @"""(.*?)""\s*:\s*""(.*?)""", RegexOptions.Multiline))
            {
                dic[re.Groups[1].Value] = re.Groups[2].Value;
            }
            return dic;
        }

        public string LoadWebText(string url)
        {
            ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(AcceptAllCertifications);
            var request = System.Net.WebRequest.Create(new System.Uri(url));
            request.Credentials = CredentialCache.DefaultCredentials;
            var response = request.GetResponse();
            var dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);
            string responseFromServer = reader.ReadToEnd();
            reader.Close();
            dataStream.Close();
            response.Close();
            return responseFromServer;
        }
        public byte[] LoadWebBytes(string url)
        {
            ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(AcceptAllCertifications);
            var request = System.Net.WebRequest.Create(new System.Uri(url));
            request.Credentials = CredentialCache.DefaultCredentials;
            var response = request.GetResponse();
            var dataStream = response.GetResponseStream();
            var responseData = ReadAllBytes(dataStream);
            dataStream.Close();
            response.Close();
            return responseData;
        }
        public bool AcceptAllCertifications(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certification, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        public byte[] ReadAllBytes(Stream stream)
        {
            var size = 1024 * 16;
            var buf = new byte[size];
            using (var ms = new MemoryStream())
            {
                while (true) {
                    var res = stream.Read(buf, 0, size);
                    if (res > 0)
                    {
                        ms.Write(buf, 0, res);
                    }
                    if (res == 0)
                    {
                        break;
                    }
                }
                return ms.ToArray();
            }
        }

        public byte[] LoadGoogleXSLX(string spreadsheetID)
        {
            var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetID + "/export?format=xlsx&id=" + spreadsheetID;
            return LoadWebBytes(url);
        }
        

        public string LoadGoogleCSV(string spreadsheetID, string sheetID)
        {
            var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetID + "/export?format=csv&id=" + spreadsheetID; // +"&gid=" + sheetID;
            return LoadWebText(url);
        }
    }


}
