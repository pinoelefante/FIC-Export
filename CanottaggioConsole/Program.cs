using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace CanottaggioConsole
{
    class Program
    {
        private static Dictionary<string, string> file_config = new Dictionary<string, string>()
            {
                { "tbl_pettorale","Pettorale"},
                { "pg_barca&pg_cate&pg_sex", "Categoria2"},
                { "Categoria_int", "Categoria"},
                { "tbl_sex", "Sex"},
                { "tbl_nazione", "Nazione"},
                { "id_squadra", "id_squadra"},
                { "tbl_corsia", "Acqua"},
                { "id_batteria", "Batteria"},
                { "1", "Atleta1"},
                { "2", "Atleta2"},
                { "3", "Atleta3"},
                { "4", "Atleta4"},
                { "5", "Atleta5"},
                { "6", "Atleta6"},
                { "7", "Atleta7"},
                { "8", "Atleta8"},
                { "9", "Atleta9"},
            };
        private static Dictionary<string, Tuple<string, string>> categories = new Dictionary<string, Tuple<string, string>>();
        private static Dictionary<string, Tuple<string, string>> nations = new Dictionary<string, Tuple<string, string>>();
        private static Dictionary<string, string> teams = new Dictionary<string, string>();
        static void Main(string[] args)
        {
            var filename = string.Empty;
            var separator = string.Empty;
            var exportType = string.Empty;
            var national = false;
            var title = string.Empty;
            var base_tvg = string.Empty;
            bool startFromShell = args.Length == 6;

            AppCredits();

            if (args.Length != 6)
            {
                filename = getStringAsPathFromCommandLine("Percorso file: ");
                separator = getStringFromCommandLine("Separatore CSV [default = ;]: ", true, false, string.Empty, ";", ",");
                if(string.IsNullOrEmpty(separator))
                {
                    separator = ";";
                    Console.WriteLine("Verra' utilizzato il carattere ; come separatore CSV");
                }
                exportType = getStringFromCommandLine("Export [mispeaker/tvg/atleti/tutto]: ", true, false, "mispeaker", "tvg", "atleti", "tutto");
                national = getStringAsBoolFromCommandLine("Nazionale [true/false]: ", true, new string[] { "t", "true", "v", "vero", "s", "si", "nazionale" }, new string[] { "internazionale", "n", "no", "f", "false", "falso" });
                title = getStringFromCommandLine("Titolo manistazione: ", false, true);
                base_tvg = getStringAsPathFromCommandLine("Percorso cartella di TVG: ", false);
            }
            else
            {
                filename = args[0];
                separator = args[1];
                exportType = args[2];
                national = Boolean.Parse(args[3]);
                title = args[4];
                base_tvg = args[5];
            }
            loadFileMapping();
            var content = readFile(filename);
            if (content == null || content.Length <= 1)
            {
                Console.WriteLine("Il file non esiste o è vuoto");
                Console.ReadKey();
                return;
            }
            var contentDictionary = parseFile(content, separator);
            switch(exportType.ToLower())
            {
                case "mispeaker":
                    MiSpeakerConverter(contentDictionary, national);
                    break;
                case "tvg":
                    TVGConverter(contentDictionary, national, title);
                    break;
                case "atleti":
                    AthletesCreditsConvert(contentDictionary);
                    break;
                case "tutto":
                    MiSpeakerConverter(contentDictionary, national);
                    TVGConverter(contentDictionary, national, title);
                    AthletesCreditsConvert(contentDictionary);
                    break;
            }
            if (!national)
                VerifyFlagsInt(contentDictionary, base_tvg);
            if (!startFromShell)
            {
                Console.WriteLine("\nFile esportato/i. Premere un tasto per chiudere la finestra");
                Console.ReadKey();
            }
        }
        private static void loadFileMapping()
        {
            var configLines = readFile("fileconfig.csv");
            if (configLines == null || configLines.Count() <= 1)
            {
                Console.WriteLine("\nMapping non trovato. Verra' usato il mapping di default");
                return;
            }
            Console.WriteLine($"Ci sono {configLines.Count()} righe nel file di mapping");
            file_config.Clear();
            var splitVal = new char[] { ';' };
            foreach(var line in configLines)
            {
                var values = line.Split(splitVal, StringSplitOptions.None);
                file_config.Add(values[1], values[0]);
            }
        }
        private static string getStringFromCommandLine(string text, bool lower=true, bool acceptAll = false, params string[] validValues)
        {
            var line = string.Empty;
            do
            {
                Console.Write(text);
                line = Console.ReadLine().Trim();
                if (lower)
                    line = line.ToLower();
            } while (!isValidValue(line, validValues) && !acceptAll);
            return line;
        }
        private static string getStringAsPathFromCommandLine(string text, bool file = true)
        {
            var line = string.Empty;
            do
            {
                Console.Write(text);
                line = Console.ReadLine().Trim();
                if (file && File.Exists(line))
                    break;
                if (!file && Directory.Exists(line))
                    break;
            } while (true);
            return line;
        }
        private static bool getStringAsBoolFromCommandLine(string text, bool lower, string[] positiveValues, string[] negativeValues)
        {
            var line = string.Empty;
            do
            {
                Console.Write(text);
                line = Console.ReadLine().Trim();
                if (lower)
                    line = line.ToLower();
                if (isValidValue(line, positiveValues))
                    return true;
                if (isValidValue(line, negativeValues))
                    return false;
            } while (true);
        }
        private static bool isValidValue(string input, string[] validValues)
        {
            foreach (var vValue in validValues)
                if (vValue.CompareTo(input) == 0)
                    return true;
            return false;
        }
        private static void loadCategories()
        {
            categories.Clear();
            var lines = File.ReadAllLines("categorie.csv");
            var splitVal = new char[] { ';' };
            foreach (var line in lines)
            {
                var values = line.Split(splitVal, StringSplitOptions.None);
                categories.Add(values[0].Trim(), new Tuple<string, string>(values[1].Trim(), values[2].Trim()));
            }
        }
        private static void loadNations()
        {
            nations.Clear();
            var lines = File.ReadAllLines("nazioni.csv");
            var splitVal = new char[] { ';' };
            foreach(var line in lines)
            {
                var values = line.Split(splitVal, StringSplitOptions.None);
                nations.Add(values[1].Trim(), new Tuple<string, string>(values[0].Trim(), values[2].Trim()));
            }
        }
        private static void loadTeams()
        {
            teams.Clear();
            var lines = File.ReadAllLines("teams.csv");
            var splitVal = new char[] { ';' };
            foreach(var line in lines)
            {
                var values = line.Split(splitVal, StringSplitOptions.None);
                teams.Add(values[0], values[1]);
            }
        }
        private static string[] readFile(string path)
        {
            if(File.Exists(path))
                return File.ReadAllLines(path);
            return null;
        }
        private static List<Dictionary<string,string>> parseFile(string[] content, string separator)
        {
            var separatorPar = separator.ToCharArray();
            var headers = content[0].Split(separatorPar, StringSplitOptions.None);
            Dictionary<int, string> header_assoc = new Dictionary<int, string>();
            for (int i = 0;i<headers.Length;i++)
            {
                var h_value = headers[i];
                if(file_config.ContainsKey(h_value))
                    header_assoc.Add(i, h_value);
            }
            List<Dictionary<string, string>> listContent = new List<Dictionary<string, string>>(content.Length);
            for(int i=1;i<content.Length;i++)
            {
                var rowContent = content[i].Split(separatorPar, StringSplitOptions.None);
                var rowDictionary = new Dictionary<string, string>();
                foreach(var index in header_assoc.Keys)
                    rowDictionary.Add(file_config[header_assoc[index]], rowContent[index]);
                listContent.Add(rowDictionary);
            }
            return listContent.OrderBy(x => Int32.Parse(x["Batteria"])).ThenBy(x=> Int32.Parse(x["Acqua"])).ToList();
        }
        private static void MiSpeakerConverter(List<Dictionary<string, string>> fields, bool isNational)
        {
            Console.WriteLine("\nAvvio esportazione MiSpeaker");
            var now = DateTime.Now;
            var filename = string.Format("Export_MiSpeaker_{0}_{1:D2}{2:D2}{3:D4}.csv", (isNational ? "Naz" : "Int"), now.Day, now.Month, now.Year);
            try
            {
                using (var file = new StreamWriter(new FileStream($@"{getDesktopPath()}\{filename}", FileMode.Create), Encoding.UTF8))
                {
                    var buffer = new StringBuilder(16384); //16384 = 16KB
                    if (isNational)
                    {
                        Console.WriteLine("ATTENZIONE: L'export nazionale potrebbe non essere corretto in quanto non testato");
                        buffer.AppendLine($"Batteria;Acqua;Pettorale;Atleta;Societa;Societa1;Atleta1;Atleta2;Atleta3;Atleta4;soc;Categoria[Categoria2];Descr_Cat");
                    }
                    else
                        buffer.AppendLine($"Batteria;Acqua;Pettorale;Flag;Cognome;SecondoNome;Società;Atleta1;Atleta2;Atleta3;Atleta4");
                    foreach (var row in fields)
                    {
                        var isTeam = row.ContainsKey("Atleta3") && !string.IsNullOrEmpty(row["Atleta3"].Trim()); //ci sono più di due atleti
                        if (isNational)
                        {
                            buffer.AppendLine($"{row["Batteria"]};{row["Acqua"]};{row["Pettorale"]};Atleta;Societa;Societa1;{row["Atleta1"].Replace("|", " ")};{row["Atleta2"].Replace("|", " ")};{row["Atleta3"].Replace("|", " ")};{row["Atleta4"].Replace("|", " ")};{row["Atleta5"].Replace("|", " ")};{row["Atleta6"].Replace("|", " ")};{row["Atleta7"].Replace("|", " ")};{row["Atleta8"].Replace("|", " ")};{row["Atleta9"].Replace("|", " ")};soc;{row["Categoria2"]};{getCategoryDescription(row["Categoria2"],row["Categoria"])}");
                        }
                        else
                        {
                            var flag = $@"Flags3D\{(getFlagName(row["Nazione"].Trim()))}.png";
                            var teamName = getTeamNameInt(row["Nazione"].Trim());
                            var surname = getSurnameInt(row["Nazione"].Trim(), isTeam);
                            buffer.AppendLine(string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}", 
                                (row.ContainsKey("Batteria")?row["Batteria"]:""),
                                (row.ContainsKey("Acqua") ? row["Acqua"] : ""),
                                (row.ContainsKey("Pettorale") ? row["Pettorale"] : ""),
                                flag,
                                surname,
                                (!isTeam && row.ContainsKey("Atleta1") ? row["Atleta1"].Replace("|", " ") : ""),
                                teamName,
                                (row.ContainsKey("Atleta1") ? row["Atleta1"].Replace('|',' ') : ""),
                                (row.ContainsKey("Atleta2") ? row["Atleta2"].Replace('|', ' ') : ""),
                                (row.ContainsKey("Atleta3") ? row["Atleta3"].Replace('|', ' ') : ""),
                                (row.ContainsKey("Atleta4") ? row["Atleta4"].Replace('|', ' ') : "")
                            ));
                        }
                    }
                    Console.WriteLine($"Salvataggio file {filename} sul desktop");
                    file.Write(buffer.ToString());
                }
                Console.WriteLine("Esportazione MiSpeaker completata con successo");
            }
            catch(Exception e)
            {
                Console.WriteLine($"Si e' verificato un errore durante l'esportazione di MiSpeaker\n{e.Message}\n{e.StackTrace}");
            }
        }
        private static string getFlagName(string nation)
        {
            if (!nations.Any())
                loadNations();
            var shortNation = getShortSurnameInt(nation);
            if (nations.ContainsKey(nation))
                return nations[nation].Item1;
            else if (nations.ContainsKey(shortNation))
                return nations[shortNation].Item1;
            else
                Console.WriteLine($"Bandiera non trovata: {nation}");
            return nation.Substring(0,3);
        }
        private static string getTeamNameInt(string nation)
        {
            if (!nations.Any())
                loadNations();
            var shortNation = getShortSurnameInt(nation);
            if (nations.ContainsKey(nation))
                return nations[nation].Item2.ToUpper();
            else if (nations.ContainsKey(shortNation))
                return nations[shortNation].Item2.ToUpper();
            else
                Console.WriteLine($"Team - Nazione ({nation}) non trovata - Short ({nation})");
            return string.Empty;
        }
        private static string getSurnameInt(string nation, bool isTeam)
        {
            if (!nations.Any())
                loadNations();
            var shortNation = getShortSurnameInt(nation);
            if (nations.ContainsKey(nation) && isTeam)
                return nations[nation].Item2.ToUpper();
            else if(nations.ContainsKey(shortNation) && isTeam)
                return nations[shortNation].Item2.ToUpper();
            else if(isTeam)
                Console.WriteLine($"Surname - Nazione ({nation}) non trovata - Short ({shortNation})");
            return shortNation;
        }
        private static string getShortSurnameInt(string nation)
        {
            var shortNation = nation.ToString();
            if (shortNation.Where(x => x == ' ').Count() >= 2)
            {
                var parts = shortNation.Split(new char[] { ' ' });
                shortNation = $"{parts[0]} {parts[1].Substring(0, 1)}{parts[2]}";
            }
            return shortNation;
        }
        private static string getCategoryDescription(string catId, string cat2)
        {
            if (!categories.Any())
                loadCategories();
            if (categories.ContainsKey(catId))
                return categories[catId].Item1;
            else if (categories.ContainsKey(cat2))
                return categories[cat2].Item1;
            else
                Console.WriteLine($"Descrizione categoria ({catId}) non trovata. Categoria2 ({cat2})");
            return string.Empty;
        }
        private static string getCategorySex(string catId, string cat2Id)
        {
            if (!categories.Any())
                loadCategories();
            if (categories.ContainsKey(catId))
                return categories[catId].Item2;
            else if (categories.ContainsKey(cat2Id))
                return categories[cat2Id].Item2;
            else
                Console.WriteLine($"Sex non trovato per la categoria ({catId} - {cat2Id})");
            return string.Empty;
        }
        private static string getTeamNameNational(string code)
        {
            if (teams.ContainsKey(code))
                return teams[code];
            else
                Console.WriteLine($"Squadra non trovata ({code})");
            return string.Empty;
        }
        private static char[] splitName = new char[] { '|' };
        private static void TVGConverter(List<Dictionary<string, string>> fields, bool isNational, string title)
        {
            Console.WriteLine("\nAvvio esportazione TVG");
            try
            {
                var groups = fields.GroupBy(x => x["Batteria"]);
                using (var excelPackage = new ExcelPackage())
                {
                    for (int i = groups.Count() - 1; i >= 0; i--)
                    {
                        var group = groups.ElementAt(i);
                        var batteryNum = group.First()["Batteria"];
                        if(string.IsNullOrEmpty(batteryNum))
                        {
                            Console.WriteLine("Verifica che il file csv non contenga righe vuote");
                            continue;
                        }
                        var category = group.First()["Categoria"];
                        var isTeam = !string.IsNullOrEmpty(group.First()["Atleta3"].Trim());
                        var worksheet = excelPackage.Workbook.Worksheets.Add(batteryNum);
                        worksheet.Cells["A1"].Value = title;
                        worksheet.Cells["A2"].Value = $"Starting list {category}";
                        if (isNational)
                        {
                            Console.WriteLine("ATTENZIONE: L'export nazionale potrebbe non essere corretto in quanto non testato");
                            worksheet.Cells["A3"].Value = "Acqua";
                            worksheet.Cells["B3"].Value = "Pettorale";
                            worksheet.Cells["C3"].Value = "Atleta";
                            worksheet.Cells["D3"].Value = "Societa";
                            worksheet.Cells["E3"].Value = "Societa1";
                            worksheet.Cells["F3"].Value = "Atleta1";
                            worksheet.Cells["G3"].Value = "Atleta2";
                            worksheet.Cells["H3"].Value = "Atleta3";
                            worksheet.Cells["I3"].Value = "Atleta4";
                            worksheet.Cells["J3"].Value = "Atleta5";
                            worksheet.Cells["K3"].Value = "Atleta6";
                            worksheet.Cells["L3"].Value = "Atleta7";
                            worksheet.Cells["M3"].Value = "Atleta8";
                            worksheet.Cells["N3"].Value = "Atleta9";
                            worksheet.Cells["O3"].Value = "Descr_Cat";
                            worksheet.Cells["P3"].Value = "Descr_Turno";
                            worksheet.Cells["Q3"].Value = "Categoria"; //Categoria 2
                            worksheet.Cells["R3"].Value = "Gara";
                            worksheet.Cells["S3"].Value = "Cognome";
                            worksheet.Cells["T3"].Value = "Nome";
                            worksheet.Cells["U3"].Value = "Sex";
                            worksheet.Cells["V3"].Value = "Batteria";
                            worksheet.Cells["W3"].Value = "Codsoc";
                            worksheet.Cells["X3"].Value = "corsia";
                            worksheet.Cells["Y3"].Value = "Cognome1";
                            worksheet.Cells["Z3"].Value = "Cognome2";
                            worksheet.Cells["AA3"].Value = "Cognome3";
                            worksheet.Cells["AB3"].Value = "Cognome4";
                            worksheet.Cells["AC3"].Value = "equipaggio";
                            var sex = getCategorySex(group.First()["Categoria2"], group.First()["Categoria"]);
                            for (int j = 0; j < group.Count(); j++)
                            {
                                var atleta = group.ElementAt(j);
                                worksheet.Cells[$"A{4 + j}"].Value = Int32.Parse(atleta["Acqua"]);
                                worksheet.Cells[$"B{4 + j}"].Value = Int32.Parse(atleta["Pettorale"]);
                                worksheet.Cells[$"C{4 + j}"].Value = getSurnameInt(atleta["Nazione"], isTeam);
                                worksheet.Cells[$"D{4 + j}"].Value = isTeam ? "" : atleta["Atleta1"].Replace("|", " ");
                                worksheet.Cells[$"E{4 + j}"].Value = getTeamNameNational(atleta["id_squadra"]);
                                worksheet.Cells[$"F{4 + j}"].Value = atleta["Atleta1"].Replace("|", " ");
                                worksheet.Cells[$"G{4 + j}"].Value = atleta["Atleta2"].Replace("|", " ");
                                worksheet.Cells[$"H{4 + j}"].Value = atleta["Atleta3"].Replace("|", " ");
                                worksheet.Cells[$"I{4 + j}"].Value = atleta["Atleta4"].Replace("|", " ");
                                worksheet.Cells[$"J{4 + j}"].Value = atleta["Atleta5"].Replace("|", " ");
                                worksheet.Cells[$"K{4 + j}"].Value = atleta["Atleta6"].Replace("|", " ");
                                worksheet.Cells[$"L{4 + j}"].Value = atleta["Atleta7"].Replace("|", " ");
                                worksheet.Cells[$"M{4 + j}"].Value = atleta["Atleta8"].Replace("|", " ");
                                worksheet.Cells[$"N{4 + j}"].Value = atleta["Atleta9"].Replace("|", " ") + " (COX)";
                                worksheet.Cells[$"O{4 + j}"].Value = getCategoryDescription(atleta["Categoria2"], atleta["Categoria"]);
                                worksheet.Cells[$"P{4 + j}"].Value = ""; //descr_turno
                                worksheet.Cells[$"Q{4 + j}"].Value = atleta["Categoria2"];
                                worksheet.Cells[$"R{4 + j}"].Value = Int32.Parse(atleta["Batteria"]);
                                worksheet.Cells[$"S{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta1"]) ? atleta["Atleta1"].Split(splitName)[0] : "";
                                worksheet.Cells[$"T{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta1"]) ? atleta["Atleta1"].Split(splitName)[1] : "";
                                worksheet.Cells[$"U{4 + j}"].Value = sex;
                                worksheet.Cells[$"V{4 + j}"].Value = Int32.Parse(atleta["Batteria"]);
                                worksheet.Cells[$"W{4 + j}"].Value = atleta["id_squadra"];
                                worksheet.Cells[$"X{4 + j}"].Value = Int32.Parse(atleta["Acqua"]);
                                worksheet.Cells[$"Y{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta1"])? atleta["Atleta1"].Split(splitName)[0] : "";
                                worksheet.Cells[$"Z{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta2"]) ? atleta["Atleta2"].Split(splitName)[0] : "";
                                worksheet.Cells[$"AA{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta3"]) ? atleta["Atleta3"].Split(splitName)[0] : "";
                                worksheet.Cells[$"AB{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta4"]) ? atleta["Atleta4"].Split(splitName)[0] : "";
                                worksheet.Cells[$"AC{4 + j}"].Value = ""; //equipaggio
                            }
                        }
                        else
                        {
                            worksheet.Cells["A3"].Value = "Acqua";
                            worksheet.Cells["B3"].Value = "Pettorale";
                            worksheet.Cells["C3"].Value = "Flag";
                            worksheet.Cells["D3"].Value = "Atleta";
                            worksheet.Cells["E3"].Value = "Societa";
                            worksheet.Cells["F3"].Value = "Societa1";
                            worksheet.Cells["G3"].Value = "Atleta1";
                            worksheet.Cells["H3"].Value = "Atleta2";
                            worksheet.Cells["I3"].Value = "Atleta3";
                            worksheet.Cells["J3"].Value = "Atleta4";
                            worksheet.Cells["K3"].Value = "Atleta5";
                            worksheet.Cells["L3"].Value = "Atleta6";
                            worksheet.Cells["M3"].Value = "Atleta7";
                            worksheet.Cells["N3"].Value = "Atleta8";
                            worksheet.Cells["O3"].Value = "Atleta9";
                            for (int j = 0; j < group.Count(); j++)
                            {
                                var atleta = group.ElementAt(j);
                                worksheet.Cells[$"A{4 + j}"].Value = Int32.Parse(atleta["Acqua"]);
                                worksheet.Cells[$"B{4 + j}"].Value = Int32.Parse(atleta["Pettorale"]);
                                worksheet.Cells[$"C{4 + j}"].Value = $@"Flags3D\{getFlagName(atleta["Nazione"])}.png";
                                worksheet.Cells[$"D{4 + j}"].Value = getSurnameInt(atleta["Nazione"], isTeam);
                                worksheet.Cells[$"E{4 + j}"].Value = isTeam ? "" : atleta["Atleta1"].Replace("|", " ");
                                worksheet.Cells[$"F{4 + j}"].Value = getTeamNameInt(atleta["Nazione"]);
                                worksheet.Cells[$"G{4 + j}"].Value = atleta.ContainsKey("Atleta1") ? atleta["Atleta1"].Replace("|", " ") : "";
                                worksheet.Cells[$"H{4 + j}"].Value = atleta.ContainsKey("Atleta2") ? atleta["Atleta2"].Replace("|", " ") : "";
                                worksheet.Cells[$"I{4 + j}"].Value = atleta.ContainsKey("Atleta3") ? atleta["Atleta3"].Replace("|", " ") : "";
                                worksheet.Cells[$"J{4 + j}"].Value = atleta.ContainsKey("Atleta4") ? atleta["Atleta4"].Replace("|", " ") : "";
                                worksheet.Cells[$"K{4 + j}"].Value = atleta.ContainsKey("Atleta5") ? atleta["Atleta5"].Replace("|", " ") : "";
                                worksheet.Cells[$"L{4 + j}"].Value = atleta.ContainsKey("Atleta6") ? atleta["Atleta6"].Replace("|", " ") : "";
                                worksheet.Cells[$"M{4 + j}"].Value = atleta.ContainsKey("Atleta7") ? atleta["Atleta7"].Replace("|", " ") : "";
                                worksheet.Cells[$"N{4 + j}"].Value = atleta.ContainsKey("Atleta8") ? atleta["Atleta8"].Replace("|", " ") : "";
                                worksheet.Cells[$"O{4 + j}"].Value = atleta.ContainsKey("Atleta9") ? atleta["Atleta9"].Replace("|", " ") + " (COX)" : "";
                                worksheet.Cells[$"P{4 + j}"].Value = atleta["Categoria"];
                                worksheet.Cells[$"Q{4 + j}"].Value = getCategoryDescription(atleta["Categoria"], atleta["Categoria2"]);
                                worksheet.Cells[$"R{4 + j}"].Value = Int32.Parse(atleta["Batteria"]);
                                worksheet.Cells[$"S{4 + j}"].Value = Int32.Parse(atleta["Acqua"]);
                                worksheet.Cells[$"T{4 + j}"].Value = 1;
                            }
                        }
                    }
                    var dateNow = DateTime.Now;
                    var filename = string.Format("Export_TVG_{0}_{1:D2}{2:D2}{3:D4}.xlsx", (isNational ? "Naz" : "Int"), dateNow.Day, dateNow.Month, dateNow.Year);
                    var fileinfo = new FileInfo($@"{getDesktopPath()}\{filename}");
                    excelPackage.SaveAs(fileinfo);
                    Console.WriteLine($"Il file {filename} e' stato salvato sul desktop");
                    Console.WriteLine("Esportazione TVG completata con successo");
                }
            }
            catch(Exception e)
            {
                Console.WriteLine($"Si e' verificato un errore durante l'esportazione di TVG\n{e.Message}\n{e.StackTrace}");
            }
        }
        private static void AthletesCreditsConvert(List<Dictionary<string, string>> fields)
        {
            Console.WriteLine("\nCarico la lista degli atleti");
            var list = new List<string>();
            foreach(var dict in fields)
            {
                for(int i=1;i<=9 && dict.ContainsKey($"Atleta{i}"); i++)
                {

                    if(!string.IsNullOrEmpty(dict[$"Atleta{i}"]))
                        list.Add(dict[$"Atleta{i}"].Trim().Replace('|', ' '));
                    else
                        break;
                }
            }
            list = list.OrderBy(x => x).ToList();
            var now = DateTime.Now;
            var filename = string.Format("Export_Credits_{0:D2}{1:D2}{2:D4}.xlsx", now.Day, now.Month, now.Year);

            using (var excel = new ExcelPackage())
            
            {
                var sheet = excel.Workbook.Worksheets.Add("Atleti");
                for (int i = 0; i < list.Count; i++)
                    sheet.Cells[$"A{(i + 1)}"].Value = list[i];
                Console.WriteLine($"Salvo la lista degli atleti sul desktop nel file {filename}");
                var fileinfo = new FileInfo($@"{getDesktopPath()}\{filename}");
                excel.SaveAs(fileinfo);
            }
        }
        private static void VerifyFlagsInt(List<Dictionary<string, string>> contents, string folder)
        {
            Console.WriteLine("\nVerifico la presenza dei file delle bandiere");
            Dictionary<string, bool> verifyStatus = new Dictionary<string, bool>();
            foreach(var atl in contents)
            {
                if(atl.ContainsKey("Nazione"))
                {
                    var flag = $@"Flags3D\{getFlagName(atl["Nazione"])}.png";
                    if(!verifyStatus.ContainsKey(flag))
                    {
                        var path = Path.Combine(folder, flag);
                        var flagExists = File.Exists(path);
                        verifyStatus.Add(flag, flagExists);
                        if (!flagExists)
                            Console.WriteLine($"Bandiera non presente ({flag})");
                    }
                }
            }
        }
        private static void AppCredits()
        {
            Console.WriteLine("Sviluppato da Giuseppe Elefante <giuseppe.elefante90@gmail.com>\nFICr Salerno - A.S.D. Cronometristi Salernitani \"R. Marra\"\n");
        }
        private static string getDesktopPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }
    }
    
}
