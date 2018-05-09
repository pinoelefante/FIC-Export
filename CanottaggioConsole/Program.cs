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
        static void Main(string[] args)
        {
            var filename = string.Empty;
            var separator = string.Empty;
            var exportType = string.Empty;
            var national = false;
            var title = string.Empty;

            credits();

            if (args.Length != 5)
            {
                do
                {
                    Console.Write("Percorso file: ");
                    filename = Console.ReadLine().Trim();
                    if (File.Exists(filename))
                        break;
                    else
                        Console.WriteLine("File inesistente");
                } while (true);
                
                Console.Write("Separatore CSV [default = ;]: ");
                separator = Console.ReadLine().Trim();
                if(string.IsNullOrEmpty(separator))
                {
                    separator = ";";
                    Console.WriteLine("Verra' utilizzato il carattere ; come separatore CSV");
                }

                do
                {
                    Console.Write("Export [mispeaker/tvg/entrambi]: ");
                    exportType = Console.ReadLine().Trim().ToLower();
                    if (exportType.CompareTo("mispeaker") == 0 || exportType.CompareTo("tvg") == 0 || exportType.CompareTo("entrambi")==0)
                        break;
                    else
                        Console.WriteLine("Valore non valido");
                } while (true);

                do
                {
                    Console.Write("Nazionale [true/false]: ");
                    var consoleValue = Console.ReadLine().Trim().ToLower();
                    switch(consoleValue)
                    {
                        case "t":
                        case "true":
                        case "v":
                        case "vero":
                        case "s":
                        case "si":
                        case "nazionale":
                            consoleValue = "true";
                            break;
                        case "internazionale":
                        case "n":
                        case "no":
                        case "f":
                        case "false":
                        case "falso":
                            consoleValue = "false";
                            break;
                        default:
                            continue;
                    }
                    national = Boolean.Parse(consoleValue);
                    break;
                } while (true);

                Console.Write("Nome competizione: ");
                title = Console.ReadLine().Trim();
            }
            else
            {
                filename = args[0];
                separator = args[1];
                exportType = args[2];
                national = Boolean.Parse(args[3]);
                title = args[4];
            }
            Debug.WriteLine($"Comando={filename} {separator} {exportType} {national} {title}");
            
            var content = readFile(filename);
            if (content == null || content.Length <= 1)
            {
                Console.WriteLine("Il file non esiste o è vuoto");
                Console.ReadKey();
                return;
            }
            var contentDictionary = parseFile(content, separator);
            loadCategories();
            loadNations();
            switch(exportType.ToLower())
            {
                case "mispeaker":
                    MiSpeakerConverter(contentDictionary, national);
                    break;
                case "tvg":
                    TVGConverter(contentDictionary, national, title);
                    break;
                case "entrambi":
                    MiSpeakerConverter(contentDictionary, national);
                    TVGConverter(contentDictionary, national, title);
                    break;
            }
            Console.WriteLine("\nFile esportato/i. Premere un tasto per chiudere la finestra");
            Console.ReadKey();
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
            return listContent.OrderBy(x => x["Batteria"]).ThenBy(x=>x["Acqua"]).ToList();
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
                        buffer.AppendLine($"Batteria;Acqua;Pettorale;Atleta;Societa;Societa1;Atleta1;Atleta2;Atleta3;Atleta4;soc;Categoria[Categoria2];Descr_Cat");
                    }
                    else
                        buffer.AppendLine($"Batteria;Acqua;Pettorale;Flag;Cognome;SecondoNome;Società;Atleta1;Atleta2;Atleta3;Atleta4");
                    foreach (var row in fields)
                    {
                        var isTeam = !string.IsNullOrEmpty(row["Atleta3"].Trim()); //ci sono più di due atleti
                        if (isNational)
                        {
                            buffer.AppendLine($"{row["Batteria"]};{row["Acqua"]};{row["Pettorale"]};Atleta;Societa;Societa1;{row["Atleta1"]};{row["Atleta2"]};{row["Atleta3"]};{row["Atleta4"]};soc;{row["Categoria2"]};{getCategoryDescription(row["Categoria2"],row["Categoria"])}");
                        }
                        else
                        {
                            var flag = $@"Flags3D\{(getFlagName(row["Nazione"].Trim()))}.png";
                            var teamName = getTeamName(row["Nazione"].Trim());
                            var surname = getSurname(row["Nazione"].Trim(), isTeam);
                            buffer.AppendLine($"{row["Batteria"]};{row["Acqua"]};{row["Pettorale"]};{flag};{surname};{(isTeam ? "" : row["Atleta1"].Replace("|", " "))};{teamName};{row["Atleta1"].Replace("|", " ")};{row["Atleta2"].Replace("|", " ")};{row["Atleta3"].Replace("|", " ")};{row["Atleta4"].Replace("|", " ")};");
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
            var shortNation = getShortSurname(nation);
            if (nations.ContainsKey(nation))
                return nations[nation].Item1;
            else if (nations.ContainsKey(shortNation))
                return nations[shortNation].Item1;
            else
                Console.WriteLine($"Bandiera non trovata: {nation}");
            return nation.Substring(0,3);
        }
        private static string getTeamName(string nation)
        {
            var shortNation = getShortSurname(nation);
            if (nations.ContainsKey(nation))
                return nations[nation].Item2.ToUpper();
            else if (nations.ContainsKey(shortNation))
                return nations[shortNation].Item2.ToUpper();
            else
                Console.WriteLine($"Team - Nazione ({nation}) non trovata - Short ({nation})");
            return string.Empty;
        }
        private static string getSurname(string nation, bool isTeam)
        {
            var shortNation = getShortSurname(nation);
            if (nations.ContainsKey(nation) && isTeam)
                return nations[nation].Item2.ToUpper();
            else if(nations.ContainsKey(shortNation) && isTeam)
                return nations[shortNation].Item2.ToUpper();
            else if(isTeam)
                Console.WriteLine($"Surname - Nazione ({nation}) non trovata - Short ({shortNation})");
            return shortNation;
        }
        private static string getShortSurname(string nation)
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
            if (categories.ContainsKey(catId))
                return categories[catId].Item1;
            else if (categories.ContainsKey(cat2))
                return categories[cat2].Item1;
            else
                Console.WriteLine($"Descrizione categoria ({catId}) non trovata. Categoria2 ({cat2})");
            return string.Empty;
        }
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
                        var category = group.First()["Categoria"];
                        var isTeam = !string.IsNullOrEmpty(group.First()["Atleta3"].Trim());
                        var worksheet = excelPackage.Workbook.Worksheets.Add(batteryNum);
                        worksheet.Cells["A1"].Value = title;
                        worksheet.Cells["A2"].Value = $"Starting list {category}";
                        if (isNational)
                        {

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
                                worksheet.Cells[$"D{4 + j}"].Value = getSurname(atleta["Nazione"], isTeam);
                                worksheet.Cells[$"E{4 + j}"].Value = isTeam ? "" : atleta["Atleta1"].Replace("|", " ");
                                worksheet.Cells[$"F{4 + j}"].Value = getTeamName(atleta["Nazione"]);
                                worksheet.Cells[$"G{4 + j}"].Value = atleta["Atleta1"].Replace("|", " ");
                                worksheet.Cells[$"H{4 + j}"].Value = atleta["Atleta2"].Replace("|", " ");
                                worksheet.Cells[$"I{4 + j}"].Value = atleta["Atleta3"].Replace("|", " ");
                                worksheet.Cells[$"J{4 + j}"].Value = atleta["Atleta4"].Replace("|", " ");
                                worksheet.Cells[$"K{4 + j}"].Value = atleta["Atleta5"].Replace("|", " ");
                                worksheet.Cells[$"L{4 + j}"].Value = atleta["Atleta6"].Replace("|", " ");
                                worksheet.Cells[$"M{4 + j}"].Value = atleta["Atleta7"].Replace("|", " ");
                                worksheet.Cells[$"N{4 + j}"].Value = atleta["Atleta8"].Replace("|", " ");
                                worksheet.Cells[$"O{4 + j}"].Value = atleta["Atleta9"].Replace("|", " ") + " (COX)";
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
        private static void credits()
        {
            Console.WriteLine("Sviluppato da Giuseppe Elefante <giuseppe.elefante90@gmail.com>\nFICr Salerno - A.S.D. Cronometristi Salernitani \"R. Marra\"\n");
        }
        private static string getDesktopPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }
    }
    
}
