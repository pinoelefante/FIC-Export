using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CanottaggioGui.DataConverters
{
    public class TVGConverter : ConverterBasics
    {
        public string BaseFolder { get; set; } = @"C:\tvg";
        private static readonly char[] splitName = new char[] { '|' };

        public override bool ConvertInternational(List<Dictionary<string, string>> fields, string title = "")
        {
            OutputStream.AppendLine("\nAvvio esportazione TVG Internazionale");
            try
            {
                var groups = fields.GroupBy(x => x["Batteria"]);
                using (var excelPackage = new ExcelPackage())
                {
                    for (int i = groups.Count() - 1; i >= 0; i--)
                    {
                        var group = groups.ElementAt(i);
                        var batteryNum = group.First()["Batteria"];
                        if (string.IsNullOrEmpty(batteryNum))
                        {
                            OutputStream.AppendLine("Verifica che il file excel non contenga righe vuote");
                            continue;
                        }
                        var category = group.First()["Categoria"];
                        var firstAtleta = group.First();
                        var isTeam = firstAtleta.ContainsKey("Atleta3") && !string.IsNullOrEmpty(firstAtleta["Atleta3"]?.Trim());
                        var worksheet = excelPackage.Workbook.Worksheets.Add(batteryNum);
                        worksheet.Cells["A1"].Value = title;
                        worksheet.Cells["A2"].Value = $"Starting list {category}";
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
                            worksheet.Cells[$"C{4 + j}"].Value = $@"Flags3D\{GetFlagName(atleta["Nazione"])}.png";
                            worksheet.Cells[$"D{4 + j}"].Value = GetSurnameInt(atleta["Nazione"], isTeam);
                            worksheet.Cells[$"E{4 + j}"].Value = isTeam ? "" : atleta["Atleta1"].Replace("|", " ");
                            worksheet.Cells[$"F{4 + j}"].Value = GetTeamNameInt(atleta["Nazione"]);
                            worksheet.Cells[$"G{4 + j}"].Value = atleta.ContainsKey("Atleta1") && !string.IsNullOrEmpty(atleta["Atleta1"]) ? atleta["Atleta1"].Replace("|", " ") : "";
                            worksheet.Cells[$"H{4 + j}"].Value = atleta.ContainsKey("Atleta2") && !string.IsNullOrEmpty(atleta["Atleta2"]) ? atleta["Atleta2"].Replace("|", " ") : "";
                            worksheet.Cells[$"I{4 + j}"].Value = atleta.ContainsKey("Atleta3") && !string.IsNullOrEmpty(atleta["Atleta3"]) ? atleta["Atleta3"].Replace("|", " ") : "";
                            worksheet.Cells[$"J{4 + j}"].Value = atleta.ContainsKey("Atleta4") && !string.IsNullOrEmpty(atleta["Atleta4"]) ? atleta["Atleta4"].Replace("|", " ") : "";
                            worksheet.Cells[$"K{4 + j}"].Value = atleta.ContainsKey("Atleta5") && !string.IsNullOrEmpty(atleta["Atleta5"]) ? atleta["Atleta5"].Replace("|", " ") : "";
                            worksheet.Cells[$"L{4 + j}"].Value = atleta.ContainsKey("Atleta6") && !string.IsNullOrEmpty(atleta["Atleta6"]) ? atleta["Atleta6"].Replace("|", " ") : "";
                            worksheet.Cells[$"M{4 + j}"].Value = atleta.ContainsKey("Atleta7") && !string.IsNullOrEmpty(atleta["Atleta7"]) ? atleta["Atleta7"].Replace("|", " ") : "";
                            worksheet.Cells[$"N{4 + j}"].Value = atleta.ContainsKey("Atleta8") && !string.IsNullOrEmpty(atleta["Atleta8"]) ? atleta["Atleta8"].Replace("|", " ") : "";
                            worksheet.Cells[$"O{4 + j}"].Value = atleta.ContainsKey("Atleta9") && !string.IsNullOrEmpty(atleta["Atleta9"]) ? atleta["Atleta9"].Replace("|", " ") + " (COX)" : "";
                            worksheet.Cells[$"P{4 + j}"].Value = atleta["Categoria"];
                            worksheet.Cells[$"Q{4 + j}"].Value = GetCategoryDescription(atleta["Categoria"], atleta["Categoria2"]);
                            worksheet.Cells[$"R{4 + j}"].Value = Int32.Parse(atleta["Batteria"]);
                            worksheet.Cells[$"S{4 + j}"].Value = Int32.Parse(atleta["Acqua"]);
                            worksheet.Cells[$"T{4 + j}"].Value = 1;
                        }
                    }

                    var dateNow = DateTime.Now;
                    var filename = string.Format("Export_TVG_Int_{0:D2}{1:D2}{2:D4}.xlsx", dateNow.Day, dateNow.Month, dateNow.Year);
                    var fileinfo = new FileInfo($@"{GetDesktopPath()}\{filename}");
                    excelPackage.SaveAs(fileinfo);
                    OutputStream.AppendLine($"Il file {filename} e' stato salvato sul desktop");
                    OutputStream.AppendLine("Esportazione TVG completata con successo");
                    return true;
                }
            }
            catch (Exception e)
            {
                OutputStream.AppendLine($"Si e' verificato un errore durante l'esportazione di TVG\n{e.Message}\n{e.StackTrace}");
                return false;
            }
        }

        public override bool ConvertNational(List<Dictionary<string, string>> fields, string title = "")
        {
            OutputStream.AppendLine("\nAvvio esportazione TVG Nazionale");
            try
            {
                var groups = fields.GroupBy(x => x["Batteria"]);
                using (var excelPackage = new ExcelPackage())
                {
                    for (int i = groups.Count() - 1; i >= 0; i--)
                    {
                        var group = groups.ElementAt(i);
                        var batteryNum = group.First()["Batteria"];
                        if (string.IsNullOrEmpty(batteryNum))
                        {
                            OutputStream.AppendLine("Verifica che il file excel non contenga righe vuote");
                            continue;
                        }
                        var category = group.First()["Categoria"];
                        var firstAtleta = group.First();
                        var isTeam = firstAtleta.ContainsKey("Atleta3") && !string.IsNullOrEmpty(firstAtleta["Atleta3"]?.Trim());
                        var worksheet = excelPackage.Workbook.Worksheets.Add(batteryNum);
                        worksheet.Cells["A1"].Value = title;
                        worksheet.Cells["A2"].Value = $"Starting list {category}";
                        OutputStream.AppendLine("ATTENZIONE: L'export nazionale potrebbe non essere corretto in quanto non testato");
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
                        var sex = GetCategorySex(group.First()["Categoria2"], group.First()["Categoria"]);
                        for (int j = 0; j < group.Count(); j++)
                        {
                            var atleta = group.ElementAt(j);
                            worksheet.Cells[$"A{4 + j}"].Value = Int32.Parse(atleta["Acqua"]);
                            worksheet.Cells[$"B{4 + j}"].Value = Int32.Parse(atleta["Pettorale"]);
                            worksheet.Cells[$"C{4 + j}"].Value = GetSurnameInt(atleta["Nazione"], isTeam);
                            worksheet.Cells[$"D{4 + j}"].Value = isTeam ? "" : atleta["Atleta1"]?.Replace("|", " ");
                            worksheet.Cells[$"E{4 + j}"].Value = GetTeamNameNational(atleta["id_squadra"]);
                            worksheet.Cells[$"F{4 + j}"].Value = atleta["Atleta1"]?.Replace("|", " ");
                            worksheet.Cells[$"G{4 + j}"].Value = atleta["Atleta2"]?.Replace("|", " ");
                            worksheet.Cells[$"H{4 + j}"].Value = atleta["Atleta3"]?.Replace("|", " ");
                            worksheet.Cells[$"I{4 + j}"].Value = atleta["Atleta4"]?.Replace("|", " ");
                            worksheet.Cells[$"J{4 + j}"].Value = atleta["Atleta5"]?.Replace("|", " ");
                            worksheet.Cells[$"K{4 + j}"].Value = atleta["Atleta6"]?.Replace("|", " ");
                            worksheet.Cells[$"L{4 + j}"].Value = atleta["Atleta7"]?.Replace("|", " ");
                            worksheet.Cells[$"M{4 + j}"].Value = atleta["Atleta8"]?.Replace("|", " ");
                            worksheet.Cells[$"N{4 + j}"].Value = atleta["Atleta9"]?.Replace("|", " ") + " (COX)";
                            worksheet.Cells[$"O{4 + j}"].Value = GetCategoryDescription(atleta["Categoria2"], atleta["Categoria"]);
                            worksheet.Cells[$"P{4 + j}"].Value = ""; //descr_turno
                            worksheet.Cells[$"Q{4 + j}"].Value = atleta["Categoria2"];
                            worksheet.Cells[$"R{4 + j}"].Value = Int32.Parse(atleta["Batteria"]);
                            worksheet.Cells[$"S{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta1"]) ? atleta["Atleta1"].Split(splitName)[0] : "";
                            worksheet.Cells[$"T{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta1"]) ? atleta["Atleta1"].Split(splitName)[1] : "";
                            worksheet.Cells[$"U{4 + j}"].Value = sex;
                            worksheet.Cells[$"V{4 + j}"].Value = Int32.Parse(atleta["Batteria"]);
                            worksheet.Cells[$"W{4 + j}"].Value = atleta["id_squadra"];
                            worksheet.Cells[$"X{4 + j}"].Value = Int32.Parse(atleta["Acqua"]);
                            worksheet.Cells[$"Y{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta1"]) ? atleta["Atleta1"].Split(splitName)[0] : "";
                            worksheet.Cells[$"Z{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta2"]) ? atleta["Atleta2"].Split(splitName)[0] : "";
                            worksheet.Cells[$"AA{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta3"]) ? atleta["Atleta3"].Split(splitName)[0] : "";
                            worksheet.Cells[$"AB{4 + j}"].Value = !string.IsNullOrEmpty(atleta["Atleta4"]) ? atleta["Atleta4"].Split(splitName)[0] : "";
                            worksheet.Cells[$"AC{4 + j}"].Value = ""; //equipaggio
                        }

                    }
                    var dateNow = DateTime.Now;
                    var filename = string.Format("Export_TVG_Naz_{0:D2}{1:D2}{2:D4}.xlsx", dateNow.Day, dateNow.Month, dateNow.Year);
                    var fileinfo = new FileInfo($@"{GetDesktopPath()}\{filename}");
                    excelPackage.SaveAs(fileinfo);
                    OutputStream.AppendLine($"Il file {filename} e' stato salvato sul desktop");
                    OutputStream.AppendLine("Esportazione TVG completata con successo");
                    return true;
                }
            }
            catch (Exception e)
            {
                OutputStream.AppendLine($"Si e' verificato un errore durante l'esportazione di TVG\n{e.Message}\n{e.StackTrace}");
                return false;
            }
        }
        public void AthletesCreditsConvert(List<Dictionary<string, string>> fields)
        {
            OutputStream.AppendLine("\nCarico la lista degli atleti");
            var list = new List<string>();
            foreach (var dict in fields)
            {
                for (int i = 1; i <= 9 && dict.ContainsKey($"Atleta{i}"); i++)
                {

                    if (!string.IsNullOrEmpty(dict[$"Atleta{i}"]))
                        list.Add(dict[$"Atleta{i}"].Trim().Replace('|', ' '));
                    else
                        break;
                }
            }
            list = list.Distinct().OrderBy(x => x).ToList();
            var now = DateTime.Now;
            var filename = string.Format("Export_Credits_{0:D2}{1:D2}{2:D4}.xlsx", now.Day, now.Month, now.Year);

            using (var excel = new ExcelPackage())

            {
                var sheet = excel.Workbook.Worksheets.Add("Atleti");
                for (int i = 0; i < list.Count; i++)
                    sheet.Cells[$"A{(i + 1)}"].Value = list[i];
                OutputStream.AppendLine($"Salvo la lista degli atleti sul desktop nel file {filename}");
                var fileinfo = new FileInfo($@"{GetDesktopPath()}\{filename}");
                excel.SaveAs(fileinfo);
            }
        }
        public void VerifyFlagsInt(List<Dictionary<string, string>> contents)
        {
            OutputStream.AppendLine("\nVerifico la presenza dei file delle bandiere");
            Dictionary<string, bool> verifyStatus = new Dictionary<string, bool>();
            foreach (var atl in contents)
            {
                if (atl.ContainsKey("Nazione"))
                {
                    var flag = $@"Flags3D\{GetFlagName(atl["Nazione"])}.png";
                    if (!verifyStatus.ContainsKey(flag))
                    {
                        var path = Path.Combine(BaseFolder, flag);
                        var flagExists = File.Exists(path);
                        verifyStatus.Add(flag, flagExists);
                        if (!flagExists)
                            OutputStream.AppendLine($"Bandiera non presente ({flag})");
                    }
                }
            }
        }
    }
}
