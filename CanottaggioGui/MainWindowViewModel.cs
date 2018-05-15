using CanottaggioGui.Data;
using CanottaggioGui.DataConverters;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace CanottaggioGui
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private AppDataLoader dataLoader;
        private MiSpeakerConverter mispeaker;
        private TVGConverter tvg;
        public MainWindowViewModel()
        {
            Task.Factory.StartNew(() =>
            {
                dataLoader = new AppDataLoader();
                var categories = dataLoader.LoadCategoriesFile();
                var nations = dataLoader.LoadNationsFile();
                var teams = dataLoader.LoadTeamsFile();
                var fileConfig = dataLoader.LoadFileMappingConfig();
                mispeaker = new MiSpeakerConverter()
                {
                    Categories = categories,
                    Nations = nations,
                    Teams = teams
                };
                tvg = new TVGConverter()
                {
                    Categories = categories,
                    Nations = nations,
                    Teams = teams,
                    BaseFolder = TVGFolder
                };
                LoadFileConfig();
                IsProgramLoaded = true;
            });
        }
        public event PropertyChangedEventHandler PropertyChanged;

        private string _pathCsv = string.Empty, _tvgFolder = @"C:\tvg\Canottaggio_Int", _title, _exportType, _exportTypeNation, textArea;
        private bool _loaded = false;
        
        public string PathCSV { get => _pathCsv; set => Set(ref _pathCsv, value); }
        public string TVGFolder { get => _tvgFolder; set => Set(ref _tvgFolder, value); }
        public string Title { get => _title; set => Set(ref _title, value); }
        public string ExportType { get => _exportType; set => Set(ref _exportType, value); }
        public string ExportTypeNation { get => _exportTypeNation; set => Set(ref _exportTypeNation, value); }
        public string TextArea { get => textArea; set => Set(ref textArea, value); }

        public bool IsProgramLoaded { get => _loaded; set => Set(ref _loaded, value); }
        
        private RelayCommand _startCmd, _selectCSVCmd, _selectTVGCmd;

        public RelayCommand StartAction =>
            _startCmd ??
            (_startCmd = new RelayCommand(() =>
            {
                TextArea = string.Empty;
                var startTime = DateTime.Now;
                if (string.IsNullOrEmpty(Title))
                {
                    MessageBox.Show("Inserire un titolo");
                    return;
                }
                if(!File.Exists(PathCSV))
                {
                    MessageBox.Show("Scegliere un file Excel/CSV");
                    return;
                }
                if (!Directory.Exists(TVGFolder))
                {
                    MessageBox.Show("Scegliere una cartella valida di TVG");
                    return;
                }
                TextArea += $"---- INIZIO ESECUZIONE ({startTime.Hour.ToString("D2")}:{startTime.Minute.ToString("D2")}:{startTime.Second.ToString("D2")})----\n";
                var isNazionale = ExportTypeNation.Equals("Nazionale") ? true : false;
                var file_content = tvg.parseFileExcel(PathCSV);
                switch (ExportType)
                {
                    case "mispeaker":
                        mispeaker.Convert(file_content, isNazionale);
                        break;
                    case "tvg":
                        tvg.Convert(file_content, isNazionale, Title);
                        break;
                    case "atleti":
                        tvg.AthletesCreditsConvert(file_content);
                        break;
                    case "tutto":
                        mispeaker.Convert(file_content, isNazionale);
                        tvg.Convert(file_content, isNazionale, Title);
                        tvg.AthletesCreditsConvert(file_content);
                        break;
                }
                if (!isNazionale)
                    tvg.VerifyFlagsInt(file_content);

                var mispeakerText = mispeaker.GetTextFromStream();
                var tvgText = tvg.GetTextFromStream();
                var endTime = DateTime.Now;
                var diff = endTime.Subtract(startTime);
                TextArea += mispeakerText;
                TextArea += tvgText;
                TextArea += $"---- FINE ESECUZIONE (Durata: {diff.TotalSeconds} secondi) ----\n\n";
            }));

        public RelayCommand SelectCSVFileCommand =>
            _selectCSVCmd ??
            (_selectCSVCmd = new RelayCommand(() =>
            {
                var dialog = new OpenFileDialog();
                dialog.Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv";
                if(dialog.ShowDialog() == true)
                    PathCSV = dialog.FileName;
            }));
        public RelayCommand SelectTVGPathCommand =>
            _selectTVGCmd ??
            (_selectTVGCmd = new RelayCommand(() =>
            {
                var dialog = new OpenFileDialog();
                dialog.Filter = string.Empty;
                dialog.FileName = "ficr";
                dialog.CheckFileExists = false;
                dialog.CheckPathExists = false;
                dialog.InitialDirectory = @"C:\tvg";
                if (dialog.ShowDialog() == true)
                {
                    TVGFolder = dialog.FileName.Substring(0, dialog.FileName.LastIndexOf("ficr"));
                    tvg.BaseFolder = TVGFolder;
                }
            }));

        private void Set<K>(ref K k, K value, [CallerMemberName]string fieldName = "")
        {
            k = value;
            Dispatcher.CurrentDispatcher.Invoke(() =>
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(fieldName));
            });
        }
        public void LoadFileConfig()
        {
            var config = dataLoader.LoadFileMappingConfig();
            if(config != null)
            {
                mispeaker.FileConfig = config;
                tvg.FileConfig = config;
            }
        }
    }
    public class RelayCommand : ICommand
    {
        private Action action;
        public RelayCommand(Action a)
        {
            action = a;
        }
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            action.Invoke();
        }
    }

}
