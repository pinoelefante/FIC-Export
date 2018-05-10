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

namespace CanottaggioGui
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _pathCsv = string.Empty, _tvgFolder = @"C:\tvg\Canottaggio_Int", _title, _separator, _exportType, _exportTypeNation, textArea;
        
        public string PathCSV { get => _pathCsv; set => Set(ref _pathCsv, value); }
        public string TVGFolder { get => _tvgFolder; set => Set(ref _tvgFolder, value); }
        public string Title { get => _title; set => Set(ref _title, value); }
        public string Separator { get => _separator; set => Set(ref _separator, value); }
        public string ExportType { get => _exportType; set => Set(ref _exportType, value); }
        public string ExportTypeNation { get => _exportTypeNation; set => Set(ref _exportTypeNation, value); }
        public string TextArea { get => textArea; set => Set(ref textArea, value); }
        
        private RelayCommand _startCmd, _selectCSVCmd, _selectTVGCmd;

        public RelayCommand StartAction =>
            _startCmd ??
            (_startCmd = new RelayCommand(() =>
            {
                if(string.IsNullOrEmpty(Title))
                {
                    MessageBox.Show("Inserire un titolo");
                    return;
                }
                if(!File.Exists(PathCSV))
                {
                    MessageBox.Show("Scegliere un file CSV");
                    return;
                }
                if (!Directory.Exists(TVGFolder))
                {
                    MessageBox.Show("Scegliere una cartella valida di TVG");
                    return;
                }
                if (!File.Exists("CanottaggioConsole.exe"))
                {
                    MessageBox.Show("Non è possibile avviare il programma.\nManca: CanottaggioConsole.exe");
                    return;
                }
                var process = Process.Start("CanottaggioConsole.exe", $"\"{PathCSV}\" {Separator} {ExportType} {(ExportTypeNation.Equals("Nazionale") ? true : false).ToString()} \"{Title}\" \"{TVGFolder}\"");
                process.OutputDataReceived += Process_OutputDataReceived;
            }));

        private void Process_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            TextArea += e.Data;
        }

        public RelayCommand SelectCSVFileCommand =>
            _selectCSVCmd ??
            (_selectCSVCmd = new RelayCommand(() =>
            {
                var dialog = new OpenFileDialog();
                dialog.Filter = "csv (*.csv)|*.csv";
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
                    TVGFolder = dialog.FileName.Substring(0, dialog.FileName.LastIndexOf("ficr"));
            }));

        private void Set<K>(ref K k, K value, [CallerMemberName]string fieldName = "")
        {
            k = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(fieldName));
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
