/*
 * Applica, Inc.
 * By José O. Lara
 * 2023.03.09
 */

namespace wpfApp_Demo.src.core.viewModels
{
    using System;
    using System.Threading.Tasks;

    using CommunityToolkit.Mvvm.ComponentModel;
    using CommunityToolkit.Mvvm.Input;

    using wpfApp_Demo.src.ui.views;
    using System.Windows;
    using System.Collections.ObjectModel;
    using Microsoft.Win32;
    using wpfApp_Demo.src.core.interfaces.services;
    using wpfApp_Demo.src.core.services;
    using System.Data;
    using System.Data.OleDb;
    using OfficeOpenXml;
    using System.IO;

    public partial class MainVm : ObservableObject {

        #region Property
        private int _TotalImagenes = 0;
        private int _TotalReclamaciones = 0;
        private IToolsService? _ToolsService;


        private IOpenDbfService? _OpenDbfService;

        public RelayCommand<MainWpf> BtnClose { get; }
        public RelayCommand BtnVdeFile { get; }
        public RelayCommand BtnFileOut { get; }
        public RelayCommand BtnSave { get; }

        [ObservableProperty]
        private string? _TxtVdeFile;

        [ObservableProperty]
        private string _TxtFileOut = "";

        [ObservableProperty]
        ObservableCollection<string> _ListLog = new();

        [ObservableProperty]
        public ObservableCollection<int> _LibLogSelectedIndex = new() { 0 };

        [ObservableProperty]
        public string _TxtTitel = "";

        #endregion

        #region Constructor
        public MainVm() {
            BtnClose = new RelayCommand<MainWpf>(async param => await ClickBtnClose(param!));
            BtnVdeFile = new RelayCommand(async () => await ClickBtnVdeFile());
            BtnFileOut = new RelayCommand(async () => await ClickBtnFileOut());
            BtnSave = new RelayCommand(async () => await ClickBtnSave());

            TxtTitel =
            $@"Applica Demo 2023 
Version: {GetType().Assembly.GetName().Version}";


        }

        #endregion

        #region Methods or functions

        private static Task ClickBtnClose(MainWpf mainWpf) {

            try {
                mainWpf?.Close();
            } catch (Exception err) {
                MessageBox.Show(err.ToString(), "ClickBtnClose", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Windows.Application.Current.Shutdown();
            }

            return Task.CompletedTask;
        }

        private async Task ClickBtnVdeFile() {
            try {
                TimeSpan timeSpan = new(DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                ListLog.Add($"Start GetData {timeSpan.ToString()}");

                OpenFileDialog op = new() {
                    DefaultExt = ".dbf",
                    Filter = "Vde Files|*.dbf|All Files|*.*"
                };

                bool? result = op.ShowDialog();
                if (result != null && result.Value) {
                    TxtVdeFile = op.FileName;
                    _OpenDbfService = new OpenDbfService(new(TxtVdeFile));
                    DataTable T = await _OpenDbfService.GetAllAsDataTableAsync();

                    _TotalImagenes = T.Rows.Count;
                    _TotalReclamaciones = T.Select("v1page <> '99'").Length;

                    ListLog.Add($"Vde File: {TxtVdeFile}");
                    ListLog.Add($"Total de Record: {_TotalImagenes:N0}");
                    ListLog.Add($"Total de Reclamaciones: {_TotalReclamaciones:N0}");

                }

            } catch (Exception err) {
                MessageBox.Show(err.ToString(), "ClickBtnVdeFile", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            //return Task.CompletedTask;
        }

        public Task ClickBtnFileOut() {
            try {

                SaveFileDialog fb = new SaveFileDialog();
                fb.Filter = "Excel Files|*.xlsx";
                fb.FileName = "Totales.xlsx";
                bool? result = fb.ShowDialog();

                if (result != null && result.Value) {
                    TxtFileOut = fb.FileName;
                    ListLog.Add(TxtFileOut);
                }

            } catch (Exception err) {
                MessageBox.Show(err.ToString(), "ClickBtnFileOut", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return Task.CompletedTask;
        }

        private Task ClickBtnSave() {
            try {

                // TODO to Implemented
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=assets/vdeFiles/SSS1503.dbf;";

                /*
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    OleDBDataReader imagenesReader = new OleDbCommand("SELECT COUNT(v0document) FROM SSS1503.dbf", connection).ExecuteReader();
                    OleDBDataReader reclamacionesReader = new OleDbCommand("SELECT COUNT(v1page) FROM SSS1503.dbf WHERE v1page < 99", connection).ExecuteReader();
                    ExcelPackage package = new ExcelPackage(new FileInfo("Totales.xlsx"));
                    package.Workbook.Worksheets.Add("Resultados").Cells[1, 1].Value = "Date";
                    package.Workbook.Worksheets.Add("Resultados").Cells[1, 2].Value = DateTime.Now.ToString("dd/MM/yyyy");
                    package.Workbook.Worksheets.Add("Resultados").Cells[2, 1].Value = "Total Imagenes";
                    package.Workbook.Worksheets.Add("Resultados").Cells[2, 2].Value = imagenesReader.GetInt32(0);
                    package.Workbook.Worksheets.Add("Resultados").Cells[3, 1].Value = "Total Reclamaciones";
                    package.Workbook.Worksheets.Add("Resultados").Cells[3, 2].Value = reclamacionesReader.GetInt32(0);

                    package.Save();
                    connection.Close();
                }*/

                string[] info = new[] { "Numero totales", Convert.ToString(_TotalImagenes), " ", "Reclamaciones", Convert.ToString(_TotalReclamaciones) };

                _ToolsService = new ToolsService();

                System.Data.DataTable outTable = new();
                object?[] myArray = new object[5];

                _ = outTable.Columns.Add("BatchNum", typeof(string));
                _ = outTable.Columns.Add("Total Imagenes", typeof(int));
                _ = outTable.Columns.Add("Total Reclamaciones", typeof(int));
                _ = outTable.Columns.Add("Job", typeof(string));
                _ = outTable.Columns.Add("Date", typeof(string));

                myArray[0] = "2023";
                myArray[1] = _TotalImagenes.ToString();
                myArray[2] = _TotalReclamaciones.ToString();
                myArray[3] = "";
                myArray[4] = DateTime.Now.ToString("MM/dd/yyyy");
                outTable.Rows.Add(myArray);

                _ToolsService.TotalesToExcel(TxtFileOut, "Demo", "Demo 2023", DateTime.Now.ToString("MM/dd/yyyy"), outTable);

                MessageBox.Show("Saved", "saved", MessageBoxButton.OK, MessageBoxImage.Exclamation);


            } catch (Exception err) {
                MessageBox.Show(err.ToString(), "ClickBtnSave", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return Task.CompletedTask;
        }

        #endregion



    }
}
