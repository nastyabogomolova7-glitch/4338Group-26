using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using Microsoft.Win32;

namespace Group4338
{
    public partial class MainWindow : Window
    {
        private string dbPath = "Data Source=services.db";

        public MainWindow()
        {
            InitializeComponent();
            CreateDatabaseTable();
        }

        private void AuthorButton_Click(object sender, RoutedEventArgs e)
        {
            var authorWindow = new _4338_Bogomolova();
            authorWindow.ShowDialog();
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel files|*.xlsx",
                    Title = "Выберите файл для импорта"
                };

                if (openFileDialog.ShowDialog() != true)
                    return;

                string filePath = openFileDialog.FileName;
                var services = ImportFromExcel(filePath);
                SaveToDatabase(services);

                MessageBox.Show("Успешно импортировано " + services.Count + " записей!",
                    "Импорт завершён", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка импорта: " + ex.Message,
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<Service> ImportFromExcel(string filePath)
        {
            var services = new List<Service>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var range = worksheet.RangeUsed();

                for (int row = 2; row <= range.RowCount(); row++)
                {
                    services.Add(new Service
                    {
                        Id = int.Parse(worksheet.Cell(row, 1).Value.ToString()),
                        Name = worksheet.Cell(row, 2).Value.ToString(),
                        Type = worksheet.Cell(row, 3).Value.ToString(),
                        Code = worksheet.Cell(row, 4).Value.ToString(),
                        CostPerHour = decimal.Parse(worksheet.Cell(row, 5).Value.ToString())
                    });
                }
            }
            return services;
        }

        private void CreateDatabaseTable()
        {
            using (var conn = new SqliteConnection(dbPath))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                cmd.CommandText = "CREATE TABLE IF NOT EXISTS Services (Id INTEGER PRIMARY KEY, Name TEXT, Type TEXT, Code TEXT, CostPerHour REAL)";
                cmd.ExecuteNonQuery();
            }
        }

        private void SaveToDatabase(List<Service> services)
        {
            using (var conn = new SqliteConnection(dbPath))
            {
                conn.Open();

                var clearCmd = conn.CreateCommand();
                clearCmd.CommandText = "DELETE FROM Services";
                clearCmd.ExecuteNonQuery();

                foreach (var s in services)
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "INSERT INTO Services (Id, Name, Type, Code, CostPerHour) VALUES (@Id, @Name, @Type, @Code, @Cost)";
                    cmd.Parameters.AddWithValue("@Id", s.Id);
                    cmd.Parameters.AddWithValue("@Name", s.Name);
                    cmd.Parameters.AddWithValue("@Type", s.Type);
                    cmd.Parameters.AddWithValue("@Code", s.Code);
                    cmd.Parameters.AddWithValue("@Cost", s.CostPerHour);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var services = LoadAllServices();

                if (services.Count == 0)
                {
                    MessageBox.Show("Сначала импортируйте данные!",
                        "Нет данных", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files|*.xlsx",
                    FileName = "export_services.xlsx",
                    Title = "Сохранить результат экспорта"
                };

                if (saveFileDialog.ShowDialog() != true)
                    return;

                ExportGroupedToExcel(services, saveFileDialog.FileName);

                MessageBox.Show("Экспорт завершён!\nФайл: " + saveFileDialog.FileName,
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка экспорта: " + ex.Message,
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<Service> LoadAllServices()
        {
            var services = new List<Service>();

            using (var conn = new SqliteConnection(dbPath))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT Id, Name, Type, Code, CostPerHour FROM Services";

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        services.Add(new Service
                        {
                            Id = reader.GetInt32(0),
                            Name = reader.GetString(1),
                            Type = reader.GetString(2),
                            Code = reader.GetString(3),
                            CostPerHour = reader.GetDecimal(4)
                        });
                    }
                }
            }
            return services;
        }

        private void ExportGroupedToExcel(List<Service> services, string outputPath)
        {
            using (var workbook = new XLWorkbook())
            {
                var grouped = services.GroupBy(s => s.Type);

                foreach (var group in grouped)
                {
                    string sheetName = group.Key;
                    if (sheetName.Length > 31)
                        sheetName = sheetName.Substring(0, 31);

                    var worksheet = workbook.Worksheets.Add(sheetName);

                    worksheet.Cell(1, 1).Value = "Id";
                    worksheet.Cell(1, 2).Value = "Наименование услуги";
                    worksheet.Cell(1, 3).Value = "Стоимость";

                    worksheet.Range("A1:C1").Style.Font.Bold = true;

                    var sorted = group.OrderBy(s => s.CostPerHour).ToList();

                    int row = 2;
                    foreach (var item in sorted)
                    {
                        worksheet.Cell(row, 1).Value = item.Id;
                        worksheet.Cell(row, 2).Value = item.Name;
                        worksheet.Cell(row, 3).Value = item.CostPerHour;
                        row++;
                    }

                    worksheet.Columns().AdjustToContents();
                }

                workbook.SaveAs(outputPath);
            }
        }
    }

    public class Service
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public string Code { get; set; } = "";
        public decimal CostPerHour { get; set; }
    }
}