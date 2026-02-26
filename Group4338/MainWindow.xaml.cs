using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;

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

        private void BtnImportJson_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "JSON files|*.json",
                    FileName = "1.json",
                    Title = "Выберите файл 1.json"
                };

                if (openFileDialog.ShowDialog() != true)
                    return;

                string filePath = openFileDialog.FileName;
                var services = ImportFromJson(filePath);
                SaveToDatabase(services);

                MessageBox.Show($"✓ Успешно импортировано {services.Count} записей из JSON!",
                    "Импорт JSON завершён", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"✗ Ошибка импорта JSON: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExportWord_Click(object sender, RoutedEventArgs e)
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
                    Filter = "Word documents|*.docx",
                    FileName = "Services_Group4338.docx",
                    Title = "Сохранить результат экспорта в Word"
                };

                if (saveFileDialog.ShowDialog() != true)
                    return;

                ExportGroupedToWord(services, saveFileDialog.FileName);

                MessageBox.Show($"✓ Экспорт в Word завершён!\nФайл: {saveFileDialog.FileName}",
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"✗ Ошибка экспорта в Word: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<Service> ImportFromJson(string filePath)
        {
            var jsonContent = File.ReadAllText(filePath);

            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            var jsonServices = JsonSerializer.Deserialize<List<JsonService>>(jsonContent, options);

            var services = new List<Service>();
            foreach (var js in jsonServices)
            {
                services.Add(new Service
                {
                    Id = js.IdServices,
                    Name = js.NameServices,
                    Type = js.TypeOfService,
                    Code = js.CodeService,
                    CostPerHour = js.Cost
                });
            }
            return services;
        }

        private class JsonService
        {
            public int IdServices { get; set; }
            public string NameServices { get; set; }
            public string TypeOfService { get; set; }
            public string CodeService { get; set; }
            public decimal Cost { get; set; }
        }

        private void ExportGroupedToWord(List<Service> services, string outputPath)
        {
            var grouped = services
                .GroupBy(s => s.Type)
                .OrderBy(g => g.Key)
                .ToDictionary(g => g.Key, g => g.OrderBy(s => s.CostPerHour).ToList());

            using (WordprocessingDocument doc = WordprocessingDocument
                .Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                int categoryIndex = 0;
                foreach (var category in grouped)
                {
                    var titleParagraph = body.AppendChild(new Paragraph());
                    var titleRun = titleParagraph.AppendChild(new Run());
                    titleRun.AppendChild(new Text(category.Key));
                    titleRun.RunProperties = new RunProperties(new Bold());
                    body.AppendChild(new Paragraph()); 

                    var table = body.AppendChild(new Table());

                    var headerRow = new TableRow(
                        CreateTableCell("Id", true),
                        CreateTableCell("Название услуги", true),
                        CreateTableCell("Стоимость", true)
                    );
                    table.Append(headerRow);

                    foreach (var service in category.Value)
                    {
                        var row = new TableRow(
                            CreateTableCell(service.Id.ToString()),
                            CreateTableCell(service.Name),
                            CreateTableCell($"{service.CostPerHour} руб.")
                        );
                        table.Append(row);
                    }

                    categoryIndex++;
                    if (categoryIndex < grouped.Count)
                    {
                        body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    }
                }

                mainPart.Document.Save();
            }
        }

        private TableCell CreateTableCell(string text, bool isHeader = false)
        {
            var cell = new TableCell();
            var paragraph = new Paragraph();
            var run = new Run(new Text(text));

            if (isHeader)
            {
                run.RunProperties = new RunProperties(new Bold());
            }

            paragraph.Append(run);
            cell.Append(paragraph);
            return cell;
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
                    cmd.CommandText = "INSERT OR REPLACE INTO Services (Id, Name, Type, Code, CostPerHour) VALUES (@Id, @Name, @Type, @Code, @Cost)";
                    cmd.Parameters.AddWithValue("@Id", s.Id);
                    cmd.Parameters.AddWithValue("@Name", s.Name);
                    cmd.Parameters.AddWithValue("@Type", s.Type);
                    cmd.Parameters.AddWithValue("@Code", s.Code);
                    cmd.Parameters.AddWithValue("@Cost", s.CostPerHour);
                    cmd.ExecuteNonQuery();
                }
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