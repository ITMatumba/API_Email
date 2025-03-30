using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using Avalonia.Controls;
using Avalonia.Interactivity;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;

namespace AvaloniaApplication1;

public partial class MainWindow : Window
{
    private string DataFromApi = "";
    private const string FileName = "TestCase.docx";

    public MainWindow()
    {
        InitializeComponent();
    }

    private void SendTestResult_OnClick(object? sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(DataFromApi))
        {
            ValidationResultTBlock.Text = "Ошибка: нет данных для проверки!";
            return;
        }

        var pattern = @"^[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}$";
        var validationResult = Regex.IsMatch(DataFromApi, pattern);
        string validationMessage = validationResult 
            ? "ИНН содержит запрещенные символы" 
            : "ИНН не содержит запрещенные символы";

        ValidationResultTBlock.Text = validationMessage;

        try
        {
            using var doc = WordprocessingDocument.Open(FileName, true);
            var document = doc.MainDocumentPart.Document;

            bool isUpdated = ReplaceTextInTable("Result 1", validationResult, document) ||
                             ReplaceTextInTable("Result 2", validationResult, document);

            if (isUpdated)
            {
                document.Save();
            }
        }
        catch (Exception ex)
        {
            ValidationResultTBlock.Text = $"Ошибка при обработке файла: {ex.Message}";
        }
    }

    private bool ReplaceTextInTable(string replacedText, bool validationResult, Document document)
    {
        bool isTextReplaced = false;
        string replacementText = validationResult ? "Не успешно" : "Успешно";

        foreach (var table in document.Descendants<Table>())
        {
            foreach (var row in table.Descendants<TableRow>())
            {
                var cells = row.Descendants<TableCell>().ToList();

                if (cells.Count >= 3)
                {
                    var cell = cells[2];
                    var textElements = cell.Descendants<Text>().ToList();

                    if (textElements.Count > 0)
                    {
                        string fullText = string.Join("", textElements.Select(t => t.Text));

                        if (fullText.Contains(replacedText))
                        {
                            foreach (var text in textElements)
                            {
                                text.Text = string.Empty;
                            }

                            textElements[0].Text = fullText.Replace(replacedText, replacementText);
                            isTextReplaced = true;
                        }
                    }
                }
            }
        }

        return isTextReplaced;
    }

    private async void GetDataFromApi_OnClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            using var httpClient = new HttpClient();
            var content = await httpClient.GetStringAsync("http://localhost:4444/TransferSimulator/email");
            var data = JsonConvert.DeserializeObject<Dictionary<string, string>>(content);

            if (data != null && data.ContainsKey("value"))
            {
                DataFromApi = data["value"];
                DataFromApiTBlock.Text = DataFromApi;
            }
            else
            {
                DataFromApiTBlock.Text = "Ошибка: неверный формат ответа API!";
            }
        }
        catch (Exception ex)
        {
            DataFromApiTBlock.Text = $"Ошибка при получении данных: {ex.Message}";
        }
    }
}