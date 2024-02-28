using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Net;
using System.Globalization;
using System.IO;
using Search_EngineTLS;
using GemBox.Spreadsheet;
using System.Security.Cryptography.X509Certificates;
using Microsoft.AspNetCore.Http;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
//using Microsoft.AspNetCore.Html;


class Trimming
{
    public static void ExtractContent(string webpage)
    {
        string[] extractedContent = ExtractParagraphContent(webpage);

        // Print the extracted content
        foreach (string content in extractedContent)
        {
            Console.WriteLine(content);
        }
    }

    static string[] ExtractParagraphContent(string inputString)
    {
        // Define a regular expression pattern to match content between <p> and </p> tags
        string pattern = @"<p>([^<>]*)</p>";

        // Use Regex.Matches to get all matches in the input string
        MatchCollection matches = Regex.Matches(inputString, pattern, RegexOptions.Singleline);

        // Create an array to store the extracted content
        string[] extractedContent = new string[matches.Count];

        // Populate the array with the extracted content
        for (int i = 0; i < matches.Count; i++)
        {
            extractedContent[i] = matches[i].Groups[1].Value;
        }

        // Return the extracted content
        return extractedContent;

    }
}
class Database
{
    public static void Filling(string URL, string Webpage, int LinkNumber)
    {
        SpreadsheetInfo.SetLicense("FREE -LIMITED-KEY");

        // Create new empty workbook.
        var workbook = new ExcelFile();

        // Add new sheet.
        var worksheet = workbook.Worksheets.Add("SearchEngineTable");

        // Write title to Excel cell.
        worksheet.Cells["A1"].Value = "SearchEngineTable";

        // Tabular sample data for writing into an Excel file.
        var Indexables = new object[,]
        {
             {1, URL, Webpage, LinkNumber},
        };
        // Set row formatting.
        worksheet.Rows["1"].Style = workbook.Styles[BuiltInCellStyleName.Heading1];

        // Write header data to Excel cells.
        for (int col = 0; col < Indexables.GetLength(1); col++)
            worksheet.Cells[2, col].Value = Indexables[0, col];//change this to update places

        // Set header cells formatting.
        var style = new CellStyle();
        style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
        style.VerticalAlignment = VerticalAlignmentStyle.Center;
        style.FillPattern.SetSolid(SpreadsheetColor.FromArgb(255, 255, 255));
        style.Font.Weight = ExcelFont.BoldWeight;
        style.Font.Color = SpreadsheetColor.FromName(ColorName.Black);
        style.WrapText = false;
        style.Borders.SetBorders(MultipleBorders.Right | MultipleBorders.Top, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
        worksheet.Cells.GetSubrange("A3:H4").Style = style;


        // Save workbook as an Excel file.
        workbook.Save("IndexTable.xlsx");
        Console.WriteLine("IT WORKED");
    }
}


