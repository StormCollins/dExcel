namespace dExcel.Utilities;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using FuzzySharp;
using WPF;


/// <summary>
/// A collection of utility functions for creating hyperlinks between cells and sheets in Excel.
/// </summary>
public static class HyperLinkUtils
{
    [ExcelFunction(
        Name = "d.HyperlinkUtils_CreateHyperlinkToHeadingInSheet",
        Description = "Creates a hyperlink from the selected cell to a cell with the same content but styled as a heading.",
        Category = "∂Excel: Hyperlink Utils")]
    public static void CreateHyperlinkToHeadingInSheet()
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Range usedRange = xlApp.ActiveSheet.UsedRange;
        Dictionary<string, string> headings = new();
        List<char> invalidCharacters 
            = new()
            {
                ' ', ':', ',', ';', '(', ')', '[', ']', '{', '}', '<', '>', '*', '?', '!', '\'', '”', '-', '+', '|', '\\', '/'
            };

        List<string> generalExceptions = new();
        
        for (int i = 1; i < usedRange.Rows.Count + 1; i++)
        {
            for (int j = 1; j < usedRange.Columns.Count + 1; j++)
            {
                Excel.Range currentCell = usedRange.Cells[i, j];
                string style = ((Excel.Style)currentCell.Style).Value.ToUpper();
                if (style.Contains("HEADING"))
                {
                    try
                    {
                        string sheetName = currentCell.Worksheet.Name;
                        sheetName = sheetName.Replace(" ", "");
                        string cellContent = currentCell.Value2;
                        if (cellContent != null)
                        {
                            cellContent = string.Join("", cellContent.Where(x => !invalidCharacters.Contains(x)));
                            string headingCellName = $"{sheetName}.Title.{cellContent}";
                            currentCell.Name = headingCellName;
                            headings[currentCell.Value2] = ((Excel.Name)currentCell.Name).NameLocal;
                        }
                    }
                    catch (Exception e)
                    {
                        generalExceptions.Add(e.Message);
                    } 
                }
            }
        }

        Excel.Range selectedRange = (Excel.Range)xlApp.Selection;
        List<string> failedHyperlinks = new();

        for (int i = 1; i < selectedRange.Rows.Count + 1; i++)
        {
            string searchText = "";
            for (int j = 1; j < selectedRange.Columns.Count + 1; j++)
            {
                Excel.Range currentCell = selectedRange.Rows[i].Columns[j];
                if (currentCell.Value2 != null && currentCell.Value2?.ToString() != "" &&
                    currentCell.Value2?.ToString() != ExcelEmpty.Value.ToString())
                {
                    searchText = selectedRange.Rows[i].Columns[j].Value2;
                }
            }

            try
            {
                string matchedHeading = Process.ExtractTop(searchText, headings.Keys).First(x => x.Score > 90).Value;
                ((Excel.Worksheet)xlApp.ActiveSheet).Hyperlinks.Add(selectedRange.Rows[i], "",headings[matchedHeading]);
            }
            catch (Exception)
            {
                failedHyperlinks.Add(searchText);
            }
        }

        string generalExceptionsMessage = "";
        if (generalExceptions.Count > 0)
        {
            generalExceptionsMessage =
                "\n\nThe following general exceptions were encountered." +
                "\nPlease report these to the developer(s): \n  • " +
                string.Join("\n  • ", generalExceptions);
        }
        
        if (failedHyperlinks.Count > 0)
        {
            Alert alert = new()
            {
                AlertCaption =
                {
                    Text = "Warning: Section Headings Not Found"
                },
                AlertBody =
                {
                    Text = 
                        "The following headings could not be found in this sheet: \n  • " +
                        string.Join("\n  • ", failedHyperlinks) +
                        generalExceptionsMessage
                },
            };

            alert.Show();
        }
    }
}
