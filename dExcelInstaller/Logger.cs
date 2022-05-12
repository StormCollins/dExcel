namespace dExcelInstaller;

using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Brushes = System.Windows.Media.Brushes;

public class Logger
{
    private string loggedText = "";
    public readonly RichTextBox LogWindow;

    public string WarningText
    {
        get => loggedText;
        set
        {
            var loggerText = new FlowDocument();
            loggedText += "\n" + value;
            var timeStamp = new Run($"[{DateTime.Now:hh:mm:ss}]  ")
            {
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.Orange,
                FontFamily = new FontFamily("Courier")
            };

            var message = new Run($"{value}")
            {
                Foreground = Brushes.Orange,
                FontFamily = new FontFamily("Courier")
            };

            var paragraph = new Paragraph();
            paragraph.Inlines.Add(timeStamp);
            paragraph.Inlines.Add(message);
            loggerText.Blocks.Add(paragraph);
            LogWindow.Document.Blocks.Add(paragraph);
            LogWindow.ScrollToEnd();
        }
    }

    public string ErrorText
    {
        get => loggedText;
        set
        {
            var loggerText = new FlowDocument();
            loggedText += "\n" + value;
            var timeStamp = new Run($"{DateTime.Now:hh:mm:ss}:  ")
            {
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.Red,
                FontFamily = new FontFamily("Courier")
            };

            var message = new Run($"{value}")
            {
                FontWeight = FontWeights.DemiBold,
                Foreground = Brushes.Red,
                FontFamily = new FontFamily("Courier")
            };

            var paragraph = new Paragraph();
            paragraph.Inlines.Add(timeStamp);
            paragraph.Inlines.Add(message);
            loggerText.Blocks.Add(paragraph);
            LogWindow.Document.Blocks.Add(paragraph);
            LogWindow.ScrollToEnd();
        }
    }

    public string DontPanicErrorText
    {
        get => loggedText;
        set
        {
            var loggerText = new FlowDocument();
            loggedText += "\n" + value;
            var timeStamp = new Run($"{DateTime.Now:hh:mm:ss}:  ")
            {
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.Red,
                FontFamily = new FontFamily("Courier")
            };

            var dontPanic = new Run($"DON'T PANIC: ")
            {
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.Red,
                FontFamily = new FontFamily("Courier")
            };

            var dontPanicExplanation =
                new Run("This error is being 'gracefully' dealt with.")
                {
                    Foreground = Brushes.Red,
                    FontWeight = FontWeights.DemiBold,
                    FontFamily = new FontFamily("Courier")
                };

            var message = new Run($"{value}")
            {
                Foreground = Brushes.Red,
                FontFamily = new FontFamily("Courier")
            };

            var paragraph = new Paragraph();
            paragraph.Inlines.Add(timeStamp);
            paragraph.Inlines.Add(dontPanic);
            paragraph.Inlines.Add(dontPanicExplanation);
            loggerText.Blocks.Add(paragraph);
            LogWindow.Document.Blocks.Add(paragraph);
            LogWindow.ScrollToEnd();

            paragraph = new Paragraph();
            paragraph.Inlines.Add(timeStamp);
            paragraph.Inlines.Add(message);
            loggerText.Blocks.Add(paragraph);
            LogWindow.Document.Blocks.Add(paragraph);
            LogWindow.ScrollToEnd();
        }
    }

    public static Run CreateTimeStamp()
    {
        return new Run($"[{DateTime.Now:HH:mm:ss}]  ")
        {
            FontWeight = FontWeights.Bold,
            FontFamily = new FontFamily("Calibri")
        };
    }

    public static Run CreateMessage(string message)
    {
        return new Run($"{message}")
        {
            FontFamily = new FontFamily("Calibri"),
            FontWeight = FontWeights.Regular,
        };
    }

    public static Run CreateBoldMessage(string message)
    {
        return new Run($"{message}")
        {
            FontFamily = new FontFamily("Calibri"),
            FontSize = 16,
            FontWeight = FontWeights.ExtraBold,
        };
    }

    public string OkayText
    {
        get => loggedText;
        set
        {
            loggedText += "\n" + value;
            var timeStamp = CreateTimeStamp();
            var message = CreateMessage(value);
            var paragraph = new Paragraph();
            paragraph.Inlines.Add(timeStamp);
            paragraph.Inlines.Add(message);
            LogWindow.Document.Blocks.Add(paragraph);
            LogWindow.ScrollToEnd();
        }
    }

    /// <summary>
    /// Constructor creating instance of <see cref="Logger"/>.
    /// </summary>
    /// <param name="logWindow">XAML log window.</param>
    public Logger(RichTextBox logWindow)
    {
        LogWindow = logWindow;
    }

    public void NewProcess(string message)
    {
        var timeStamp = CreateTimeStamp();
        var messageRun = CreateBoldMessage(message);
        var paragraph = new Paragraph();
        paragraph.Inlines.Clear();
        paragraph.Inlines.Add(timeStamp);
        paragraph.Inlines.Add(messageRun);
        LogWindow.Document.Blocks.Add(paragraph);
        LogWindow.ScrollToEnd();
    }

    public void HorizontalLine(Brush color, char lineCharacter = '-', int lineLength = 97)
    {
        var horizontalLine = new Run(new string(lineCharacter, lineLength))
        {
            FontWeight = FontWeights.Bold,
            Foreground = color
        };

        var paragraph = new Paragraph();
        paragraph.Inlines.Add(horizontalLine);
        LogWindow.Document.Blocks.Add(paragraph);
    }

    public void DashedHorizontalLine(Brush? color = null)
    {
        if (color == null)
        {
            color = (Brush)Application.Current.Resources["SecondaryHueMidBrush"];
        }
        var dashedLine = string.Concat(Enumerable.Repeat("-  ", 56));
        var horizontalLine = new Run(dashedLine)
        {
            FontWeight = FontWeights.Bold,
            Foreground = color
        };

        var paragraph = new Paragraph();
        paragraph.Inlines.Add(horizontalLine);
        LogWindow.Document.Blocks.Add(paragraph);
    }

    public void ExtractionComplete(string message)
    {
        DashedHorizontalLine((Brush)Application.Current.Resources["SecondaryAccentBrush"]);

        var completeRun = new Run($"\n>>>>>>>>  Complete : ")
        {
            FontWeight = FontWeights.Bold,
            Foreground = (Brush)Application.Current.Resources["SecondaryAccentBrush"],
            FontFamily = new FontFamily("Courier")
        };


        var messageRun = new Run($" {message} ")
        {
            Foreground = (Brush)Application.Current.Resources["SecondaryAccentBrush"],
            FontFamily = new FontFamily("Courier")
        };

        var chevrons = new Run("  <<<<<<<<")
        {
            FontWeight = FontWeights.Bold,
            Foreground = (Brush)Application.Current.Resources["SecondaryAccentBrush"],
            FontFamily = new FontFamily("Courier")
        };

        var paragraph = new Paragraph();
        paragraph.Inlines.Clear();
        paragraph.Inlines.Add(completeRun);
        paragraph.Inlines.Add(messageRun);
        paragraph.Inlines.Add(chevrons);
        LogWindow.Document.Blocks.Add(paragraph);
        LogWindow.ScrollToEnd();
    }

    
    }
