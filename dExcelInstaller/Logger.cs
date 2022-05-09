namespace dExcelInstaller;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Brushes = System.Windows.Media.Brushes;

public class Logger
{
    private string loggedText = "";
    public readonly RichTextBox LogWindow;

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

    public void DashedHorizontalLine(Brush color)
    {
        var dashedLine = string.Concat(Enumerable.Repeat("-  ", 43));
        var horizontalLine = new Run(dashedLine)
        {
            FontWeight = FontWeights.Bold,
            Foreground = color
        };

        var paragraph = new Paragraph();
        paragraph.Inlines.Add(horizontalLine);
        LogWindow.Document.Blocks.Add(paragraph);
    }

    public void NewProcess(string message)
    {
        DashedHorizontalLine((Brush)Application.Current.Resources["SecondaryAccentBrush"]);

        var timeStamp = new Run($"{DateTime.Now:hh:mm:ss}:  ")
        {
            FontWeight = FontWeights.Bold,
            Foreground = (Brush)Application.Current.Resources["SecondaryAccentBrush"],
            FontFamily = new FontFamily("Courier")
        };

        var messageRun = new Run($"{message}")
        {
            FontWeight = FontWeights.Bold,
            Foreground = (Brush)Application.Current.Resources["SecondaryAccentBrush"],
            FontFamily = new FontFamily("Courier")
        };

        var paragraph = new Paragraph();
        paragraph.Inlines.Clear();
        paragraph.Inlines.Add(timeStamp);
        paragraph.Inlines.Add(messageRun);
        LogWindow.Document.Blocks.Add(paragraph);
        LogWindow.ScrollToEnd();
    }

    public void NewExtraction(string message)
    {
        HorizontalLine((Brush)Application.Current.Resources["SecondaryAccentBrush"]);

        var timeStamp = new Run($"{DateTime.Now:hh:mm:ss}:  ")
        {
            FontWeight = FontWeights.Bold,
            Foreground = (Brush)Application.Current.Resources["SecondaryAccentBrush"],
            FontFamily = new FontFamily("Courier")
        };

        var messageRun = new Run($"{message}")
        {
            FontWeight = FontWeights.Bold,
            FontStyle = FontStyles.Italic,
            Foreground = (Brush)Application.Current.Resources["SecondaryAccentBrush"],
            FontFamily = new FontFamily("Courier")
        };

        var paragraph = new Paragraph();
        paragraph.Inlines.Clear();
        paragraph.Inlines.Add(timeStamp);
        paragraph.Inlines.Add(messageRun);
        LogWindow.Document.Blocks.Add(paragraph);
        LogWindow.ScrollToEnd();
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

    public string WarningText
    {
        get => loggedText;
        set
        {
            var loggerText = new FlowDocument();
            loggedText += "\n" + value;
            var timeStamp = new Run($"{DateTime.Now:hh:mm:ss}:  ")
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

        public string OkayText
        {
            get => loggedText;
            set
            {
                var loggerText = new FlowDocument();
                loggedText += "\n" + value;
                var timeStamp = new Run($"{DateTime.Now:hh:mm:ss}:  ")
                {
                    FontWeight = FontWeights.Bold,
                    Foreground = (Brush)Application.Current.Resources["SecondaryAccentBrush"],
                    FontFamily = new FontFamily("Courier")
                };

                var message = new Run($"{value}")
                {
                    Foreground = (Brush)Application.Current.Resources["SecondaryAccentBrush"],
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

        public string WarningTextWithoutTime
        {
            get => loggedText;
            set
            {
                var loggerText = new FlowDocument();
                loggedText += "\n" + value;

                var message = new Run($"{value}")
                {
                    Foreground = Brushes.Orange,
                    FontFamily = new FontFamily("Courier")
                };

                var paragraph = new Paragraph();
                paragraph.Inlines.Add(message);
                loggerText.Blocks.Add(paragraph);
                LogWindow.Document.Blocks.Add(paragraph);
                LogWindow.ScrollToEnd();
            }
        }


        public string ErrorTextWithoutTime
        {
            get => loggedText;
            set
            {
                var loggerText = new FlowDocument();
                loggedText += "\n" + value;

                var message = new Run($"{value}")
                {
                    FontWeight = FontWeights.DemiBold,
                    Foreground = Brushes.Red,
                    FontFamily = new FontFamily("Courier")
                };

                var paragraph = new Paragraph();
                paragraph.Inlines.Add(message);
                loggerText.Blocks.Add(paragraph);
                LogWindow.Document.Blocks.Add(paragraph);
                LogWindow.ScrollToEnd();
            }
        }

        public string OkayTextWithoutTime
        {
            get => loggedText;
            set
            {
                var loggerText = new FlowDocument();
                loggedText += "\n" + value;

                var message = new Run($"{value}")
                {
                    Foreground = (Brush)Application.Current.Resources["SecondaryAccentBrush"],
                    FontFamily = new FontFamily("Courier")
                };

                var paragraph = new Paragraph();
                paragraph.Inlines.Add(message);
                loggerText.Blocks.Add(paragraph);
                LogWindow.Document.Blocks.Add(paragraph);
                LogWindow.ScrollToEnd();
            }
        }

        public Logger(RichTextBox rtbLogger)
        {
            LogWindow = rtbLogger;
        }
    }
