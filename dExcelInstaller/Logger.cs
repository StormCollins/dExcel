namespace dExcelInstaller;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

public class Logger
{
    private readonly RichTextBox _logWindow;
    private static readonly Brush ErrorBrush = (Brush)Application.Current.Resources["ErrorBrush"];
    private static readonly Brush WarningBrush = (Brush)Application.Current.Resources["WarningBrush"];
    private static readonly Brush PrimaryHueMidBrush = (Brush)Application.Current.Resources["PrimaryHueMidBrush"];
    private static readonly Brush PrimaryHueLightBrush = (Brush)Application.Current.Resources["PrimaryHueLightBrush"];

    /// <summary>
    /// Sets the current 'error' text, indicating an error to the user, in the log window.
    /// </summary>
    public string ErrorText
    {
        set => CreateTimeStampedMessage(value, ErrorBrush);
    }

    /// <summary>
    /// Sets the current 'Okay' text, indicating no errors nor warnings to the user, in the log window.
    /// </summary>
    public string OkayText
    {
        set => CreateTimeStampedMessage(value);
    }

    /// <summary>
    /// Sets the current 'Warning' text, indicating warnings to the user, in the log window.
    /// </summary>
    public string WarningText
    {
        set => CreateTimeStampedMessage(value, WarningBrush);
    }

    /// <summary>
    /// Creates an instance of <see cref="Logger"/>.
    /// </summary>
    /// <param name="logWindow">XAML log window.</param>
    public Logger(RichTextBox logWindow)
    {
        _logWindow = logWindow;
    }

    /// <summary>
    /// Creates a formatted time stamp.
    /// </summary>
    /// <param name="fontColor">The font color.</param>
    /// <returns>A formatted time stamp.</returns>
    private static Run CreateTimeStamp(Brush? fontColor = null) =>
        CreateMessage($"[{DateTime.Now:HH:mm:ss}]  ", fontColor, true)[0];

    /// <summary>
    /// Creates a formatted message without a time stamp. 
    /// Place a path between 
    /// </summary>
    /// <param name="message">The message.</param>
    /// <param name="fontColor">The font color.</param>
    /// <param name="isBold">Set to true to make the font bold.</param>
    /// <returns>A formatted message.</returns>
    private static List<Run> CreateMessage(string message, Brush? fontColor = null, bool isBold = false)
    {
        var paths = Regex.Matches(message, @"(?<=\[\[).+?(?=\]\])").Select(x => x.Value).ToList();
        var highlightedStrings = Regex.Matches(message, @"(?<=\*\*).+?(?=\*\*)").Select(x => x.Value).ToList();  
        
        if (paths.Any() || highlightedStrings.Any())
        {
            var subMessages = Regex.Split(message, @"(\[\[)|(\]\])|(\*\*)").ToList();
            subMessages.RemoveAll(x => x is "[[" or "]]" or "**");

            List<Run> output = new();
            foreach (var subMessage in subMessages)
            {
                if (paths.Contains(subMessage))
                {
                    var run = new Run($"{subMessage}")
                    {
                        FontFamily = new FontFamily("Courier New"),
                        FontWeight = isBold ? FontWeights.ExtraBold : FontWeights.Regular,
                        Foreground = fontColor ?? PrimaryHueMidBrush,
                        TextDecorations = TextDecorations.Underline,
                    };

                    run.MouseEnter += LoggerPath_MouseEnter;
                    run.MouseLeave += LoggerPath_MouseLeave;
                    run.MouseLeftButtonDown += LoggerPath_MouseLeftButtonDown;
                    output.Add(run);
                }
                else if (highlightedStrings.Contains(subMessage))
                {
                    var run = new Run($"{subMessage}")
                    {
                        FontStyle = FontStyles.Italic,
                        FontWeight = FontWeights.Bold,
                        Foreground = fontColor ?? PrimaryHueMidBrush,
                    };
                    output.Add(run);
                }
                else
                {
                    output.Add(
                        new Run($"{subMessage}")
                        {
                            FontFamily = new FontFamily("Calibri"),
                            FontWeight = isBold ? FontWeights.ExtraBold : FontWeights.Regular,
                            Foreground = fontColor ?? PrimaryHueMidBrush,
                        });
                }
            }
        
            return output; 
        }
        
        return new List<Run>
        {
            new($"{message}")
            {
                FontFamily = new FontFamily("Calibri"),
                FontWeight = isBold ? FontWeights.ExtraBold : FontWeights.Regular,
                Foreground = fontColor ?? PrimaryHueMidBrush,
            } 
        };
    }

    private static void LoggerPath_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        var path = ((Run)sender).Text;
        if (Path.HasExtension(path))
        {
            Process.Start(new ProcessStartInfo(Path.GetDirectoryName(path)) { UseShellExecute = true });
        }
        else
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
    }

    private static void LoggerPath_MouseLeave(object sender, MouseEventArgs e)
    {
        ((Run)sender).FontWeight = FontWeights.Regular;
        if (((Run)sender).Foreground == PrimaryHueLightBrush)
        {
            ((Run)sender).Foreground = PrimaryHueMidBrush;
        }
    }

    private static void LoggerPath_MouseEnter(object sender, MouseEventArgs e)
    {
        ((Run)sender).FontWeight = FontWeights.Bold;
        ((Run)sender).Cursor = Cursors.Hand;

        if (((Run)sender).Foreground == PrimaryHueMidBrush)
        {
            ((Run)sender).Foreground = PrimaryHueLightBrush;
        }
    }

    /// <summary>
    /// Creates a formatted, time-stamped message.
    /// </summary>
    /// <param name="message">The message.</param>
    /// <param name="fontColor">The font color.</param>
    private void CreateTimeStampedMessage(string message, Brush? fontColor = null)
    {
        var loggerText = new FlowDocument();
        var timeStamp = CreateTimeStamp(fontColor ?? PrimaryHueMidBrush);
        var messagesWithoutTimeStamp = CreateMessage(message, fontColor ?? PrimaryHueMidBrush);
        var paragraph = new Paragraph();
        paragraph.Inlines.Add(timeStamp);
        foreach (var subMessage in messagesWithoutTimeStamp)
        {
            paragraph.Inlines.Add(subMessage);
        }
        loggerText.Blocks.Add(paragraph);
        _logWindow.Document.Blocks.Add(paragraph);
        _logWindow.ScrollToEnd();
    }

    public void NewProcess(string message)
    {
        var timeStamp = CreateTimeStamp(PrimaryHueLightBrush);
        var messageRuns = CreateMessage(message, PrimaryHueLightBrush, true);
        var paragraph = new Paragraph();
        paragraph.Inlines.Clear();
        paragraph.Inlines.Add("\n");
        paragraph.Inlines.Add(timeStamp);
        foreach (var messageRun in messageRuns)
        {
            paragraph.Inlines.Add(messageRun);
        }
        _logWindow.Document.Blocks.Add(paragraph);
        _logWindow.ScrollToEnd();
    }

    public void NewSubProcess(string message)
    {
        var timeStamp = CreateTimeStamp(PrimaryHueLightBrush);
        var messageRuns = CreateMessage(message, PrimaryHueLightBrush, true);
        var paragraph = new Paragraph();
        paragraph.Inlines.Clear();
        paragraph.Inlines.Add(DashedHorizontalLine());
        paragraph.Inlines.Add(timeStamp);
        foreach (var messageRun in messageRuns)
        {
            paragraph.Inlines.Add(messageRun);
        }
        _logWindow.Document.Blocks.Add(paragraph);
        _logWindow.ScrollToEnd();
    }

    private Run DashedHorizontalLine(Brush? fontColor = null)
    {
        var repeats = _logWindow.Width;
        return CreateMessage(string.Concat(Enumerable.Repeat("-  ", (int) repeats / 11)) + '\n', fontColor)[0];
    }

    public void InstallationFailed() => EndProcessMessage("Installation Failed", ErrorBrush);

    public void InstallationSucceeded() => EndProcessMessage("Installation Succeeded", PrimaryHueLightBrush);

    public void UninstallationFailed() => EndProcessMessage("Uninstallation Failed", ErrorBrush);
    
    public void UninstallationSucceeded() => EndProcessMessage("Uninstallation Succeeded", PrimaryHueLightBrush);

    public void ProcessSucceeded(string message) => EndProcessMessage(message, PrimaryHueLightBrush);

    public void ProcessFailed(string message) => EndProcessMessage(message, ErrorBrush);

    private void EndProcessMessage(string message, Brush fontColor)
    {
        var formattedMessage =
            CreateMessage(
                message: $">>>>>> {message} <<<<<<<   \n",
                isBold: true,
                fontColor: fontColor);
        var paragraph = new Paragraph();
        paragraph.Inlines.Clear();
        paragraph.TextAlignment = TextAlignment.Center;
        paragraph.Inlines.Add(DashedHorizontalLine(fontColor));
        paragraph.Inlines.Add(formattedMessage[0]);
        paragraph.Inlines.Add(DashedHorizontalLine(fontColor));
        _logWindow.Document.Blocks.Add(paragraph);
        _logWindow.ScrollToEnd();
    }
}
