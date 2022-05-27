namespace dExcelInstaller;

using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
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
        CreateMessage($"[{DateTime.Now:HH:mm:ss}]  ", fontColor, true);

    /// <summary>
    /// Creates a formatted message without a time stamp.
    /// </summary>
    /// <param name="message">The message.</param>
    /// <param name="fontColor">The font color.</param>
    /// <param name="isBold">Set to true to make the font bold.</param>
    /// <returns>A formatted message.</returns>
    private static Run CreateMessage(string message, Brush? fontColor = null, bool isBold = false)
    {
        return new Run($"{message}")
        {
            FontFamily = new FontFamily("Calibri"),
            FontWeight = isBold ? FontWeights.ExtraBold : FontWeights.Regular,
            Foreground = fontColor ?? PrimaryHueMidBrush,
        };
    }

    private void CreateTimeStampedMessage(string message, Brush? fontColor = null)
    {
        var loggerText = new FlowDocument();
        var timeStamp = CreateTimeStamp(fontColor ?? PrimaryHueMidBrush);
        var messageWithoutTimeStamp = CreateMessage(message, fontColor ?? PrimaryHueMidBrush);
        var paragraph = new Paragraph();
        paragraph.Inlines.Add(timeStamp);
        paragraph.Inlines.Add(messageWithoutTimeStamp);
        loggerText.Blocks.Add(paragraph);
        _logWindow.Document.Blocks.Add(paragraph);
        _logWindow.ScrollToEnd();
    }

    public void NewProcess(string message)
    {
        var timeStamp = CreateTimeStamp(PrimaryHueLightBrush);
        var messageRun = CreateMessage(message, PrimaryHueLightBrush, true);
        var paragraph = new Paragraph();
        paragraph.Inlines.Clear();
        paragraph.Inlines.Add("\n");
        paragraph.Inlines.Add(timeStamp);
        paragraph.Inlines.Add(messageRun);
        _logWindow.Document.Blocks.Add(paragraph);
        _logWindow.ScrollToEnd();
    }

    public void NewSubProcess(string message)
    {
        var timeStamp = CreateTimeStamp(PrimaryHueLightBrush);
        var messageRun = CreateMessage(message, PrimaryHueLightBrush, true);
        var paragraph = new Paragraph();
        paragraph.Inlines.Clear();
        paragraph.Inlines.Add(DashedHorizontalLine());
        paragraph.Inlines.Add(timeStamp);
        paragraph.Inlines.Add(messageRun);
        _logWindow.Document.Blocks.Add(paragraph);
        _logWindow.ScrollToEnd();
    }

    private Run DashedHorizontalLine(Brush? fontColor = null)
    {
        var repeats = _logWindow.Width;
        return CreateMessage(string.Concat(Enumerable.Repeat("-  ", (int) repeats / 11)) + '\n', fontColor);
    }

    public void InstallationFailed() => EndProcessMessage("Installation Failed", ErrorBrush);

    public void InstallationSucceeded() => EndProcessMessage("Installation Succeeded", PrimaryHueLightBrush);

    public void UninstallationFailed() => EndProcessMessage("Uninstallation Failed", ErrorBrush);
    
    public void UninstallationSucceeded() => EndProcessMessage("Uninstallation Succeeded", PrimaryHueLightBrush);

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
        paragraph.Inlines.Add(formattedMessage);
        paragraph.Inlines.Add(DashedHorizontalLine(fontColor));
        _logWindow.Document.Blocks.Add(paragraph);
        _logWindow.ScrollToEnd();
    }
}
