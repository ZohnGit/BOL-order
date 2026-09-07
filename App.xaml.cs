using System;
using System.IO;
using System.Windows;

namespace BolOrderExporter;

public partial class App : Application
{
    public App()
    {
        DispatcherUnhandledException += (_, args) =>
        {
            var message = args.Exception.ToString();
            try
            {
                var directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "BolOrderExporter");
                Directory.CreateDirectory(directory);
                File.WriteAllText(Path.Combine(directory, "startup-error.log"), message);
            }
            catch
            {
                // The message box below still reports the original startup failure.
            }

            MessageBox.Show(
                "应用发生错误：\n\n" + args.Exception.Message +
                "\n\n错误记录位置：%LOCALAPPDATA%\\BolOrderExporter\\startup-error.log",
                "BOL 订单导出工具",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            args.Handled = true;
            Shutdown(-1);
        };
    }
}
