using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using System.Diagnostics;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Avalonia.Platform.Storage;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using Avalonia;

namespace RemoveDocxMetadataUI
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            BeginMoveDrag(e);
        }

        private void Exit(object? sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OnTitleClick(object? sender, RoutedEventArgs e)
        {
            OpenLinkInBrowser("https://github.com/krolchonok/DocxRemoverMeta");
        }

        private void OnAuthorClick(object? sender, RoutedEventArgs e)
        {
            OpenLinkInBrowser("https://github.com/krolchonok");
        }

        private async void SelectAndProcessFiles(object? sender, RoutedEventArgs e)
        {
            var files = await SelectDocxFiles();
            if (files != null && files.Count > 0)
            {
                var results = await RemoveMetadataFromFiles(files);
                ShowResults(results);
            }
        }

        private async Task<List<string>> SelectDocxFiles()
        {
            var selectedFiles = new List<string>();

            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null) return selectedFiles;

            var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Выберите DOCX файлы",
                AllowMultiple = true,
                FileTypeFilter = new[]
                {
                    new FilePickerFileType("DOCX документы")
                    {
                        Patterns = new[] { "*.docx" }
                    }
                }
            });

            if (files != null)
            {
                foreach (var file in files)
                {
                    selectedFiles.Add(file.Path.LocalPath);
                }
            }

            return selectedFiles;
        }

        private async Task<Dictionary<string, bool>> RemoveMetadataFromFiles(List<string> filePaths)
        {
            var results = new Dictionary<string, bool>();

            await Task.Run(() =>
            {
                foreach (var filePath in filePaths)
                {
                    try
                    {
                        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                        {
                            var coreProps = doc.PackageProperties;
                            if (coreProps != null)
                            {
                                coreProps.Creator = null;
                                coreProps.LastModifiedBy = null;
                                coreProps.Revision = null;
                                coreProps.Category = null;
                                coreProps.ContentStatus = null;
                                coreProps.Created = null;
                                coreProps.Description = null;
                                coreProps.Identifier = null;
                                coreProps.Keywords = null;
                                coreProps.Language = null;
                                coreProps.LastPrinted = null;
                                coreProps.Modified = null;
                                coreProps.Subject = null;
                                coreProps.Title = null;
                                coreProps.Version = null;
                            }

                            if (doc.CustomFilePropertiesPart != null)
                            {
                                doc.DeletePart(doc.CustomFilePropertiesPart);
                            }

                            doc.Save();
                        }
                        results[filePath] = true;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Ошибка при обработке файла {filePath}: {ex.Message}");
                        results[filePath] = false;
                    }
                }
            });

            return results;
        }

        private async void ShowResults(Dictionary<string, bool> results)
        {
            var successCount = results.Count(r => r.Value);
            var failCount = results.Count - successCount;

            var messageText = $"Processed File: {results.Count}\n" +
                             $"Sucsess: {successCount}\n" +
                             $"Error: {failCount}";

            var exitButton = new Button
            {
                Content = "[e]xit",
                Margin = new Thickness(10),
                Background = Avalonia.Media.Brushes.Black,
                BorderBrush = Avalonia.Media.Brushes.White,
                BorderThickness = new Thickness(2),
                CornerRadius = new CornerRadius(0),
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Bottom,
            };

            var resultWindow = new Window
            {
                Title = "Result",
                Width = 300,
                Height = 150,
                FontFamily = "Hermit",
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                SystemDecorations = SystemDecorations.None, // Убираем встроенный статус-бар
                Content = new Border
                {
                    BorderBrush = Avalonia.Media.Brushes.White, // Цвет обводки
                    BorderThickness = new Thickness(1), // Толщина обводки
                    Child = new Grid
                    {
                        RowDefinitions = new RowDefinitions("*,Auto"),
                        Children =
                {
                    new TextBlock
                    {
                        Text = messageText,
                        HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                        VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
                    },
                    exitButton
                }
                    }
                }
            };

            resultWindow.PointerPressed += (sender, e) =>
            {
                if (e.GetCurrentPoint(resultWindow).Properties.IsLeftButtonPressed)
                {
                    resultWindow.BeginMoveDrag(e);
                }
            };

            exitButton.Click += (sender, e) =>
            {
                resultWindow.Close();
            };
            await resultWindow.ShowDialog(this);
        }

        private void OpenLinkInBrowser(string url)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Не удалось открыть ссылку: {ex.Message}");
            }
        }

        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.S: // Нажатие клавиши "S"
                    SelectAndProcessFiles(sender, new RoutedEventArgs()); // Передаем RoutedEventArgs
                    break;
                case Key.E: // Нажатие клавиши "E"
                    Exit(sender, new RoutedEventArgs()); // Передаем RoutedEventArgs
                    break;
                default:
                    Debug.WriteLine($"Нажата клавиша: {e.Key}");
                    break;
            }
        }
    }
}