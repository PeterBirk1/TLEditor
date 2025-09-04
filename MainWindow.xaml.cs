using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Input;
using Microsoft.Win32;
using Markdig;

namespace TLEditor
{
    public partial class MainWindow : Window
    {
        private string currentFilePath;
        private string currentFolderPath;
        private Dictionary<TreeViewItem, int> structureLineMap;
        private bool hasUnsavedChanges;
        private string configFilePath;
        private bool showFoldersState = true;
        private bool showStructureState = true;

        public MainWindow()
        {
            InitializeComponent();
            structureLineMap = new Dictionary<TreeViewItem, int>();
            configFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TLEditor", "config.txt");
            InitializeEditor();
            LoadSavedFolder();
        }

        private void InitializeEditor()
        {
            // Enable undo/redo
            Editor.IsUndoEnabled = true;
        }

        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Markdown files (*.md)|*.md|All files (*.*)|*.*",
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                // Auto-save current file before opening new one
                AutoSaveCurrentFile();
                
                LoadFile(openFileDialog.FileName);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(currentFilePath))
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Markdown files (*.md)|*.md|All files (*.*)|*.*",
                    RestoreDirectory = true,
                    DefaultExt = "md"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    currentFilePath = saveFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }

            SaveCurrentFile();
        }

        private void SaveAsButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Markdown files (*.md)|*.md|All files (*.*)|*.*",
                RestoreDirectory = true,
                DefaultExt = "md"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                currentFilePath = saveFileDialog.FileName;
                SaveCurrentFile();
            }
        }

        private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            
            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                currentFolderPath = folderDialog.SelectedPath;
                LoadFolderStructure(currentFolderPath);
                SaveFolderToConfig(currentFolderPath);
            }
        }




        private void LoadFile(string filePath)
        {
            try
            {
                currentFilePath = filePath;
                // Load as plain text (.md files)
                string content = File.ReadAllText(filePath, Encoding.UTF8);
                Editor.Document.Blocks.Clear();
                Editor.Document.Blocks.Add(new Paragraph(new Run(content)));
                
                // Update document structure based on plain text content
                // Add a small delay to ensure the RichTextBox has finished loading the content
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    string plainTextContent = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text;
                    UpdateDocumentStructure(plainTextContent);
                }), System.Windows.Threading.DispatcherPriority.Background);
                
                Title = $"TL Markdown Editor - {Path.GetFileName(filePath)}";
                hasUnsavedChanges = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveCurrentFile()
        {
            try
            {
                // Save as plain text (.md files)
                string content = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text;
                File.WriteAllText(currentFilePath, content, Encoding.UTF8);
                
                Title = $"TL Markdown Editor - {Path.GetFileName(currentFilePath)}";
                hasUnsavedChanges = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void UpdateDocumentStructure(string content)
        {
            DocumentStructure.Items.Clear();
            structureLineMap.Clear();
            
            if (string.IsNullOrEmpty(content))
                return;
                
            // Use consistent line splitting method
            string[] lines = content.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            
            for (int i = 0; i < lines.Length; i++)
            {
                string originalLine = lines[i];
                string trimmedLine = originalLine.Trim();
                
                // Skip empty lines but keep processing for accurate line numbers
                if (string.IsNullOrWhiteSpace(trimmedLine))
                    continue;
                
                // Check for markdown headers (# ## ### etc.)
                if (trimmedLine.StartsWith("#"))
                {
                    int level = 0;
                    while (level < trimmedLine.Length && trimmedLine[level] == '#')
                    {
                        level++;
                    }
                    
                    if (level <= 6 && level < trimmedLine.Length && (trimmedLine[level] == ' ' || level == trimmedLine.Length))
                    {
                        string headerText = level < trimmedLine.Length ? trimmedLine.Substring(level).Trim() : "";
                        if (!string.IsNullOrEmpty(headerText))
                        {
                            TreeViewItem headerItem = new TreeViewItem
                            {
                                Header = $"{new string(' ', (level - 1) * 2)}{headerText}",
                                Tag = $"Line {i + 1}",
                                ToolTip = $"Line {i + 1}: {headerText}"
                            };
                            
                            // Store the actual line number (0-based for NavigateToLine)
                            structureLineMap[headerItem] = i;
                            DocumentStructure.Items.Add(headerItem);
                        }
                    }
                }
            }
            
            // If no headers found, add a fallback message
            if (DocumentStructure.Items.Count == 0)
            {
                TreeViewItem noHeaders = new TreeViewItem
                {
                    Header = "No headers found",
                    Foreground = Brushes.LightGray,
                    IsEnabled = false
                };
                DocumentStructure.Items.Add(noHeaders);
            }
        }


        private void DocumentStructure_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            TreeViewItem selectedItem = e.NewValue as TreeViewItem;
            if (selectedItem != null && selectedItem.IsEnabled && structureLineMap.ContainsKey(selectedItem))
            {
                try
                {
                    int lineNumber = structureLineMap[selectedItem];
                    NavigateToLine(lineNumber);
                }
                catch (Exception ex)
                {
                    // Log navigation error but don't show to user
                    System.Diagnostics.Debug.WriteLine($"Navigation error: {ex.Message}");
                }
            }
        }

        private void FolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            TreeViewItem selectedItem = e.NewValue as TreeViewItem;
            if (selectedItem?.Tag is string filePath && File.Exists(filePath) && 
                filePath.EndsWith(".md"))
            {
                // Auto-save current file before switching
                AutoSaveCurrentFile();
                
                LoadFile(filePath);
            }
        }

        private void AutoSaveCurrentFile()
        {
            if (!string.IsNullOrEmpty(currentFilePath) && hasUnsavedChanges)
            {
                try
                {
                    // Save as plain text (.md files)
                    string content = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text;
                    File.WriteAllText(currentFilePath, content, Encoding.UTF8);
                    
                    // Update title to remove asterisk
                    Title = $"TL Markdown Editor - {Path.GetFileName(currentFilePath)}";
                    hasUnsavedChanges = false;
                }
                catch (Exception)
                {
                    // If auto-save fails, just log it but don't show error to user
                    // Could optionally show a brief status message
                }
            }
        }

        private void LoadFolderStructure(string folderPath)
        {
            FolderTree.Items.Clear();
            
            try
            {
                TreeViewItem rootItem = new TreeViewItem
                {
                    Header = CreateFolderHeader(Path.GetFileName(folderPath), true),
                    Tag = folderPath
                };
                
                LoadDirectoryItems(rootItem, folderPath);
                rootItem.IsExpanded = true;
                FolderTree.Items.Add(rootItem);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading folder: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadDirectoryItems(TreeViewItem parentItem, string directoryPath)
        {
            try
            {
                var directories = Directory.GetDirectories(directoryPath)
                    .OrderBy(d => Path.GetFileName(d));
                
                foreach (string directory in directories)
                {
                    TreeViewItem dirItem = new TreeViewItem
                    {
                        Header = CreateFolderHeader(Path.GetFileName(directory), false),
                        Tag = directory
                    };
                    
                    LoadDirectoryItems(dirItem, directory);
                    parentItem.Items.Add(dirItem);
                }

                // Get only .md files
                var allFiles = Directory.GetFiles(directoryPath, "*.md").OrderBy(f => Path.GetFileName(f));
                
                foreach (string file in allFiles)
                {
                    TreeViewItem fileItem = new TreeViewItem
                    {
                        Header = CreateFileHeader(Path.GetFileName(file), file),
                        Tag = file
                    };
                    
                    parentItem.Items.Add(fileItem);
                }
            }
            catch (Exception)
            {
                // Ignore access denied errors for certain directories
            }
        }

        private void NavigateToLine(int lineNumber)
        {
            try
            {
                // Get the plain text content
                string content = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text;
                if (string.IsNullOrEmpty(content))
                    return;
                
                // Use the same line splitting method as UpdateDocumentStructure
                string[] lines = content.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                
                if (lineNumber < 0 || lineNumber >= lines.Length)
                    return;
                
                // Calculate character position by counting characters up to the target line
                int charPosition = 0;
                for (int i = 0; i < lineNumber; i++)
                {
                    charPosition += lines[i].Length;
                    // Add 1 for line break (simplified approach)
                    if (i < lines.Length - 1) // Don't add line break after last line
                    {
                        charPosition += 1;
                    }
                }
                
                // Get text pointer at the calculated position
                TextPointer start = Editor.Document.ContentStart;
                TextPointer target = start.GetPositionAtOffset(Math.Min(charPosition, content.Length));
                
                if (target != null)
                {
                    // Position cursor at the calculated position
                    Editor.CaretPosition = target;
                    Editor.Focus();
                    
                    // Clear selection
                    Editor.Selection.Select(target, target);
                    
                    // Force scroll to make the position visible
                    Rect rect = target.GetCharacterRect(LogicalDirection.Forward);
                    if (!rect.IsEmpty)
                    {
                        double targetY = rect.Y - Editor.ActualHeight / 3;
                        Editor.ScrollToVerticalOffset(Math.Max(0, targetY));
                    }
                }
            }
            catch (Exception)
            {
                // Navigation failed silently
            }
        }

        private void Editor_TextChanged(object sender, TextChangedEventArgs e)
        {
            hasUnsavedChanges = true;
            if (!string.IsNullOrEmpty(currentFilePath) && !Title.EndsWith(" *"))
            {
                Title += " *";
            }
            
            // Update document structure when text changes
            try
            {
                string content = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text;
                UpdateDocumentStructure(content);
            }
            catch (Exception)
            {
                // If structure update fails, just continue - it's not critical
            }
        }

        // Rich Formatting Event Handlers
        private void UndoButton_Click(object sender, RoutedEventArgs e)
        {
            if (Editor.CanUndo)
                Editor.Undo();
        }

        private void RedoButton_Click(object sender, RoutedEventArgs e)
        {
            if (Editor.CanRedo)
                Editor.Redo();
        }

        private void FindReplaceButton_Click(object sender, RoutedEventArgs e)
        {
            FindReplaceDialog dialog = new FindReplaceDialog(Editor);
            dialog.Owner = this;
            dialog.Show();
        }


        private void BulletListButton_Click(object sender, RoutedEventArgs e)
        {
            InsertList(false);
        }

        private void NumberListButton_Click(object sender, RoutedEventArgs e)
        {
            InsertList(true);
        }

        private void InsertList(bool isNumbered)
        {
            if (Editor.Selection != null)
            {
                List list = new List();
                if (isNumbered)
                {
                    list.MarkerStyle = TextMarkerStyle.Decimal;
                }
                else
                {
                    list.MarkerStyle = TextMarkerStyle.Disc;
                }

                // Add 3 list items as default
                for (int i = 0; i < 3; i++)
                {
                    ListItem item = new ListItem();
                    item.Blocks.Add(new Paragraph(new Run($"List item {i + 1}")));
                    list.ListItems.Add(item);
                }

                // Insert the list at the current position
                var insertionPosition = Editor.CaretPosition;
                var paragraph = insertionPosition.Paragraph;
                
                if (paragraph != null)
                {
                    paragraph.SiblingBlocks.InsertAfter(paragraph, list);
                }
                else
                {
                    Editor.Document.Blocks.Add(list);
                }
            }
        }

        private void InsertTableButton_Click(object sender, RoutedEventArgs e)
        {
            TableInsertDialog dialog = new TableInsertDialog();
            dialog.Owner = this;
            
            if (dialog.ShowDialog() == true)
            {
                InsertTable(dialog.Rows, dialog.Columns);
            }
        }

        private void InsertTable(int rows, int columns)
        {
            // Build markdown table
            StringBuilder markdownTable = new StringBuilder();
            
            // Header row
            markdownTable.Append("|");
            for (int c = 0; c < columns; c++)
            {
                markdownTable.Append($" Header {c + 1} |");
            }
            markdownTable.AppendLine();
            
            // Separator row
            markdownTable.Append("|");
            for (int c = 0; c < columns; c++)
            {
                markdownTable.Append("---------|");
            }
            markdownTable.AppendLine();
            
            // Data rows
            for (int r = 0; r < rows - 1; r++) // -1 because header is already created
            {
                markdownTable.Append("|");
                for (int c = 0; c < columns; c++)
                {
                    markdownTable.Append($" Cell {r + 1},{c + 1} |");
                }
                markdownTable.AppendLine();
            }
            
            // Insert the markdown table at the current position
            var insertionPosition = Editor.CaretPosition;
            
            // Insert the table text
            insertionPosition.InsertTextInRun(markdownTable.ToString());
            
            // Move cursor to after the inserted table
            var newPosition = insertionPosition.GetPositionAtOffset(markdownTable.ToString().Length);
            if (newPosition != null)
            {
                Editor.CaretPosition = newPosition;
            }
            
            Editor.Focus();
        }

        // Menu Event Handlers
        private void ExitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            // Auto-save current file before closing
            AutoSaveCurrentFile();
            base.OnClosing(e);
        }

        private void ShowFoldersMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // Toggle the state
            showFoldersState = !showFoldersState;
            
            // Update button text to reflect state
            if (sender is Button button)
            {
                button.Content = showFoldersState ? "📁 Hide Folders" : "📁 Show Folders";
            }
            
            // Get the main grid
            Grid mainGrid = (Grid)this.Content;
            
            if (showFoldersState)
            {
                // Show folders panel
                LeftPanel.Visibility = Visibility.Visible;
                LeftSplitter.Visibility = Visibility.Visible;
                mainGrid.ColumnDefinitions[0].Width = new GridLength(250, GridUnitType.Pixel);
                mainGrid.ColumnDefinitions[0].MinWidth = 200;
                mainGrid.ColumnDefinitions[1].Width = new GridLength(5, GridUnitType.Pixel);
            }
            else
            {
                // Hide folders panel
                LeftPanel.Visibility = Visibility.Collapsed;
                LeftSplitter.Visibility = Visibility.Collapsed;
                mainGrid.ColumnDefinitions[0].MinWidth = 0;
                mainGrid.ColumnDefinitions[0].Width = new GridLength(0, GridUnitType.Pixel);
                mainGrid.ColumnDefinitions[1].Width = new GridLength(0, GridUnitType.Pixel);
            }
        }

        private void ShowStructureMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // Toggle the state
            showStructureState = !showStructureState;
            
            // Update button text to reflect state
            if (sender is Button button)
            {
                button.Content = showStructureState ? "📋 Hide Structure" : "📋 Show Structure";
            }
            
            // Get the main grid
            Grid mainGrid = (Grid)this.Content;
            
            if (showStructureState)
            {
                // Show structure panel
                RightPanel.Visibility = Visibility.Visible;
                RightSplitter.Visibility = Visibility.Visible;
                mainGrid.ColumnDefinitions[4].Width = new GridLength(250, GridUnitType.Pixel);
                mainGrid.ColumnDefinitions[4].MinWidth = 200;
                mainGrid.ColumnDefinitions[3].Width = new GridLength(5, GridUnitType.Pixel);
            }
            else
            {
                // Hide structure panel
                RightPanel.Visibility = Visibility.Collapsed;
                RightSplitter.Visibility = Visibility.Collapsed;
                mainGrid.ColumnDefinitions[4].MinWidth = 0;
                mainGrid.ColumnDefinitions[4].Width = new GridLength(0, GridUnitType.Pixel);
                mainGrid.ColumnDefinitions[3].Width = new GridLength(0, GridUnitType.Pixel);
            }
        }

        // Helper methods for creating headers with icons
        private string CreateFolderHeader(string name, bool isRootFolder)
        {
            string icon = "📁";
            return $"{icon} {name}";
        }

        private string CreateFileHeader(string fileName, string fullPath)
        {
            string icon = GetFileIcon(fullPath);
            return $"{icon} {fileName}";
        }

        private string GetFileIcon(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLower();
            switch (extension)
            {
                case ".md":
                    return "●"; // Markdown file
                default:
                    return "◎"; // Generic file
            }
        }

        private void SaveFolderToConfig(string folderPath)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(configFilePath));
                File.WriteAllText(configFilePath, folderPath);
            }
            catch (Exception)
            {
                // Silently ignore config save errors
            }
        }

        private void LoadSavedFolder()
        {
            try
            {
                if (File.Exists(configFilePath))
                {
                    string savedPath = File.ReadAllText(configFilePath).Trim();
                    if (Directory.Exists(savedPath))
                    {
                        currentFolderPath = savedPath;
                        LoadFolderStructure(currentFolderPath);
                    }
                }
            }
            catch (Exception)
            {
                // Silently ignore config load errors
            }
        }

    }
}