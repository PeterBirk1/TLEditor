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
    // Class to track data for each editor tab
    public class EditorTabData
    {
        public string FilePath { get; set; }
        public RichTextBox Editor { get; set; }
        public bool HasUnsavedChanges { get; set; }
        public Dictionary<TreeViewItem, int> StructureLineMap { get; set; }
        public bool IsFormattingInProgress { get; set; }
        
        public EditorTabData(string filePath = null)
        {
            FilePath = filePath;
            HasUnsavedChanges = false;
            StructureLineMap = new Dictionary<TreeViewItem, int>();
            IsFormattingInProgress = false;
            
            // Create a new RichTextBox for this tab
            Editor = new RichTextBox();
            Editor.IsUndoEnabled = true;
            Editor.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            Editor.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            Editor.Background = new SolidColorBrush(Color.FromRgb(40, 40, 40));
            Editor.Foreground = Brushes.White;
            Editor.FontFamily = new FontFamily("Consolas");
            Editor.FontSize = 14;
            
            // Event handler will be added when tab is created in MainWindow
        }
    }

    public partial class MainWindow : Window
    {
        private string currentFilePath;
        private string currentFolderPath;
        private Dictionary<TreeViewItem, int> structureLineMap;
        private bool hasUnsavedChanges;
        private string configFilePath;
        private bool showFoldersState = true;
        private bool showStructureState = true;
        
        // Tab management
        private List<EditorTabData> editorTabs;
        private int currentTabIndex = -1;

        public MainWindow()
        {
            InitializeComponent();
            structureLineMap = new Dictionary<TreeViewItem, int>();
            configFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TLEditor", "config.txt");
            
            // Initialize tab management
            editorTabs = new List<EditorTabData>();
            InitializeEditor();
            LoadSavedFolder();
        }

        private void InitializeEditor()
        {
            // Start with no tabs open
            currentTabIndex = -1;
            currentFilePath = null;
            hasUnsavedChanges = false;
            structureLineMap.Clear();
            DocumentStructure.Items.Clear();
            Title = "TL Markdown Editor";
        }

        private void NewFileButton_Click(object sender, RoutedEventArgs e)
        {
            // Create a new empty tab
            var tabData = CreateNewTab();
            
            // Create UI tab
            TabItem newTab = new TabItem();
            newTab.Header = "New Document";
            newTab.ToolTip = "Untitled document";
            newTab.Content = tabData.Editor;
            
            // Apply consistent styling to the new editor
            tabData.Editor.Background = new SolidColorBrush(Color.FromRgb(40, 40, 40));
            tabData.Editor.Foreground = Brushes.White;
            tabData.Editor.FontFamily = new FontFamily("Consolas");
            tabData.Editor.FontSize = 14;
            
            EditorTabs.Items.Add(newTab);
            EditorTabs.SelectedIndex = EditorTabs.Items.Count - 1;
            
            // Switch to new tab
            currentTabIndex = editorTabs.Count - 1;
            SwitchToTab(currentTabIndex);
            
            // Clear document structure
            DocumentStructure.Items.Clear();
            
            // Focus on the new editor
            tabData.Editor.Focus();
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
            // Check if there's a current tab to save
            if (currentTabIndex < 0 || currentTabIndex >= editorTabs.Count)
                return;

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
            // Check if there's a current tab to save
            if (currentTabIndex < 0 || currentTabIndex >= editorTabs.Count)
                return;

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

        private void CloseTabButton_Click(object sender, RoutedEventArgs e)
        {
            // Close the currently selected tab
            if (currentTabIndex >= 0 && currentTabIndex < editorTabs.Count)
            {
                CloseTab(currentTabIndex);
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
                // Check if file is already open in a tab
                for (int i = 0; i < editorTabs.Count; i++)
                {
                    if (editorTabs[i].FilePath == filePath)
                    {
                        // File already open, switch to that tab
                        EditorTabs.SelectedIndex = i;
                        SwitchToTab(i);
                        return;
                    }
                }
                
                // Create new tab for this file
                var tabData = CreateNewTab(filePath);
                
                // Load file content into the tab's editor
                string content = File.ReadAllText(filePath, Encoding.UTF8);
                tabData.Editor.Document.Blocks.Clear();
                
                // Create paragraph with content and apply proper formatting
                Paragraph paragraph = new Paragraph(new Run(content));
                tabData.Editor.Document.Blocks.Add(paragraph);
                
                // Apply base formatting to the entire document immediately
                TextRange entireDocument = new TextRange(tabData.Editor.Document.ContentStart, tabData.Editor.Document.ContentEnd);
                entireDocument.ApplyPropertyValue(TextElement.FontSizeProperty, 14.0);
                entireDocument.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal);
                entireDocument.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(Colors.White));
                entireDocument.ApplyPropertyValue(TextElement.FontFamilyProperty, new FontFamily("Consolas"));
                
                // Create UI tab with close button functionality
                TabItem newTab = new TabItem();
                newTab.Header = Path.GetFileName(filePath);
                newTab.ToolTip = filePath;
                newTab.Content = tabData.Editor;
                newTab.Tag = editorTabs.Count - 1; // Store the index for close button reference
                
                // Apply consistent styling to the new editor
                tabData.Editor.Background = new SolidColorBrush(Color.FromRgb(40, 40, 40));
                tabData.Editor.Foreground = Brushes.White;
                tabData.Editor.FontFamily = new FontFamily("Consolas");
                tabData.Editor.FontSize = 14;
                
                EditorTabs.Items.Add(newTab);
                EditorTabs.SelectedIndex = EditorTabs.Items.Count - 1;
                
                // Switch to new tab
                currentTabIndex = editorTabs.Count - 1;
                SwitchToTab(currentTabIndex);
                
                // Update document structure based on plain text content
                // Add a delay to ensure the RichTextBox has finished loading the content
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    var currentTab = GetCurrentTab();
                    if (currentTab != null)
                    {
                        string plainTextContent = new TextRange(currentTab.Editor.Document.ContentStart, currentTab.Editor.Document.ContentEnd).Text;
                        UpdateDocumentStructure(plainTextContent);
                    }
                }), System.Windows.Threading.DispatcherPriority.Background);
                
                // Skip automatic header formatting on load for better performance
                // Headers will be formatted when saving the document
                
                tabData.HasUnsavedChanges = false;
                UpdateTabTitle(tabData);
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
                var currentTab = GetCurrentTab();
                if (currentTab == null)
                {
                    return; // No tab to save
                }
                
                // Format headers before saving (only if Hybrid mode is enabled)
                if (HybridModeCheckBox.IsChecked == true)
                {
                    FormatExistingHeaders(currentTab.Editor);
                }
                
                // Save current tab's content
                string tabContent = new TextRange(currentTab.Editor.Document.ContentStart, currentTab.Editor.Document.ContentEnd).Text;
                File.WriteAllText(currentTab.FilePath, tabContent, Encoding.UTF8);
                
                // Update tab state
                currentTab.HasUnsavedChanges = false;
                UpdateTabTitle(currentTab);
                
                // Update window title
                Title = $"TL Markdown Editor - {Path.GetFileName(currentTab.FilePath)}";
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
            
            // Get the current tab's structure map, or use the global one for backward compatibility
            var currentTab = GetCurrentTab();
            var currentStructureMap = currentTab?.StructureLineMap ?? structureLineMap;
            currentStructureMap.Clear();
            
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
                            currentStructureMap[headerItem] = i;
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
            if (selectedItem != null && selectedItem.IsEnabled)
            {
                try
                {
                    // Use the current tab's structure map, or fall back to global map
                    var currentTab = GetCurrentTab();
                    var currentStructureMap = currentTab?.StructureLineMap ?? structureLineMap;
                    
                    if (currentStructureMap.ContainsKey(selectedItem))
                    {
                        int lineNumber = currentStructureMap[selectedItem];
                        NavigateToLine(lineNumber);
                    }
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
                
                // Check if file is already open in a tab
                int existingTabIndex = -1;
                for (int i = 0; i < editorTabs.Count; i++)
                {
                    if (editorTabs[i].FilePath == filePath)
                    {
                        existingTabIndex = i;
                        break;
                    }
                }
                
                if (existingTabIndex >= 0)
                {
                    // File is already open, switch to that tab
                    EditorTabs.SelectedIndex = existingTabIndex;
                    SwitchToTab(existingTabIndex);
                }
                else
                {
                    // File is not open, load it (this will create a new tab)
                    LoadFile(filePath);
                }
            }
        }

        private void FolderTree_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Only respond to left mouse button clicks
            if (e.ChangedButton != MouseButton.Left)
                return;
                
            // Find the TreeViewItem that was clicked
            TreeViewItem clickedItem = null;
            DependencyObject originalSource = e.OriginalSource as DependencyObject;
            
            // Walk up the visual tree to find the TreeViewItem
            while (originalSource != null && clickedItem == null)
            {
                if (originalSource is TreeViewItem)
                {
                    clickedItem = originalSource as TreeViewItem;
                }
                else
                {
                    originalSource = VisualTreeHelper.GetParent(originalSource);
                }
            }
            
            if (clickedItem?.Tag is string filePath && File.Exists(filePath) && 
                filePath.EndsWith(".md"))
            {
                // Auto-save current file before switching
                AutoSaveCurrentFile();
                
                // Check if file is already open in a tab
                int existingTabIndex = -1;
                for (int i = 0; i < editorTabs.Count; i++)
                {
                    if (editorTabs[i].FilePath == filePath)
                    {
                        existingTabIndex = i;
                        break;
                    }
                }
                
                if (existingTabIndex >= 0)
                {
                    // File is already open, switch to that tab
                    EditorTabs.SelectedIndex = existingTabIndex;
                    SwitchToTab(existingTabIndex);
                }
                else
                {
                    // File is not open, load it (this will create a new tab)
                    LoadFile(filePath);
                }
            }
        }

        private void AutoSaveCurrentFile()
        {
            var currentTab = GetCurrentTab();
            if (currentTab != null && !string.IsNullOrEmpty(currentTab.FilePath) && currentTab.HasUnsavedChanges)
            {
                try
                {
                    // Format headers before auto-saving (only if Hybrid mode is enabled)
                    if (HybridModeCheckBox.IsChecked == true)
                    {
                        FormatExistingHeaders(currentTab.Editor);
                    }
                    
                    // Save current tab's content
                    string content = new TextRange(currentTab.Editor.Document.ContentStart, currentTab.Editor.Document.ContentEnd).Text;
                    File.WriteAllText(currentTab.FilePath, content, Encoding.UTF8);
                    
                    // Update tab state
                    currentTab.HasUnsavedChanges = false;
                    UpdateTabTitle(currentTab);
                    
                    // Update title to remove asterisk
                    Title = $"TL Markdown Editor - {Path.GetFileName(currentTab.FilePath)}";
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
                var currentEditor = GetCurrentEditor();
                if (currentEditor == null) return;
                
                // Get the plain text content
                string content = new TextRange(currentEditor.Document.ContentStart, currentEditor.Document.ContentEnd).Text;
                if (string.IsNullOrEmpty(content))
                    return;
                
                // Use the same line splitting method as header formatting
                string[] lines = content.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                
                if (lineNumber < 0 || lineNumber >= lines.Length)
                    return;
                
                // Get the text of the target line
                string targetLineText = lines[lineNumber];
                
                // Find this specific line using text search (same approach as header formatting)
                TextPointer lineStart = FindLineStartByText(currentEditor, targetLineText, lineNumber);
                
                if (lineStart != null)
                {
                    // Move one line down to correct positioning offset
                    TextPointer adjustedPosition = lineStart.GetLineStartPosition(1) ?? lineStart;
                    // Position cursor at the beginning of the line
                    currentEditor.CaretPosition = adjustedPosition;
                    currentEditor.Focus();
                    
                    // Clear selection
                    currentEditor.Selection.Select(adjustedPosition, adjustedPosition);
                    
                    // Force scroll to make the position visible
                    Rect rect = adjustedPosition.GetCharacterRect(LogicalDirection.Forward);
                    if (!rect.IsEmpty)
                    {
                        double targetY = rect.Y - currentEditor.ActualHeight / 3;
                        currentEditor.ScrollToVerticalOffset(Math.Max(0, targetY));
                    }
                }
            }
            catch (Exception)
            {
                // Navigation failed silently
            }
        }

        private TextPointer FindLineStartByText(RichTextBox editor, string lineText, int lineNumber)
        {
            try
            {
                // Get all lines to handle duplicates properly
                string content = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd).Text;
                string[] lines = content.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                
                // Find all occurrences of this line text
                List<int> matchingLineIndexes = new List<int>();
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i] == lineText)
                    {
                        matchingLineIndexes.Add(i);
                    }
                }
                
                if (matchingLineIndexes.Count == 0)
                    return null;
                
                // Find which occurrence index corresponds to our line number
                int occurrenceIndex = matchingLineIndexes.IndexOf(lineNumber);
                if (occurrenceIndex < 0)
                    return null; // This shouldn't happen, but safety check
                
                // Search for the specific occurrence of the text in the document
                TextPointer start = editor.Document.ContentStart;
                TextPointer end = editor.Document.ContentEnd;
                TextRange documentRange = new TextRange(start, end);
                
                TextPointer foundStart = documentRange.Start;
                int foundCount = 0;
                
                while (foundStart != null && foundStart.CompareTo(documentRange.End) < 0)
                {
                    // Try to find the text starting from current position
                    TextPointer foundEnd = foundStart;
                    
                    // Move forward by the length of text we're looking for
                    for (int i = 0; i < lineText.Length && foundEnd != null; i++)
                    {
                        foundEnd = foundEnd.GetNextInsertionPosition(LogicalDirection.Forward);
                    }
                    
                    if (foundEnd != null)
                    {
                        // Check if this range contains our text
                        TextRange testRange = new TextRange(foundStart, foundEnd);
                        string rangeText = testRange.Text;
                        
                        if (rangeText == lineText)
                        {
                            // Found a match - check if it's the occurrence we want
                            if (foundCount == occurrenceIndex)
                            {
                                // Move to the actual beginning of the line (not just the text start)
                                TextPointer lineStart = foundStart.GetLineStartPosition(0);
                                return lineStart ?? foundStart; // Fallback to text start if line start fails
                            }
                            
                            // Move past this match to find the next one
                            foundCount++;
                            foundStart = foundEnd;
                            continue;
                        }
                    }
                    
                    // Move to next position
                    foundStart = foundStart.GetNextInsertionPosition(LogicalDirection.Forward);
                }
                
                return null; // Not found
            }
            catch
            {
                return null;
            }
        }

        private void Editor_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Find which tab this editor belongs to
            var senderEditor = sender as RichTextBox;
            var tabData = editorTabs.FirstOrDefault(t => t.Editor == senderEditor);
            
            if (tabData != null)
            {
                // Don't mark as changed if we're just applying formatting
                if (!tabData.IsFormattingInProgress)
                {
                    // Tab-based text change handling
                    tabData.HasUnsavedChanges = true;
                    UpdateTabTitle(tabData);
                }
                
                // If this is the current tab, update window title and document structure
                if (editorTabs.IndexOf(tabData) == currentTabIndex)
                {
                    if (!tabData.IsFormattingInProgress)
                    {
                        hasUnsavedChanges = true;
                        if (!string.IsNullOrEmpty(tabData.FilePath) && !Title.EndsWith(" *"))
                        {
                            Title += " *";
                        }
                    }
                    
                    // Update document structure when text changes (always do this)
                    try
                    {
                        string content = new TextRange(senderEditor.Document.ContentStart, senderEditor.Document.ContentEnd).Text;
                        UpdateDocumentStructure(content);
                        
                        // Headers will be formatted when saving the document
                    }
                    catch (Exception)
                    {
                        // If structure update fails, just continue - it's not critical
                    }
                }
            }
        }

        private void FormatExistingHeaders(RichTextBox editor)
        {
            if (editor?.Document == null) return;

            // Find the tab data for this editor
            var tabData = editorTabs.FirstOrDefault(t => t.Editor == editor);
            if (tabData == null) return;

            try
            {
                // Set flag to prevent marking as changed during formatting
                tabData.IsFormattingInProgress = true;
                
                // Store current caret position
                TextPointer caretPosition = editor.CaretPosition;
                
                // Get the plain text content
                string content = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd).Text;
                
                if (string.IsNullOrEmpty(content))
                    return;
                    
                // Use fast line splitting
                string[] lines = content.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                
                // Pre-identify all headers to batch process them
                var headersToFormat = new List<(int lineIndex, int level, string lineText)>();
                
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    string trimmed = line.Trim();

                    // Fast header detection - skip non-headers quickly
                    if (string.IsNullOrWhiteSpace(trimmed) || !trimmed.StartsWith("#"))
                        continue;
                    
                    // Quick validation for valid headers
                    int level = 0;
                    while (level < trimmed.Length && trimmed[level] == '#')
                        level++;
                    
                    if (level >= 1 && level <= 6 && 
                        level < trimmed.Length && 
                        trimmed[level] == ' ')
                    {
                        string headerText = trimmed.Substring(level + 1).Trim();
                        if (!string.IsNullOrEmpty(headerText))
                        {
                            headersToFormat.Add((i, level, line));
                        }
                    }
                }
                
                // Batch format all headers using optimized approach
                FormatHeadersBatch(editor, headersToFormat);
                
                // Restore caret position
                if (caretPosition != null)
                {
                    editor.CaretPosition = caretPosition;
                }
            }
            catch (Exception)
            {
                // Silently ignore formatting errors
            }
            finally
            {
                // Always reset the formatting flag
                tabData.IsFormattingInProgress = false;
            }
        }

        private void FormatHeadersBatch(RichTextBox editor, List<(int lineIndex, int level, string lineText)> headers)
        {
            if (headers.Count == 0) return;

            try
            {
                // Group headers by their text content to handle duplicates efficiently
                var headerGroups = headers.GroupBy(h => h.lineText).ToList();
                
                foreach (var group in headerGroups)
                {
                    string lineText = group.Key;
                    var occurrences = group.ToList();
                    
                    if (occurrences.Count == 1)
                    {
                        // Single occurrence - format directly
                        var header = occurrences[0];
                        FormatSpecificTextFast(editor, lineText, header.level);
                    }
                    else
                    {
                        // Multiple occurrences - format each one
                        for (int i = 0; i < occurrences.Count; i++)
                        {
                            var header = occurrences[i];
                            FormatTextOccurrence(editor, lineText, i, header.level);
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Silently ignore batch formatting errors
            }
        }

        private void FormatSpecificTextFast(RichTextBox editor, string textToFind, int headerLevel)
        {
            try
            {
                // Fast single-occurrence formatting
                TextRange documentRange = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd);
                TextPointer foundStart = documentRange.Start;
                
                while (foundStart != null && foundStart.CompareTo(documentRange.End) < 0)
                {
                    TextPointer foundEnd = foundStart;
                    
                    // Move forward by text length
                    for (int i = 0; i < textToFind.Length && foundEnd != null; i++)
                    {
                        foundEnd = foundEnd.GetNextInsertionPosition(LogicalDirection.Forward);
                    }
                    
                    if (foundEnd != null)
                    {
                        TextRange testRange = new TextRange(foundStart, foundEnd);
                        if (testRange.Text == textToFind)
                        {
                            // Found it! Apply formatting and exit
                            double fontSize = GetHeaderFontSize(headerLevel);
                            testRange.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize);
                            testRange.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                            testRange.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(Colors.White));
                            return;
                        }
                    }
                    
                    foundStart = foundStart.GetNextInsertionPosition(LogicalDirection.Forward);
                }
            }
            catch (Exception)
            {
                // Silently ignore formatting errors
            }
        }

        private void FormatLineAsHeader(RichTextBox editor, int lineIndex, int headerLevel)
        {
            try
            {
                // Get the full document text content
                string content = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd).Text;
                
                if (string.IsNullOrEmpty(content))
                    return;
                
                // Split content into actual lines
                string[] lines = content.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                
                if (lineIndex < 0 || lineIndex >= lines.Length)
                    return;
                
                // Get the target line text
                string targetLine = lines[lineIndex];
                
                // Double-check that this line is actually a header
                string trimmedLine = targetLine.Trim();
                if (string.IsNullOrEmpty(trimmedLine) || !trimmedLine.StartsWith("#"))
                    return;
                
                // Validate header format
                int hashCount = 0;
                while (hashCount < trimmedLine.Length && trimmedLine[hashCount] == '#')
                {
                    hashCount++;
                }
                
                if (hashCount > 6 || hashCount >= trimmedLine.Length || trimmedLine[hashCount] != ' ')
                    return;
                
                string headerText = trimmedLine.Substring(hashCount + 1).Trim();
                if (string.IsNullOrEmpty(headerText))
                    return;
                
                // Find the exact text of this line in the document and format it
                FormatSpecificText(editor, targetLine, headerLevel);
            }
            catch (Exception)
            {
                // Silently ignore formatting errors
            }
        }

        private void FormatSpecificText(RichTextBox editor, string textToFind, int headerLevel)
        {
            try
            {
                // Get all lines to track which occurrence we're looking for
                string content = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd).Text;
                string[] lines = content.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                
                // Find all occurrences of this text and their line positions
                List<int> matchingLineIndexes = new List<int>();
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i] == textToFind)
                    {
                        matchingLineIndexes.Add(i);
                    }
                }
                
                if (matchingLineIndexes.Count == 0)
                    return;
                
                // If there's only one occurrence, format it simply
                if (matchingLineIndexes.Count == 1)
                {
                    FormatTextOccurrence(editor, textToFind, 0, headerLevel);
                    return;
                }
                
                // For multiple occurrences, format all of them
                for (int occurrenceIndex = 0; occurrenceIndex < matchingLineIndexes.Count; occurrenceIndex++)
                {
                    FormatTextOccurrence(editor, textToFind, occurrenceIndex, headerLevel);
                }
            }
            catch (Exception)
            {
                // Silently ignore formatting errors
            }
        }

        private void FormatTextOccurrence(RichTextBox editor, string textToFind, int occurrenceIndex, int headerLevel)
        {
            try
            {
                // Search for the specific occurrence of the text in the document
                TextPointer start = editor.Document.ContentStart;
                TextPointer end = editor.Document.ContentEnd;
                
                // Create a range for the entire document
                TextRange documentRange = new TextRange(start, end);
                
                // Find the text (skip to the correct occurrence)
                TextPointer foundStart = documentRange.Start;
                int foundCount = 0;
                
                while (foundStart != null && foundStart.CompareTo(documentRange.End) < 0)
                {
                    // Try to find the text starting from current position
                    TextPointer foundEnd = foundStart;
                    
                    // Move forward by the length of text we're looking for
                    for (int i = 0; i < textToFind.Length && foundEnd != null; i++)
                    {
                        foundEnd = foundEnd.GetNextInsertionPosition(LogicalDirection.Forward);
                    }
                    
                    if (foundEnd != null)
                    {
                        // Check if this range contains our text
                        TextRange testRange = new TextRange(foundStart, foundEnd);
                        string rangeText = testRange.Text;
                        
                        if (rangeText == textToFind)
                        {
                            // Found a match - check if it's the occurrence we want
                            if (foundCount == occurrenceIndex)
                            {
                                // This is the occurrence we want! Apply formatting
                                double fontSize = GetHeaderFontSize(headerLevel);
                                testRange.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize);
                                testRange.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                                testRange.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(Colors.White));
                                return; // Found and formatted the correct occurrence
                            }
                            
                            // Move past this match to find the next one
                            foundCount++;
                            foundStart = foundEnd;
                            continue;
                        }
                    }
                    
                    // Move to next position
                    foundStart = foundStart.GetNextInsertionPosition(LogicalDirection.Forward);
                }
            }
            catch (Exception)
            {
                // Silently ignore formatting errors
            }
        }

        private void FormatCurrentParagraphIfHeader(RichTextBox editor)
        {
            if (editor?.Document == null) return;

            // Find the tab data for this editor
            var tabData = editorTabs.FirstOrDefault(t => t.Editor == editor);
            if (tabData == null) return;

            try
            {
                // Set flag to prevent marking as changed during formatting
                tabData.IsFormattingInProgress = true;
                
                // Get the current line where the cursor is
                TextPointer caretPosition = editor.CaretPosition;
                
                // Find which line the cursor is on
                string content = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd).Text;
                if (string.IsNullOrEmpty(content))
                    return;
                
                string[] lines = content.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                
                // Find the current line index by searching from caret position
                int currentLineIndex = FindCurrentLineIndex(editor, caretPosition, lines);
                if (currentLineIndex < 0 || currentLineIndex >= lines.Length)
                    return;
                
                // Get the current line text
                string currentLineText = lines[currentLineIndex];
                string trimmedLine = currentLineText.Trim();
                
                if (!string.IsNullOrEmpty(trimmedLine) && trimmedLine.StartsWith("#"))
                {
                    // Determine header level
                    int headerLevel = 0;
                    while (headerLevel < trimmedLine.Length && trimmedLine[headerLevel] == '#')
                    {
                        headerLevel++;
                    }
                    
                    // Only format if it's a valid header (1-6 #s followed by space)
                    if (headerLevel >= 1 && headerLevel <= 6 && 
                        headerLevel < trimmedLine.Length && 
                        trimmedLine[headerLevel] == ' ')
                    {
                        string headerText = trimmedLine.Substring(headerLevel + 1).Trim();
                        if (!string.IsNullOrEmpty(headerText))
                        {
                            // Format only this specific line using the same approach as file loading
                            FormatSpecificText(editor, currentLineText, headerLevel);
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Silently ignore formatting errors
            }
            finally
            {
                // Always reset the formatting flag
                tabData.IsFormattingInProgress = false;
            }
        }

        private int FindCurrentLineIndex(RichTextBox editor, TextPointer caretPosition, string[] lines)
        {
            try
            {
                // Get all text up to the caret position
                TextRange textBeforeCaret = new TextRange(editor.Document.ContentStart, caretPosition);
                string textBefore = textBeforeCaret.Text;
                
                // Count line breaks to determine current line
                if (string.IsNullOrEmpty(textBefore))
                    return 0;
                
                // Split the text before caret to count lines
                string[] linesBefore = textBefore.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                
                // The current line index is the number of lines before the caret - 1
                int lineIndex = linesBefore.Length - 1;
                
                // Ensure we don't exceed the total number of lines
                return Math.Min(Math.Max(0, lineIndex), lines.Length - 1);
            }
            catch
            {
                return 0; // Default to first line if calculation fails
            }
        }

        private double GetHeaderFontSize(int headerLevel)
        {
            switch (headerLevel)
            {
                case 1: return 24.0;
                case 2: return 22.0;
                case 3: return 20.0;
                case 4: return 18.0;
                case 5: return 16.0;
                case 6: return 15.0;
                default: return 14.0;
            }
        }

        private void Editor_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var editor = sender as RichTextBox;
                if (editor != null)
                {
                    // Find the tab data for this editor
                    var tabData = editorTabs.FirstOrDefault(t => t.Editor == editor);
                    if (tabData != null)
                    {
                        // Set flag to prevent marking as changed during formatting
                        tabData.IsFormattingInProgress = true;
                        
                        try
                        {
                            // Set normal formatting properties for the insertion point BEFORE Enter is processed
                            editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, 14.0);
                            editor.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal);
                            editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, new FontFamily("Consolas"));
                            editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(Colors.White));
                            
                            // Also schedule cleanup after Enter is processed
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                try
                                {
                                    // Double-check that insertion formatting is normal
                                    editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, 14.0);
                                    editor.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal);
                                    editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, new FontFamily("Consolas"));
                                    editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(Colors.White));
                                }
                                finally
                                {
                                    tabData.IsFormattingInProgress = false;
                                }
                            }), System.Windows.Threading.DispatcherPriority.Background);
                        }
                        catch
                        {
                            tabData.IsFormattingInProgress = false;
                        }
                    }
                }
            }
        }
        
        private void EditorTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TabControl tabControl = sender as TabControl;
            if (tabControl != null && tabControl.SelectedIndex >= 0 && tabControl.SelectedIndex < editorTabs.Count)
            {
                // Switch to the selected tab
                SwitchToTab(tabControl.SelectedIndex);
                
                // Update document structure for the new active tab
                var currentTab = GetCurrentTab();
                if (currentTab != null && currentTab.Editor != null)
                {
                    try
                    {
                        string content = new TextRange(currentTab.Editor.Document.ContentStart, currentTab.Editor.Document.ContentEnd).Text;
                        UpdateDocumentStructure(content);
                        // Format existing headers when switching tabs (only if Hybrid mode is enabled)
                        if (HybridModeCheckBox.IsChecked == true)
                        {
                            FormatExistingHeaders(currentTab.Editor);
                        }
                    }
                    catch (Exception)
                    {
                        // If structure update fails, just continue - it's not critical
                    }
                }
            }
        }
        
        // Tab management helper methods
        private EditorTabData CreateNewTab(string filePath = null)
        {
            var tabData = new EditorTabData(filePath);
            
            // Subscribe to text changed events for this editor
            tabData.Editor.TextChanged += Editor_TextChanged;
            
            // Subscribe to key events for Enter key handling
            tabData.Editor.PreviewKeyDown += Editor_PreviewKeyDown;
            
            editorTabs.Add(tabData);
            return tabData;
        }
        
        private RichTextBox GetCurrentEditor()
        {
            if (currentTabIndex >= 0 && currentTabIndex < editorTabs.Count)
            {
                return editorTabs[currentTabIndex].Editor;
            }
            
            return null; // No current editor if no tabs are open
        }
        
        private EditorTabData GetCurrentTab()
        {
            if (currentTabIndex >= 0 && currentTabIndex < editorTabs.Count)
            {
                return editorTabs[currentTabIndex];
            }
            
            return null;
        }
        
        private void UpdateTabTitle(EditorTabData tabData)
        {
            // Find the corresponding TabItem in the UI
            for (int i = 0; i < editorTabs.Count; i++)
            {
                if (editorTabs[i] == tabData && i < EditorTabs.Items.Count)
                {
                    var tabItem = EditorTabs.Items[i] as TabItem;
                    if (tabItem != null)
                    {
                        string fileName = string.IsNullOrEmpty(tabData.FilePath) ? 
                            "Untitled" : Path.GetFileName(tabData.FilePath);
                        
                        tabItem.Header = tabData.HasUnsavedChanges ? fileName + " *" : fileName;
                        tabItem.ToolTip = tabData.FilePath ?? "Untitled document";
                    }
                    break;
                }
            }
        }
        
        private void SwitchToTab(int tabIndex)
        {
            if (tabIndex >= 0 && tabIndex < editorTabs.Count)
            {
                currentTabIndex = tabIndex;
                var tabData = editorTabs[tabIndex];
                
                // Update current file path and structure for backward compatibility
                currentFilePath = tabData.FilePath;
                hasUnsavedChanges = tabData.HasUnsavedChanges;
                structureLineMap = tabData.StructureLineMap;
                
                // Update window title
                string fileName = string.IsNullOrEmpty(tabData.FilePath) ? 
                    "Untitled" : Path.GetFileName(tabData.FilePath);
                Title = $"TL Markdown Editor - {fileName}" + (tabData.HasUnsavedChanges ? " *" : "");
            }
        }
        
        private void TabCloseButton_Click(object sender, RoutedEventArgs e)
        {
            // Find which tab this close button belongs to
            Button closeButton = sender as Button;
            if (closeButton != null)
            {
                // Navigate up the visual tree to find the TabItem
                DependencyObject parent = closeButton;
                while (parent != null && !(parent is TabItem))
                {
                    parent = VisualTreeHelper.GetParent(parent);
                }
                
                if (parent is TabItem tabItem)
                {
                    int tabIndex = EditorTabs.Items.IndexOf(tabItem);
                    if (tabIndex >= 0)
                    {
                        CloseTab(tabIndex);
                    }
                }
            }
        }
        
        private bool CloseTab(int tabIndex)
        {
            if (tabIndex < 0 || tabIndex >= editorTabs.Count)
                return false;
                
            var tabData = editorTabs[tabIndex];
            
            // Auto-save tab if it has unsaved changes
            if (tabData.HasUnsavedChanges)
            {
                if (!string.IsNullOrEmpty(tabData.FilePath))
                {
                    // Auto-save existing file
                    try
                    {
                        // Format headers before auto-saving (only if Hybrid mode is enabled)
                        if (HybridModeCheckBox.IsChecked == true)
                        {
                            FormatExistingHeaders(tabData.Editor);
                        }
                        
                        string content = new TextRange(tabData.Editor.Document.ContentStart, tabData.Editor.Document.ContentEnd).Text;
                        File.WriteAllText(tabData.FilePath, content, Encoding.UTF8);
                        tabData.HasUnsavedChanges = false;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error auto-saving file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return false; // Don't close if save failed
                    }
                }
                else
                {
                    // For new unsaved files, show save as dialog since we need a file path
                    SaveFileDialog saveDialog = new SaveFileDialog
                    {
                        Filter = "Markdown files (*.md)|*.md|All files (*.*)|*.*",
                        RestoreDirectory = true,
                        DefaultExt = "md"
                    };
                    
                    if (saveDialog.ShowDialog() == true)
                    {
                        try
                        {
                            // Format headers before saving (only if Hybrid mode is enabled)
                            if (HybridModeCheckBox.IsChecked == true)
                            {
                                FormatExistingHeaders(tabData.Editor);
                            }
                            
                            string content = new TextRange(tabData.Editor.Document.ContentStart, tabData.Editor.Document.ContentEnd).Text;
                            File.WriteAllText(saveDialog.FileName, content, Encoding.UTF8);
                            tabData.HasUnsavedChanges = false;
                            tabData.FilePath = saveDialog.FileName;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            return false; // Don't close if save failed
                        }
                    }
                    else
                    {
                        return false; // User cancelled save as dialog
                    }
                }
            }
            
            // Remove the tab from both UI and data structures
            EditorTabs.Items.RemoveAt(tabIndex);
            editorTabs.RemoveAt(tabIndex);
            
            // Update current tab index
            if (currentTabIndex == tabIndex)
            {
                // Current tab is being closed, switch to adjacent tab
                if (editorTabs.Count > 0)
                {
                    // If closing the last tab, switch to the previous one
                    int newIndex = Math.Min(tabIndex, editorTabs.Count - 1);
                    EditorTabs.SelectedIndex = newIndex;
                    SwitchToTab(newIndex);
                }
                else
                {
                    // No tabs left, reset to empty state
                    currentTabIndex = -1;
                    currentFilePath = null;
                    hasUnsavedChanges = false;
                    structureLineMap.Clear();
                    DocumentStructure.Items.Clear();
                    Title = "TL Markdown Editor";
                }
            }
            else if (currentTabIndex > tabIndex)
            {
                // Adjust current tab index if a tab before it was closed
                currentTabIndex--;
            }
            
            return true;
        }
        
        private void CreateNewEmptyTab()
        {
            // Create new empty tab data
            var tabData = CreateNewTab();
            
            // Create UI tab
            TabItem newTab = new TabItem();
            newTab.Header = "New Document";
            newTab.ToolTip = "Untitled document";
            newTab.Content = tabData.Editor;
            newTab.Name = "InitialTab"; // Mark as initial tab for special handling
            
            // Apply consistent styling to the new editor
            tabData.Editor.Background = new SolidColorBrush(Color.FromRgb(40, 40, 40));
            tabData.Editor.Foreground = Brushes.White;
            tabData.Editor.FontFamily = new FontFamily("Consolas");
            tabData.Editor.FontSize = 14;
            
            EditorTabs.Items.Add(newTab);
            EditorTabs.SelectedIndex = 0;
            
            // Switch to new tab
            currentTabIndex = 0;
            SwitchToTab(0);
            
            // Clear document structure
            DocumentStructure.Items.Clear();
            Title = "TL Markdown Editor";
        }

        // Rich Formatting Event Handlers
        private void UndoButton_Click(object sender, RoutedEventArgs e)
        {
            var currentEditor = GetCurrentEditor();
            if (currentEditor != null && currentEditor.CanUndo)
                currentEditor.Undo();
        }

        private void RedoButton_Click(object sender, RoutedEventArgs e)
        {
            var currentEditor = GetCurrentEditor();
            if (currentEditor != null && currentEditor.CanRedo)
                currentEditor.Redo();
        }

        private void FindReplaceButton_Click(object sender, RoutedEventArgs e)
        {
            var currentEditor = GetCurrentEditor();
            if (currentEditor != null)
            {
                FindReplaceDialog dialog = new FindReplaceDialog(currentEditor);
                dialog.Owner = this;
                dialog.Show();
            }
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
            var currentEditor = GetCurrentEditor();
            if (currentEditor?.Selection != null)
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
                var insertionPosition = currentEditor.CaretPosition;
                var paragraph = insertionPosition.Paragraph;
                
                if (paragraph != null)
                {
                    paragraph.SiblingBlocks.InsertAfter(paragraph, list);
                }
                else
                {
                    currentEditor.Document.Blocks.Add(list);
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
            var currentEditor = GetCurrentEditor();
            if (currentEditor == null) return;
            
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
            var insertionPosition = currentEditor.CaretPosition;
            
            // Insert the table text
            insertionPosition.InsertTextInRun(markdownTable.ToString());
            
            // Move cursor to after the inserted table
            var newPosition = insertionPosition.GetPositionAtOffset(markdownTable.ToString().Length);
            if (newPosition != null)
            {
                currentEditor.CaretPosition = newPosition;
            }
            
            currentEditor.Focus();
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
                    return "📝"; // Markdown file
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