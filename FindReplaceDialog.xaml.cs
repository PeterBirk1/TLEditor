using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace TLEditor
{
    public partial class FindReplaceDialog : Window
    {
        private RichTextBox targetEditor;
        private TextPointer lastFoundPosition;

        public FindReplaceDialog(RichTextBox editor)
        {
            InitializeComponent();
            targetEditor = editor;
        }

        private void FindNextButton_Click(object sender, RoutedEventArgs e)
        {
            FindNext();
        }

        private void ReplaceButton_Click(object sender, RoutedEventArgs e)
        {
            Replace();
        }

        private void ReplaceAllButton_Click(object sender, RoutedEventArgs e)
        {
            ReplaceAll();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private bool FindNext()
        {
            if (string.IsNullOrEmpty(FindTextBox.Text))
                return false;

            string searchText = FindTextBox.Text;
            StringComparison comparison = MatchCaseCheckBox.IsChecked == true ? 
                StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            TextPointer startPosition = lastFoundPosition ?? targetEditor.Document.ContentStart;
            TextPointer position = FindTextInRange(startPosition, targetEditor.Document.ContentEnd, searchText, comparison);

            if (position == null && lastFoundPosition != null)
            {
                position = FindTextInRange(targetEditor.Document.ContentStart, targetEditor.Document.ContentEnd, searchText, comparison);
                lastFoundPosition = null;
            }

            if (position != null)
            {
                lastFoundPosition = position.GetPositionAtOffset(searchText.Length);
                targetEditor.Selection.Select(position, lastFoundPosition);
                targetEditor.Focus();
                return true;
            }

            MessageBox.Show("Text not found.", "Find", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }

        private void Replace()
        {
            if (targetEditor.Selection.Text == FindTextBox.Text || 
                (MatchCaseCheckBox.IsChecked == false && 
                 string.Equals(targetEditor.Selection.Text, FindTextBox.Text, StringComparison.OrdinalIgnoreCase)))
            {
                targetEditor.Selection.Text = ReplaceTextBox.Text;
                FindNext();
            }
            else
            {
                FindNext();
            }
        }

        private void ReplaceAll()
        {
            if (string.IsNullOrEmpty(FindTextBox.Text))
                return;

            int count = 0;
            lastFoundPosition = null;
            
            TextRange documentRange = new TextRange(targetEditor.Document.ContentStart, targetEditor.Document.ContentEnd);
            string documentText = documentRange.Text;
            
            StringComparison comparison = MatchCaseCheckBox.IsChecked == true ? 
                StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            
            string newText = documentText.Replace(FindTextBox.Text, ReplaceTextBox.Text, comparison);
            
            if (newText != documentText)
            {
                count = (documentText.Length - newText.Length) / (FindTextBox.Text.Length - ReplaceTextBox.Text.Length);
                documentRange.Text = newText;
            }

            MessageBox.Show($"Replaced {count} occurrences.", "Replace All", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private TextPointer FindTextInRange(TextPointer start, TextPointer end, string searchText, StringComparison comparison)
        {
            TextPointer current = start;
            
            while (current != null && current.CompareTo(end) < 0)
            {
                if (current.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                {
                    string textRun = current.GetTextInRun(LogicalDirection.Forward);
                    int index = textRun.IndexOf(searchText, comparison);
                    
                    if (index >= 0)
                    {
                        return current.GetPositionAtOffset(index);
                    }
                }
                
                current = current.GetNextContextPosition(LogicalDirection.Forward);
            }
            
            return null;
        }
    }

    public static class StringExtensions
    {
        public static string Replace(this string original, string oldValue, string newValue, StringComparison comparison)
        {
            string result = original;
            int index = 0;
            
            while ((index = result.IndexOf(oldValue, index, comparison)) >= 0)
            {
                result = result.Substring(0, index) + newValue + result.Substring(index + oldValue.Length);
                index += newValue.Length;
            }
            
            return result;
        }
    }
}