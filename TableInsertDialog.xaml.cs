using System;
using System.Windows;

namespace TLEditor
{
    public partial class TableInsertDialog : Window
    {
        public int Rows { get; private set; }
        public int Columns { get; private set; }

        public TableInsertDialog()
        {
            InitializeComponent();
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(RowsTextBox.Text, out int rows) && rows > 0 && rows <= 20 &&
                int.TryParse(ColumnsTextBox.Text, out int columns) && columns > 0 && columns <= 10)
            {
                Rows = rows;
                Columns = columns;
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show("Please enter valid numbers for rows (1-20) and columns (1-10).", "Invalid Input", 
                               MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}