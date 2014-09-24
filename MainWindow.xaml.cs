using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using OfficeOpenXml;

namespace Excel_Text_Modifier
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.AllowDrop = true;
            this.DragEnter += MainWindow_DragEnter;
            this.Drop += MainWindow_Drop;
            InitializeComponent();
        }

        #region DragDropControl

        private void MainWindow_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void MainWindow_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] filePaths = (string[])(e.Data.GetData(DataFormats.FileDrop));
                foreach (string fileLoc in filePaths)
                {
                    
                }
            }
        }

        #endregion

        private void FileOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.DefaultExt = ".xlsx";
            openDialog.Filter = "Microsoft Excel Worksheet (.xlsx)|*.xlsx";
            Nullable<bool> result = openDialog.ShowDialog();
            if (result == true)
            {
                string filename = openDialog.FileName;
                FileNameTextBox.Text = filename;
            }
        }

        private void CloseProg_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            FileInfo File = new FileInfo(FileNameTextBox.Text);
            using (ExcelPackage package = new ExcelPackage(File))
            {   
                int colume = 3;
                int row = 4;
                object curCell = "";
                
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                // Goals
                // Read each Cell in Colume C
                // Compare first 3 char from string to list
                // If in list add "DT" or "LT"
                while (true)
                {
                    curCell = worksheet.Cells[row, colume].Value;                    
                    if (curCell == null)
                        break;
                    worksheet.Cells[row, colume].Value = addPrefix(curCell.ToString());                    
                    row++;
                }
                package.Save(); // Saves modifications
            }
            MessageBox.Show("Done.");
        }

        private string addPrefix(string cell)
        {
            string workingPrefix = cell.Substring(0, 3);
            if (PrefixLists.DT.Contains(workingPrefix))
            {
                 cell = string.Format("DT{0}",cell);
            }else if (PrefixLists.LT.Contains(workingPrefix))
            {
                cell = string.Format("LT{0}",cell);
            }
            return cell;
        }

        private void FileList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
