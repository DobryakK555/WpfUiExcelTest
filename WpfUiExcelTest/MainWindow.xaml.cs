using System;
using System.Collections.Generic;
using System.Security;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace WpfUiExcelTest
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ChooseFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files(.xls)|*.xls; *.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                ExcelService.FileName = openFileDialog.FileName;
                ValidLabel();
            }
            else
            {
                MessageBox.Show("Please, choose correct .xls or .xlsx file");
                InvalidLabel();
            }
        }

        private void InvalidLabel()
        {
            IsFileValid.Content = "Your file is invalid!";
            IsFileValid.Background = Brushes.DarkRed;
            IsFileValid.Foreground = Brushes.White;
            IsFileValid.Visibility = Visibility.Visible;
        }

        private void ValidLabel()
        {
            IsFileValid.Content = "Your file is valid!";
            IsFileValid.Background = Brushes.DarkGreen;
            IsFileValid.Foreground = Brushes.White;
            IsFileValid.Visibility = Visibility.Visible;
        }

        private void Calculate_Click(object sender, RoutedEventArgs e)
        {
            var firstDate = startPeriod.SelectedDate;
            var secondDate = endPeriod.SelectedDate;

            if (ExcelService.CheckDates(firstDate, secondDate))
            {
                var result = ExcelService.GetFilteredSet();

                if (result.Count == 0)
                {
                    MessageBox.Show("No data for choosen dates");
                }
                else
                {
                    ExcelFile.ItemsSource = ExcelService.DoWork(result).DefaultView;
                }
            }
            else
            {
                MessageBox.Show("Enter correct dates. First date must be earlier then second.");
                InvalidLabel();
            }
        }
    }
}
