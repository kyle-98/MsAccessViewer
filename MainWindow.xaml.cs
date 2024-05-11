using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Data.OleDb;
using MSAccessViewer.Resources;
using System.Diagnostics;
using System.Windows.Input;
using System.Windows.Controls;

namespace MSAccessViewer
{
     /// <summary>
     /// Interaction logic for MainWindow.xaml
     /// </summary>
     public partial class MainWindow : Window
     {

          OleDbConnection? access_connection;

          public MainWindow()
          {
               InitializeComponent();

               string[] sys_args = ((App)Application.Current).AppArgs;
               if (sys_args != null) { access_connection = Access.Connect(sys_args[0]); }
               else { throw new Exception("Could not connect to the access file. It is either not an access file or there was failure upon connecting to it."); }
               
               
               
          }

          //Take user back to the main page of the application from the ViewController.SelectedIndex = 1
          private void FindFieldsBackBtn_Click(object sender, RoutedEventArgs e)
          {
               ViewController.SelectedIndex = 0;
          }

          //take user to find fields button back (this button is on the main page)
          private void FindFieldnamesBtn_Click(object sender, RoutedEventArgs e)
          {
               TablenamesCombobox.ItemsSource = Access.GetAccessTableNames(access_connection);
               TablenamesCombobox.SelectedIndex = 0;
               ViewController.SelectedIndex = 1;
          }

          private void Fieldnames_GridView_PrviewMouseMove(object sender, MouseEventArgs e)
          {
               e.Handled = true;
          }

          private void TablenamesCombobox_SelectionChanged(object sender, RoutedEventArgs e)
          {
               Fieldnames_ListView.ItemsSource = Access.GetFieldNames(access_connection, TablenamesCombobox.SelectedItem.ToString());
               foreach(GridViewColumn column in FieldNames_GridView.Columns)
               {
                    if (double.IsNaN(column.Width)) { column.Width = column.ActualWidth; }
                    column.Width = double.NaN;
               }
          }
     }
}
