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
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;

namespace MSAccessViewer
{
     /// <summary>
     /// Interaction logic for MainWindow.xaml
     /// </summary>
     public partial class MainWindow : Window
     {

          OleDbConnection? access_connection;
          GridViewColumnHeader _last_header_clicked = null;
          ListSortDirection _last_direction = ListSortDirection.Ascending;

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

          private void SortItems(string sort_by, ListSortDirection direction, ListView list_view)
          {
               ICollectionView data_view = CollectionViewSource.GetDefaultView(list_view.ItemsSource);

               data_view.SortDescriptions.Clear();
               SortDescription sort_desc = new(sort_by, direction);
               data_view.SortDescriptions.Add(sort_desc);
               data_view.Refresh();
          }

          private void GridViewColumnHeader_Click(object sender, RoutedEventArgs e)
          {
               var header_clicked = e.OriginalSource as GridViewColumnHeader;
               ListSortDirection direction;
               if(header_clicked == null) { return; }
               if(header_clicked.Role == GridViewColumnHeaderRole.Padding) { return; }

               if(header_clicked != _last_header_clicked)
               {
                    direction = ListSortDirection.Ascending;
               }
               else
               {
                    if(_last_direction == ListSortDirection.Ascending) { direction = ListSortDirection.Descending; }
                    else { direction = ListSortDirection.Ascending; }
               }
               var column_binding = header_clicked.Column.DisplayMemberBinding as Binding;
               var sort_by = column_binding?.Path.Path ?? header_clicked.Column.Header as string;
               SortItems(sort_by, direction, Fieldnames_ListView);
               if(direction == ListSortDirection.Ascending)
               {
                    header_clicked.Column.HeaderTemplate = Resources["HeaderTemplateArrowUp"] as DataTemplate;
               }
               else
               {
                    header_clicked.Column.HeaderTemplate = Resources["HeaderTemplateArrowDown"] as DataTemplate;
               }

               if(_last_header_clicked != null && _last_header_clicked != header_clicked)
               {
                    _last_header_clicked.Column.HeaderTemplate = null;
               }
               _last_header_clicked = header_clicked;
               _last_direction = direction;

              
          }
     }
}
