﻿using System;
using System.Collections.Generic;
using System.Windows;
using System.Data.OleDb;
using MSAccessViewer.Resources;
using System.Windows.Input;
using System.Windows.Controls;
using System.ComponentModel;
using System.Windows.Data;
using System.Windows.Media;
using System.Data;
using System.Diagnostics;

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
          List<string> access_tablenames = new();
          string[] sys_args;

          public MainWindow()
          {
               InitializeComponent();

               sys_args = ((App)Application.Current).AppArgs;
               if (sys_args != null) { access_connection = Access.Connect(sys_args[0]); }
               else { throw new Exception("Could not connect to the access file. It is either not an access file or there was failure upon connecting to it."); }

               access_tablenames = Access.GetAccessTableNames(access_connection);
               TablenamesListbox.ItemsSource = access_tablenames;

               AccessFilePathTextbox.Text = sys_args[0];
          }

          /*
           * ----------------------
           * General helper methods
           * ----------------------
          */

          private void PopulateFieldNamesListView()
          {
               if (TablenamesListbox != null)
               {
                    if (TablenamesListbox.SelectedItem != null && TablenamesListbox.SelectedItem.ToString() != string.Empty)
                    {
                         FieldNames_DataGrid.ItemsSource = null;
                         if (Table_ColumnCombobox.SelectedIndex == 0)
                         {
                              DataTable access_dt = Access.GetDatatable(access_connection, TablenamesListbox.SelectedItem.ToString());
                              Access.CorrectDataTypes(access_dt, access_connection);
                              FieldNames_DataGrid.ItemsSource = access_dt.DefaultView;
                         }
                         else
                         {
                              FieldNames_DataGrid.ItemsSource = Access.GetFieldNames(access_connection, TablenamesListbox.SelectedItem.ToString());
                         }
                    }
                    else { if (FieldNames_DataGrid != null) { FieldNames_DataGrid.ItemsSource = null; } }
               }
               else { return; }
          }

          //sorting columns based on the header clicked
          private void SortItems(string sort_by, ListSortDirection direction, ListView list_view)
          {
               ICollectionView data_view = CollectionViewSource.GetDefaultView(list_view.ItemsSource);
               if (data_view == null) { return; }
               data_view.SortDescriptions.Clear();
               SortDescription sort_desc = new(sort_by, direction);
               data_view.SortDescriptions.Add(sort_desc);
               data_view.Refresh();
          }

          //get the listview where the column the user clicked exists in
          private ListView FindListViewFromHeader(GridViewColumnHeader header)
          {
               DependencyObject parent = VisualTreeHelper.GetParent(header);
               while (parent != null && parent is not ListView) { parent = VisualTreeHelper.GetParent(parent); }
               return parent as ListView;
          }


          /*
           * --------------------------
           * General user event methods
           * --------------------------
          */

          //Take user back to the main page of the application from the ViewController.SelectedIndex = 1
          private void BackBtn_Click(object sender, RoutedEventArgs e)
          {
               ViewController.SelectedIndex = 0;
          }

          private void GridView_PreviewMouseMove(object sender, MouseEventArgs e)
          {
               e.Handled = true;
          }

          //handle click event on column headers
          private void GridViewColumnHeader_Click(object sender, RoutedEventArgs e)
          {
               var header_clicked = e.OriginalSource as GridViewColumnHeader;
               ListSortDirection direction;
               if (header_clicked == null) { return; }
               if (header_clicked.Role == GridViewColumnHeaderRole.Padding) { return; }

               if (header_clicked != _last_header_clicked)
               {
                    direction = ListSortDirection.Ascending;
               }
               else
               {
                    if (_last_direction == ListSortDirection.Ascending) { direction = ListSortDirection.Descending; }
                    else { direction = ListSortDirection.Ascending; }
               }
               var column_binding = header_clicked.Column.DisplayMemberBinding as Binding;
               var sort_by = column_binding?.Path.Path ?? header_clicked.Column.Header as string;
               SortItems(sort_by, direction, FindListViewFromHeader(header_clicked));
               if (direction == ListSortDirection.Ascending)
               {
                    header_clicked.Column.HeaderTemplate = Resources["HeaderTemplateArrowUp"] as DataTemplate;
               }
               else
               {
                    header_clicked.Column.HeaderTemplate = Resources["HeaderTemplateArrowDown"] as DataTemplate;
               }

               if (_last_header_clicked != null && _last_header_clicked != header_clicked)
               {
                    _last_header_clicked.Column.HeaderTemplate = null;
               }
               _last_header_clicked = header_clicked;
               _last_direction = direction;
          }

          
          private void CheckTabIndex_TableData(object sender, RoutedEventArgs e)
          {
               //make sure user is on the correct tab and has a table selected in the listbox
               if (ViewController.SelectedIndex == 1 && TablenamesListbox.SelectedItem != null)
               {
                    //content selected in combobox
                    if (Table_ColumnCombobox.SelectedIndex == 0)
                    {
                         UpdateData_Header.IsEnabled = true;
                         ExportTableData_Header.IsEnabled = true;
                    }
                    else
                    {
                         UpdateData_Header.IsEnabled = false;
                         ExportTableData_Header.IsEnabled = false;
                    }
               }
               else
               {
                    UpdateData_Header.IsEnabled = false;
                    ExportTableData_Header.IsEnabled = false;
               }
          }


          /*
           * --------------------------------------------------------------
           * MainPage user event methods (ViewController.SelectedIndex = 0
           * --------------------------------------------------------------
          */

          //take user to find fields button back
          private void FindFieldnamesBtn_Click(object sender, RoutedEventArgs e)
          {
               access_tablenames = Access.GetAccessTableNames(access_connection);
               ViewController.SelectedIndex = 1;
          }

          private void FindTablenamesBtn_Click(object sender, RoutedEventArgs e)
          {
               FieldnameEntry.ItemsSource = Access.GetAllFieldNames(access_connection);
               FieldnameEntry.SelectedIndex = 0;
               ViewController.SelectedIndex = 2;
          }

          /// <summary>
          /// Open the file directory where the current access file is stored in
          /// </summary>
          /// <param name="sender">Object being passed into the method by a user event</param>
          /// <param name="e">Event that is triggered based on a user interaction</param>
          private void AccessFilePathLbl_DoubleClick(object sender, RoutedEventArgs e)
          {
               string[] path_arr = sys_args[0].Split('\\');
               Array.Resize(ref path_arr, path_arr.Length - 1);
               Process.Start("explorer.exe", $"{string.Join('\\', path_arr)}");
          }


          /*
           * ---------------------------------------------------------------
           * FieldNames user event methods (ViewController.SelectedIndex = 1
           * ---------------------------------------------------------------
          */

          private void Table_ColumnCombobox_SelectionChanged(object sender, RoutedEventArgs e)
          {
               PopulateFieldNamesListView();
          }

          private void SearchTablenameBox_TextChanged(object sender, TextChangedEventArgs e)
          {
               string search_text = SearchTablenameInput.Text.ToLower();
               if (string.IsNullOrEmpty(SearchTablenameInput.Text.ToLower()))
               {
                    TablenamesListbox.ItemsSource = access_tablenames;
               }
               else
               {
                    List<string> filtered_items = new();
                    foreach (var item in TablenamesListbox.Items)
                    {
                         if (item.ToString().ToLower().Contains(search_text))
                         {
                              filtered_items.Add(item.ToString());
                         }
                    }
                    TablenamesListbox.ItemsSource = filtered_items;
               }
          }

          private void TablenamesListbox_SelectionChanged(object sender, SelectionChangedEventArgs e)
          {
               if (TablenamesListbox.SelectedItems.Count == 0) { FieldNames_DataGrid.ItemsSource = null; }
               else { PopulateFieldNamesListView(); }
          }

          private void ExportTableDataHeader_Click(object sender, RoutedEventArgs e)
          {
               Access.ExportTableData(access_connection, TablenamesListbox.SelectedItem.ToString(), FieldNames_DataGrid);
          }


          /*
           * ---------------------------------------------------------------
           * TableNames user event methods (ViewController.SelectedIndex = 2
           * ---------------------------------------------------------------
          */

          private void FieldnameEntry_SelectionChanged(object sender, RoutedEventArgs e)
          {
               if(FieldnameEntry.SelectedItem != null && FieldnameEntry.SelectedItem.ToString() != string.Empty) 
               {
                    Tablenames_ListView.ItemsSource = Access.GetTablenameViaField(access_connection, FieldnameEntry.SelectedItem.ToString());

                    foreach (GridViewColumn column in TableNames_GridView.Columns)
                    {
                         if (double.IsNaN(column.Width)) { column.Width = column.ActualWidth; }
                         column.Width = double.NaN;
                    }
               }
               else { Tablenames_ListView.ItemsSource = null; }
          }

          private void UpdateAccessDataHeader_Click(object sender, RoutedEventArgs e)
          {
               if(TablenamesListbox.SelectedItem != null && Table_ColumnCombobox.SelectedIndex == 0)
               {
                    Access.UpdateTable(access_connection, TablenamesListbox.SelectedItem.ToString(), ((DataView)FieldNames_DataGrid.ItemsSource).Table);
               }
          }



     }
}
