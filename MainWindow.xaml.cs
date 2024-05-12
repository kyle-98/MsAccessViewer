﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Data.OleDb;
using MSAccessViewer.Resources;
using System.Windows.Input;
using System.Windows.Controls;
using System.ComponentModel;
using System.Windows.Data;
using System.Windows.Media;

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
          private void BackBtn_Click(object sender, RoutedEventArgs e)
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

          private void GridView_PreviewMouseMove(object sender, MouseEventArgs e)
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
               while(parent != null && parent is not ListView) { parent = VisualTreeHelper.GetParent(parent); }
               return parent as ListView;
          }


          //handle click event on column headers
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
               SortItems(sort_by, direction, FindListViewFromHeader(header_clicked));
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

          //copy selected fields to clipboard
          private void Fieldnames_ListView_KeyDown(object sender, KeyEventArgs e)
          {
               if(e.Key == Key.C && Keyboard.Modifiers == ModifierKeys.Control)
               {
                    dynamic selected_items = Fieldnames_ListView.SelectedItems;
                    int counter = 0;
                    if(selected_items.Count > 0)
                    {
                         string clipboard_data = string.Empty;
                         foreach (var item in selected_items)
                         {
                              if(selected_items.Count == 1) { clipboard_data = item.FieldName; }
                              else if(counter == selected_items.Count - 1) { clipboard_data += item.FieldName; }
                              else { clipboard_data += $"{item.FieldName}\n"; }
                              counter++;
                         }
                         Clipboard.SetText(clipboard_data);
                    }
               }
          }

          private void FindTablenamesBtn_Click(object sender, RoutedEventArgs e)
          {
               FieldnameEntry.ItemsSource = Access.GetAllFieldNames(access_connection);
               FieldnameEntry.SelectedIndex = 0;
               ViewController.SelectedIndex = 2;
          }

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
     }
}
