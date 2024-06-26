﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace MSAccessViewer.Resources
{
     public class AccessField : IComparable<AccessField>
     {
          public string FieldName { get; set; }
          public string DataType { get; set; }
          public int OrdinalPosition { get; set; }
          public string IsNullable { get; set; }
          public string Description { get; set; }


          public AccessField(string fieldname, string datatype, int ordposition, string isnullable, string description)
          {
               FieldName = fieldname;
               DataType = datatype;
               OrdinalPosition = ordposition;
               IsNullable = isnullable;
               Description = description;
          }

          public int CompareTo(AccessField other)
          {
               return string.Compare(FieldName, other.FieldName, StringComparison.Ordinal);
          }
     }

     public class AccessTableField : IComparable<AccessTableField>
     {
          public string TableName { get; set; }
          public int OrdinalPosition { get; set; }
          public string DataType { get; set; }
          public string IsNullable { get; set; }
          public string Description { get; set; }

          public AccessTableField(string tablename, int ordpositon, string datatype, string isnullable, string desc)
          {
               TableName = tablename;
               OrdinalPosition = ordpositon;
               DataType = datatype;
               IsNullable = isnullable;
               Description = desc;
          }
          public int CompareTo(AccessTableField other)
          {
               return string.Compare(TableName, other.TableName, StringComparison.Ordinal);
          }

     }



     public static class Access
     {
          public static OleDbConnection? Connect(string filepath)
          {
               const string FILE_PASSWORD = "patrick_star";
               string connection_string = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filepath};Jet OLEDB:Database Password={FILE_PASSWORD};Mode=12;";
               OleDbConnection access_connection = new(connection_string);
               try
               {
                    access_connection.Open();
                    return access_connection;
               }
               catch (Exception ex)
               {
                    throw new Exception($"Could not open access database file: \n\n{ex.Message}");
               }
          }

          /// <summary>
          /// Convert the numeric datatypes retieved from access into readable data types to present to the user
          /// </summary>
          /// <param name="dataType">Numeric data type retrieved from the access column</param>
          /// <returns>String that is a readable data type</returns>
          private static string GetDataTypeName(int dataType)
          {
               switch (dataType)
               {
                    case 2:
                         return "Short Integer";
                    case 3:
                         return "Long Integer";
                    case 4:
                         return "Single";
                    case 5:
                         return "Double";
                    case 6:
                         return "Currency";
                    case 7:
                         return "Date/Time";
                    case 11:
                         return "Yes/No";
                    case 17:
                         return "Byte";
                    case 72:
                         return "GUID";
                    case 128:
                         return "Binary";
                    case 130:
                         return "Text";
                    case 131:
                         return "Long Text";
                    default:
                         return "Unknown";
               }
          }

          /// <summary>
          /// Close a connection to the access file passed as a system parameter
          /// </summary>
          /// <param name="access_connection"><c>OleDbConnection</c> to an access file</param>
          public static void CloseConnection(OleDbConnection access_connection) { access_connection.Close(); }


          public static ObservableCollection<AccessField> GetFieldNames(OleDbConnection access_connection, string tablename)
          {
               ObservableCollection<AccessField> fields = new();
               DataTable schema = access_connection.GetSchema("Columns", new[] { null, null, tablename });
               if (schema != null)
               {
                    foreach (DataRow row in schema.Rows)
                    {
                         if (row != null) 
                         { 
                              fields.Add(
                                   new AccessField(
                                        row["COLUMN_NAME"].ToString(), 
                                        GetDataTypeName(Convert.ToInt32(row["DATA_TYPE"])), 
                                        Convert.ToInt32(row["ORDINAL_POSITION"]), 
                                        row["IS_NULLABLE"].ToString(), 
                                        row["DESCRIPTION"].ToString() == string.Empty ? "N/A" : row["DESCRIPTION"].ToString()
                                   )
                              ); 
                         }
                    }
               }
               return fields;
          }

          public static List<string> GetAccessTableNames(OleDbConnection access_connection)
          {
               List<string> tablenames = new();
               DataTable schema = access_connection.GetSchema("Tables");
               DataRow[] table_rows = schema.Select("TABLE_TYPE = 'TABLE'");
               foreach (DataRow row in table_rows) { tablenames.Add(row["TABLE_NAME"].ToString()); }
               return tablenames;
          }

          public static List<string> GetAllFieldNames(OleDbConnection access_connection)
          {
               List<string> field_names = new();
               DataTable schema = access_connection.GetSchema("Tables");
               DataRow[] table_rows = schema.Select("TABLE_TYPE = 'TABLE'");
               foreach (DataRow row in table_rows)
               {
                    if(row != null) 
                    {
                         DataTable cols_schema = access_connection.GetSchema("Columns", new[] { null, null, row["TABLE_NAME"].ToString() });
                         foreach(DataRow cs_rows in cols_schema.Rows)
                         {
                              if (!field_names.Contains(cs_rows["COLUMN_NAME"].ToString()))
                              {
                                   field_names.Add(cs_rows["COLUMN_NAME"].ToString());
                              }
                         }
                    }
               }
               return field_names;
          }

          public static ObservableCollection<AccessTableField> GetTablenameViaField(OleDbConnection access_connection, string field_name)
          {
               ObservableCollection<AccessTableField> table_names = new();
               DataTable schema = access_connection.GetSchema("Tables");
               DataRow[] table_rows = schema.Select("TABLE_TYPE = 'TABLE'");
               foreach (DataRow row in table_rows)
               {
                    if (row != null)
                    {
                         DataTable col_schema = access_connection.GetSchema("Columns", new[] { null, null, row["TABLE_NAME"].ToString(), field_name });
                         if(col_schema.Rows.Count == 1) 
                         {
                              foreach(DataRow col_row in col_schema.Rows)
                              {
                                   table_names.Add(
                                        new AccessTableField(
                                             row["TABLE_NAME"].ToString(),
                                             Convert.ToInt32(col_row["ORDINAL_POSITION"]),
                                             GetDataTypeName(Convert.ToInt32(col_row["DATA_TYPE"])),
                                             col_row["IS_NULLABLE"].ToString(),
                                             col_row["DESCRIPTION"].ToString() == string.Empty ? "N/A" : col_row["DESCRIPTION"].ToString()
                                        )
                                   );
                              }
                         }
                    }
               }
               return table_names;
          }

          public static DataTable GetDatatable(OleDbConnection access_connection, string tablename) 
          {
               OleDbDataAdapter adapter = new($"select * from [{tablename}]", access_connection);
               DataTable dt = new();
               adapter.Fill(dt);
               return dt;
          }

          private static void ConvertColumnType(DataTable dt, string column_name, Type new_type)
          {
               DataColumn new_col = new($"{column_name}_new", new_type);
               dt.Columns.Add(new_col);
               
               foreach(DataRow row in dt.Rows) 
               {
                    if (row[column_name] != DBNull.Value)
                    {
                         row[new_col] = Convert.ChangeType(row[column_name], new_type);
                    }
               }
               dt.Columns.Remove(column_name);
               new_col.ColumnName = column_name;
          }

          public static void CorrectDataTypes(DataTable dt, OleDbConnection access_connection)
          {
               DataTable schema = access_connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, dt.TableName, null });
               foreach(DataColumn col in dt.Columns)
               {
                    foreach(DataRow row in schema.Rows)
                    {
                         if (row["COLUMN_NAME"].ToString() == col.ColumnName)
                         {
                              string data_type = row["DATA_TYPE"].ToString();
                              switch(data_type)
                              {
                                   case "3": // Long Integer
                                        ConvertColumnType(dt, col.ColumnName, typeof(int));
                                        break;
                                   case "6": // Currency
                                        ConvertColumnType(dt, col.ColumnName, typeof(decimal));
                                        break;
                                   case "7": // Date/Time   
                                        ConvertColumnType(dt, col.ColumnName, typeof(DateTime));
                                        break;
                                   case "4": // Single
                                        ConvertColumnType(dt, col.ColumnName, typeof(float));
                                        break;
                                   case "5": // Double
                                        ConvertColumnType(dt, col.ColumnName, typeof(double));
                                        break;
                                   case "11": // Boolean
                                        ConvertColumnType(dt, col.ColumnName, typeof(bool));
                                        break;
                                   case "17": // Byte
                                        ConvertColumnType(dt, col.ColumnName, typeof(byte));
                                        break;
                                   case "130": // Text
                                        ConvertColumnType(dt, col.ColumnName, typeof(string));
                                        break;
                                   
                                   default:
                                        break;
                              }
                         }
                    }
               }
          }

          public static void UpdateTable(OleDbConnection access_connection, string tablename, DataTable datagrid_dt)
          {
               try
               {
                    OleDbDataAdapter adapter = new();
                    OleDbCommandBuilder cmd_builder = new(adapter);
                    cmd_builder.DataAdapter = adapter;
                    cmd_builder.QuotePrefix = "[";
                    cmd_builder.QuoteSuffix = "]";
                    adapter.SelectCommand = new OleDbCommand($"select * from {tablename}", access_connection);
                    adapter.Update(datagrid_dt);
                    MessageBox.Show($"Successfully updated: {tablename}", "Update Success", MessageBoxButton.OK, MessageBoxImage.Information);
               }
               catch (Exception ex) { MessageBox.Show($"Error when updating access table:\n{ex.Message}", "Update error", MessageBoxButton.OK, MessageBoxImage.Error); }
          }


          private static bool DatagridEqualAccessTable(DataTable access_dt, DataTable datagrid_dt)
          {
               if(access_dt.Rows.Count != datagrid_dt.Rows.Count) { return false; }
               if (access_dt.Columns.Count != datagrid_dt.Columns.Count) { return false; }

               for (int i = 0; i < access_dt.Columns.Count; i++)
               {
                    if (
                         access_dt.Columns[i].ColumnName != datagrid_dt.Columns[i].ColumnName || 
                         access_dt.Columns[i].DataType != datagrid_dt.Columns[i].DataType
                    ) { return false; }
               }

               for (int row = 0; row < access_dt.Rows.Count; row++)
               {
                    for(int col = 0; col < access_dt.Columns.Count; col++)
                    {
                         if (!Equals(access_dt.Rows[row][col], datagrid_dt.Rows[row][col])) { return false; }
                    }
               }
               return true;
          }


          private static void ExportToCSV(DataTable export_table, string filepath)
          {
               StringBuilder sb = new();
               string[] column_names = export_table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
               sb.AppendLine(string.Join(',', column_names));

               foreach(DataRow row in export_table.Rows)
               {
                    string[] fields = row.ItemArray.Select(field => field.ToString()).ToArray();
                    sb.AppendLine(string.Join(',', fields));
               }

               File.WriteAllText(filepath, sb.ToString());
          }



          private static void ShowSaveDialog(DataTable export_table)
          {
               SaveFileDialog save_file_dialog = new SaveFileDialog {
                    Filter = "CSV Files (*.csv|*.csv)",
                    Title = "Export to CSV",
                    FileName = "Hentai.csv"
               };

               if(save_file_dialog.ShowDialog() == true)
               {
                    string filepath = save_file_dialog.FileName;
                    try 
                    { 
                         ExportToCSV(export_table, filepath); 
                         MessageBoxResult result = MessageBox.Show("Successfully exported table.\n\nWould you like to open the csv?", "Export to CSV", MessageBoxButton.YesNo, MessageBoxImage.Information);
                         if(result == MessageBoxResult.Yes) { Process.Start("explorer.exe", filepath); }
                    }
                    catch(Exception ex) { MessageBox.Show($"Failed to export csv:\n\n{ex.Message}"); }
               }
          }


          public static void ExportTableData(OleDbConnection access_connection, string tablename, DataGrid dg)
          {
               DataTable access_dt = new();
               using(OleDbCommand cmd = new($"select * from [{tablename}]", access_connection))
               {
                    OleDbDataAdapter adapter = new(cmd);
                    adapter.Fill(access_dt);
               }

               DataTable dg_dt = ((DataView)dg.ItemsSource).Table;

               if (DatagridEqualAccessTable(access_dt, dg_dt))
               {
                    ShowSaveDialog(access_dt);
               }
               else
               {
                    MessageBoxResult result = MessageBox.Show(
                         "The datagrid and the table stored in access do not match.\n\n Would you like to commit your changes in the datagrid to the access file before exporting?\n\nSelecting Yes will update then export\nSelecting No will export without updating\nSelecting Cancel will abort the operation", 
                         "Export to CSV Warning", 
                         MessageBoxButton.YesNoCancel, 
                         MessageBoxImage.Warning
                    );

                    if(result == MessageBoxResult.Yes) { UpdateTable(access_connection, tablename, dg_dt); }

                    if(result == MessageBoxResult.Cancel) { return; }

                    if(result == MessageBoxResult.Yes || result == MessageBoxResult.No)
                    {
                         ShowSaveDialog(access_dt);
                    }
                    

               }

          }

     }
}
