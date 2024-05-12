using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace MSAccessViewer.Resources
{
     public class AccessField : IComparable<AccessField>
     {
          public string FieldName { get; set; }
          public string DataType { get; set; }
          public string OrdinalPosition { get; set; }
          public string IsNullable { get; set; }
          public string Description { get; set; }


          public AccessField(string fieldname, string datatype, string ordposition, string isnullable, string description)
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

          public static void CloseConnection(OleDbConnection access_connection) { access_connection.Close(); }

          public static ObservableCollection<AccessField> GetFieldNames(OleDbConnection access_connection, string tablename)
          {
               ObservableCollection<AccessField> fields = new();
               DataTable schema = access_connection.GetSchema("Columns", new[] { null, null, tablename });
               if (schema != null)
               {
                    foreach (DataRow row in schema.Rows)
                    {
                         if (row != null) { fields.Add(new AccessField(row["COLUMN_NAME"].ToString(), GetDataTypeName(Convert.ToInt32(row["DATA_TYPE"])), row["ORDINAL_POSITION"].ToString(), row["IS_NULLABLE"].ToString(), row["DESCRIPTION"].ToString() == string.Empty ? "N/A" : row["DESCRIPTION"].ToString())); }
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

          
     }
}
