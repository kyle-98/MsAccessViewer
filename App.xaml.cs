using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace MSAccessViewer
{
     /// <summary>
     /// Interaction logic for App.xaml
     /// </summary>
     public partial class App : Application
     {
          public string[] AppArgs { get;  private set; }

          protected override void OnStartup(StartupEventArgs e)
          {
               base.OnStartup(e);
               AppArgs = e.Args;
               this.DispatcherUnhandledException += App_DispatcherUnhandledException;
          }

          private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
          {
               MessageBox.Show(e.Exception.Message);
               e.Handled = true;
          }
     }
}
