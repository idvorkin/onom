using OneNoteMenu;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace OneNotePieMenu
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // This should be in a seperate application - but overloading for now, since I'd refactor this anyway before long. 
            var isGenerateDayAndWeekTemplate = e.Args.Contains("template");
            {
                GenerateDayAndWeekTemplate();
            }
        }

        private static void GenerateDayAndWeekTemplate()
        {
            var capabilities = new AllOneNoteCapabilities();

            // Goto X page will create the page if it does not exist.
            capabilities.DailyPages.GotoThisWeekPage(); // I think this might break on Sunday night
            capabilities.DailyPages.GotoTodayPage();
            System.Windows.Application.Current.Shutdown();
        }
    }
}
