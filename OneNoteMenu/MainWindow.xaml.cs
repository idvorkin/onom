using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

// A collection of onenote capabilities with a button for each.
using OnenoteCapabilities;
using OneNoteMenu.Properties;
using Settings = OnenoteCapabilities.Settings;

namespace OneNoteMenu
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        DailyPages dailyPages = new DailyPages(new Settings());
        EraseEmpty erase = new EraseEmpty();
        public MainWindow()
        {
            InitializeComponent();
            CreateButtons();
        }
        // UX Helpers - These should move to an alternate assembly
        Button CreateButton(string Content, Action OnClick)
        {
            var button = new Button() { Content = Content };
            button.Click += (o, e) => { OnClick(); };

            button.FontSize = 20;
            return button;
        }

        void CreateButtons()
        {
            var buttons = new[]{
                CreateButton("Goto _Today", dailyPages.GotoTodayPage),
                CreateButton("Goto This _Week", dailyPages.GotoThisWeekPage),
                CreateButton("Erase Empty Section", erase.DeleteEmptySections),
            }.ToList();

            buttons.ForEach((b) => this.StackPanel.Children.Add(b));
        }
    }

}
