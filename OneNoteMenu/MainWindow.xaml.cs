using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using OneNoteObjectModel;
using Settings = OnenoteCapabilities.Settings;

namespace OneNoteMenu
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static readonly OneNoteApp ona = new OneNoteApp();
        readonly EraseEmpty erase = new EraseEmpty();
        readonly DailyPages dailyPages = new DailyPages(ona, new Settings());
        private static string[] _people = "SeanSe;AlaksS;MaSudame;AmmonL;LarryS;IgorD;ToriS".Split(';');
        private ObservableCollection<string> _observablePeople = new ObservableCollection<string>(_people);

        // should be done in the XAML, but I'm lazy.
        private double defaultFontSize = 20;

        public MainWindow()
        {
            InitializeComponent();
            DrawDynamicUXElements();
        }
        // UX Helpers - These should move to an alternate assembly
        Button CreateButton(string Content, Action OnClick)
        {
            var button = new Button() { Content = Content };
            button.Click += (o, e) => { OnClick(); };

            button.FontSize = defaultFontSize;
            return button;
        }

        void MessageBoxPerson(string person)
        {
            MessageBox.Show(this, "Showing:" + person);
        }

        void DrawDynamicUXElements()
        {
            var dailyPagesButtons = new[]
            {
                CreateButton("_Today", dailyPages.GotoTodayPage),
                CreateButton("This _Week", dailyPages.GotoThisWeekPage),
                CreateButton("_Yesterday", dailyPages.GotoYesterday),
                CreateButton("Cleanup", erase.DeleteEmptySections),
            }.ToList();

            dailyPagesButtons.ForEach((b) => this.GridDailyPages.Children.Add(b));

            Action doNothing = ()=> { } ;

            var peoplePagesButtons = new[]{
                CreateButton("_Next", () =>
                {
                    MessageBoxPerson(this.PeopleList.SelectedValue as string);
                }),
                CreateButton("_Current", doNothing),
                CreateButton("_Last", doNothing),
            }.ToList();

            this.PeopleList.FontSize = defaultFontSize;
            this.PeopleList.ItemsSource = _observablePeople;
            // set the default item to a person.
            this.PeopleList.SelectedIndex = 0;
            peoplePagesButtons.ForEach((b) => this.GridPeoplePages.Children.Add(b));
        }
    }

}
