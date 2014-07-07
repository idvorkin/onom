using System;
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
using List = OneNoteObjectModel.List;

namespace OneNoteMenu
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<string> _observablePeople;
        public AllOneNoteCapabilities capabilities = null;

        // should be done in the XAML, but I'm lazy.
        private double defaultFontSize = 20;

        public MainWindow()
        {
            InitializeComponent();
            capabilities = new AllOneNoteCapabilities();
            _observablePeople = new ObservableCollection<string>(capabilities.ListOfPeople);
            DrawDynamicUXElements();
            CrashDumpWriter.InstallReportAndCreateCrashDumpUnhandledExceptionHandler();
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

        string selectedPerson()
        {
            return this.PeopleList.SelectedValue as string;
        }

        void Augment()
        {
            capabilities.Augmenter.AugmentCurrentPage();
        }

        void DrawDynamicUXElements()
        {
            var dailyPagesButtons = new[]
            {
                CreateButton("_Today", capabilities.DailyPages.GotoTodayPage),
                CreateButton("This _Week", capabilities.DailyPages.GotoThisWeekPage),
                CreateButton("_Yesterday", capabilities.DailyPages.GotoYesterday),
                CreateButton("_Augment", Augment),
            }.ToList();

            dailyPagesButtons.ForEach((b) => this.GridDailyPages.Children.Add(b));

            Action doNothing = ()=> { } ;

            var peoplePagesButtons = new[]{
                CreateButton("Next", ()=> capabilities.PeoplePages.GotoPersonNextPage(selectedPerson())),
                CreateButton("_Previous", ()=> capabilities.PeoplePages.GotoPersonPreviousMeetingPage(selectedPerson())),
                CreateButton("_NewDay", ()=> capabilities.PeoplePages.GotoPersonCurrentMeetingPage(selectedPerson())),
            }.ToList();

            this.PeopleList.FontSize = defaultFontSize;
            this.PeopleList.ItemsSource = _observablePeople;
            // set the default item to a person.
            this.PeopleList.SelectedIndex = 0;
            peoplePagesButtons.ForEach((b) => this.GridPeoplePages.Children.Add(b));
        }
    }

}
