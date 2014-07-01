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
using List = OneNoteObjectModel.List;

namespace OneNoteMenu
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static readonly OneNoteApp ona = new OneNoteApp();
        static readonly SettingsPeoplePages settingsPeoplePages = new SettingsPeoplePages();
        static readonly SettingsDailyPages settingsDailyPages = new SettingsDailyPages();
        readonly EraseEmpty erase = new EraseEmpty();
        readonly DailyPages dailyPages = new DailyPages(ona, settingsDailyPages);
        private readonly PeoplePages peoplePages = new PeoplePages(ona, settingsPeoplePages);
        private static string[] _people = new SettingsPeoplePages().People().ToArray();
        private ObservableCollection<string> _observablePeople = new ObservableCollection<string>(_people);
        private Augmenter augmenter;

        // should be done in the XAML, but I'm lazy.
        private double defaultFontSize = 20;

        public MainWindow()
        {
            InitializeComponent();
            DrawDynamicUXElements();

            // TBD: Look up a dependency injection mechanism.
            var smartTagProcessors = new List<ISmartTagProcessor>()
            {
                new TwitterSmartTagProcessor(), 
                new WikipediaSmartTagProcessor(), 
                new PeopleSmartTagProcessor(ona, settingsPeoplePages), 
                new DailySmartTagProcessor(ona,settingsDailyPages),
                // Topic smarttag processor needs to go last as it will create a topic page for any un-processed tag.
                new TopicSmartTagTopicProcessor(ona, settingsDailyPages)
            };
            var smartTagAugmentor = new SmartTagAugmenter(ona, new SettingsSmartTags(), smartTagProcessors);
            augmenter = new Augmenter(ona, new List<IPageAugmenter> {smartTagAugmentor});
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
            augmenter.AugmentCurrentPage();
        }

        void DrawDynamicUXElements()
        {
            var dailyPagesButtons = new[]
            {
                CreateButton("_Today", dailyPages.GotoTodayPage),
                CreateButton("This _Week", dailyPages.GotoThisWeekPage),
                CreateButton("_Yesterday", dailyPages.GotoYesterday),
                CreateButton("_Augment", Augment),
            }.ToList();

            dailyPagesButtons.ForEach((b) => this.GridDailyPages.Children.Add(b));

            Action doNothing = ()=> { } ;

            var peoplePagesButtons = new[]{
                CreateButton("Next", ()=> peoplePages.GotoPersonNextPage(selectedPerson())),
                CreateButton("_Previous", ()=> peoplePages.GotoPersonPreviousMeetingPage(selectedPerson())),
                CreateButton("_NewDay", ()=> peoplePages.GotoPersonCurrentMeetingPage(selectedPerson())),
            }.ToList();

            this.PeopleList.FontSize = defaultFontSize;
            this.PeopleList.ItemsSource = _observablePeople;
            // set the default item to a person.
            this.PeopleList.SelectedIndex = 0;
            peoplePagesButtons.ForEach((b) => this.GridPeoplePages.Children.Add(b));
        }
    }

}
