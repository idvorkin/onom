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
using OneNoteMenu;
using OneNoteObjectModel;
using OnenoteCapabilities;

namespace OneNotePieMenu
{
    /// <summary>
    /// Interaction logic for PieMenuWindow.xaml
    /// </summary>
    public partial class PieMenuWindow : Window
    {

        public ObservableCollection<string> _observablePeople;
        public AllOneNoteCapabilities capabilities = null;

        public PieMenuWindow()
        {

            InitializeComponent();
            capabilities = new AllOneNoteCapabilities();
            _observablePeople = new ObservableCollection<string>(capabilities.ListOfPeople);
            this.PeopleList.ItemsSource = _observablePeople;
            this.PeopleList.SelectedIndex = 0;
            RefreshTopicLruMenu();
        }


        /// <summary>
        /// Procsesing to enable the menu to be moved while pressing C-LButton
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainWindow_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // The goal is to be able to come up with a way to move the window given it has windowstyle=none (which is required for transparency).
            // DragMove is the way to instantiate a move, but the only time I can get DragMove() to work is from primary mouse button down.
            // without the if statement on the control key, the window will eat the mouse clicks, and the pie menu clicks will not occur.
            // Surprisingly, C-Click only works if i'm within the window but not on a pie menu.
            base.OnMouseLeftButtonDown(e);
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                this.DragMove();
            }
        }

        // DailyPages UX - 

        private void TimeMachine_OnClick(object sender, RoutedEventArgs e)
        {
            ToggleTimeMachineVisibility();
        }
        private void Today_Click(object sender, RoutedEventArgs e)
        {
            capabilities.DailyPages.GotoTodayPage();
        }

        private void Yesterday_OnClick(object sender, RoutedEventArgs e)
        {
            capabilities.DailyPages.GotoYesterday();
        }

        private void ThisWeek_OnClick(object sender, RoutedEventArgs e)
        {
            capabilities.DailyPages.GotoThisWeekPage();
        }

        private void ToggleTimeMachineVisibility()
        {
            if (TimeMachineDatePicker.IsVisible)
            {
                TimeMachineDatePicker.Visibility = Visibility.Hidden;
            }
            else
            {
                TimeMachineDatePicker.Visibility = Visibility.Visible;
            }
        }

        string selectedPerson()
        {
            return this.PeopleList.SelectedValue as string;
        }

        private void Augment_OnClick(object sender, RoutedEventArgs e)
        {
            capabilities.Augmenter.AugmentCurrentPage();
        }

        private void PersonNext_OnClick(object sender, RoutedEventArgs e)
        {
            capabilities.PeoplePages.GotoPersonNextPage(selectedPerson());
        }

        private void PersonToday_OnClick(object sender, RoutedEventArgs e)
        {
            capabilities.PeoplePages.GotoPersonCurrentMeetingPage(selectedPerson());
        }

        private void PersonPrev_OnClick(object sender, RoutedEventArgs e)
        {
            capabilities.PeoplePages.GotoPersonPreviousMeetingPage(selectedPerson());
        }

        private void RefreshTopicLruMenu()
        {
// HACKY - but will do for the demo.
            // find the topic menu
            var topicMenu = this.Menu1.Items.OfType<PieInTheSky.PieMenuItem>().First(mi => (string) mi.Header == "Topic");

            // All menu items that aren't picker and Refresh are LRU entires.
            var topicLRUs =
                topicMenu.Items.OfType<PieInTheSky.PieMenuItem>()
                    .Where(mi => ! "Refresh;Picker".Split(';').Contains((string) mi.Header))
                    .ToArray();

            // Debug.Assert(topicLRUs.Count() > 0 );

            // Get the latest topic entries from the notebook.
            var topicSection =
                capabilities.ona.GetNotebook(capabilities.SettingsTopicPages.TopicNotebook)
                    .PopulatedSection(capabilities.ona, capabilities.SettingsTopicPages.TopicSection);
            var topicTitleInLRUOrder = topicSection.Page.OrderByDescending(p => p.lastModifiedTime).Select(p => p.name);
            var replacementList = topicTitleInLRUOrder.Take(Math.Min(topicTitleInLRUOrder.Count(), topicLRUs.Count())).ToArray();
            foreach (var i in Enumerable.Range(0, replacementList.Count()))
            {
                topicLRUs[i].Header = replacementList[i];
            }
        }

        private void GotoTopic_OnClick(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as PieInTheSky.PieMenuItem;
            var pageTitle = menuItem.Header as string;

            var page = capabilities.ona.GetNotebook(capabilities.SettingsTopicPages.TopicNotebook)
                .PopulatedSection(capabilities.ona, capabilities.SettingsTopicPages.TopicSection)
                .Page.FirstOrDefault(p => p.name == pageTitle);
            if (page != null)
            {
                capabilities.ona.OneNoteApplication.NavigateTo(page.ID);
            }
        }

        private void Topic_OnClick(object sender, RoutedEventArgs e)
        {
            RefreshTopicLruMenu();
        }
    }
}
