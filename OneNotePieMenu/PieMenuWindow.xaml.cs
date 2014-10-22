using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
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
using CSharpAnalytics;
using CSharpAnalytics.Sessions;
using OneNoteMenu;
using OneNoteObjectModel;
using OnenoteCapabilities;
using Path = System.IO.Path;

namespace OneNotePieMenu
{
    /// <summary>
    /// Interaction logic for PieMenuWindow.xaml
    /// </summary>
    public partial class PieMenuWindow : Window
    {
        // copied from: http://stackoverflow.com/questions/1600962/displaying-the-build-date
        private DateTime RetrieveLinkerTimestamp()
        {
            string filePath = System.Reflection.Assembly.GetCallingAssembly().Location;
            const int c_PeHeaderOffset = 60;
            const int c_LinkerTimestampOffset = 8;
            byte[] b = new byte[2048];
            System.IO.Stream s = null;

            try
            {
                s = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                s.Read(b, 0, 2048);
            }
            finally
            {
                if (s != null)
                {
                    s.Close();
                }
            }

            int i = System.BitConverter.ToInt32(b, c_PeHeaderOffset);
            int secondsSince1970 = System.BitConverter.ToInt32(b, i + c_LinkerTimestampOffset);
            DateTime dt = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            dt = dt.AddSeconds(secondsSince1970);
            dt = dt.ToLocalTime();
            return dt;
        }


        public static RoutedCommand AugmentCommand = new RoutedCommand();
        public ObservableCollection<string> _observablePeople;
        public AllOneNoteCapabilities capabilities = null;

        public void InitAnalytics()
        {
            AutoMeasurement.Instance = new WpfAutoMeasurement();
            AutoMeasurement.DebugWriter = d => Debug.WriteLine(d);

            // UA... is the APP-ID for OneNoteLabs/OneNotePieMenu - to get access mail idvorkin
            var config = new MeasurementConfiguration("UA-53825421-1","OneNotePieMenu", RetrieveLinkerTimestamp().ToString());
            AutoMeasurement.Start(config);
        }

        public PieMenuWindow()
        {
            var initTimer = Stopwatch.StartNew();
            InitializeComponent();
            InitAnalytics();
            capabilities = new AllOneNoteCapabilities();
            _observablePeople = new ObservableCollection<string>(capabilities.ListOfPeople);
            this.PeopleList.ItemsSource = _observablePeople;
            this.PeopleList.SelectedIndex = 0;
            RefreshTopicLruMenu();
            CrashDumpWriter.InstallReportAndCreateCrashDumpUnhandledExceptionHandler();
            AugmentCommand.InputGestures.Add( new KeyGesture( Key.A , ModifierKeys.Control ));
            AutoMeasurement.Client.TrackTimedEvent("Init","Duration",initTimer.Elapsed);
        }



        /// <summary>
        /// Procsesing to enable the menu to be moved while pressing C-LButton
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainWindow_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("Windows", MethodBase.GetCurrentMethod().Name);
            // The goal is to be able to come up with a way to move the window given it has windowstyle=none (which is required for transparency).
            // DragMove is the way to instantiate a move, but the only time I can get DragMove() to work is from primary mouse button down.
            base.OnMouseLeftButtonDown(e);

            var point = e.MouseDevice.GetPosition(this.Menu1);
            var isOnMenu = this.Menu1.IsMenuRelativePointOnMenu(point);
            if (isOnMenu)
            {
                // do nothing if the point is a menu interaction
                return;
            }

            // left clicked not on menu, initiate the drag.
            this.DragMove();
        }

        private void PieMenuWindow_OnMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("Windows", MethodBase.GetCurrentMethod().Name);
            var point = e.MouseDevice.GetPosition(this.Menu1);
            var isOnMenu = this.Menu1.IsMenuRelativePointOnMenu(point);
            if (isOnMenu)
            {
                // do nothing if the point is a menu interaction
                return;
            }

            // right clicked not on menu, show context menu.
            var contextMenu = this.FindResource("DefaultContextMenu") as ContextMenu;
            contextMenu.PlacementTarget = sender as Window;
            contextMenu.IsOpen = true;
        }

        // DailyPages UX - 

        private void TimeMachine_OnClick(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("DailyPages", MethodBase.GetCurrentMethod().Name);
            ToggleTimeMachineVisibility();
        }
        private void Today_Click(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("DailyPages", MethodBase.GetCurrentMethod().Name);
            capabilities.DailyPages.GotoTodayPage();
        }

        private void Yesterday_OnClick(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("DailyPages", MethodBase.GetCurrentMethod().Name);
            capabilities.DailyPages.GotoYesterday();
        }

        private void ThisWeek_OnClick(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("DailyPages", MethodBase.GetCurrentMethod().Name);
            capabilities.DailyPages.GotoThisWeekPage();
        }

        private void ToggleTimeMachineVisibility()
        {
            AutoMeasurement.Client.TrackEvent("DailyPages", MethodBase.GetCurrentMethod().Name);
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
            var timer = Stopwatch.StartNew();
            capabilities.Augmenter.AugmentCurrentPage();
            AutoMeasurement.Client.TrackTimedEvent("Augment","Duration",timer.Elapsed);
        }

        private void PersonNext_OnClick(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("PeoplePages", MethodBase.GetCurrentMethod().Name);
            capabilities.PeoplePages.GotoPersonNextPage(selectedPerson());
        }

        private void PersonToday_OnClick(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("PeoplePages", MethodBase.GetCurrentMethod().Name);
            capabilities.PeoplePages.GotoPersonCurrentMeetingPage(selectedPerson());
        }

        private void PersonPrev_OnClick(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("PeoplePages", MethodBase.GetCurrentMethod().Name);
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
            var topicSection = OneNoteApplication.Instance.GetNotebook(capabilities.SettingsTopicPages.TopicNotebook)
                    .PopulatedSection(capabilities.SettingsTopicPages.TopicSection);
            var topicTitleInLRUOrder = topicSection.Page.OrderByDescending(p => p.lastModifiedTime).Select(p => p.name);

            // UX order of menu-items is backwards, so fill in from back.
            var replacementList = topicTitleInLRUOrder.Take(Math.Min(topicTitleInLRUOrder.Count(), topicLRUs.Count())).Reverse().ToArray();
            foreach (var i in Enumerable.Range(0, replacementList.Count()))
            {
                topicLRUs[i].Header = replacementList[i];
            }
        }

        private void GotoTopic_OnClick(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("TopicPages", MethodBase.GetCurrentMethod().Name);
            var menuItem = sender as PieInTheSky.PieMenuItem;
            var pageTitle = menuItem.Header as string;

            var page = OneNoteApplication.Instance.GetNotebook(capabilities.SettingsTopicPages.TopicNotebook)
                .PopulatedSection(capabilities.SettingsTopicPages.TopicSection)
                .GetPage(pageTitle);
            if (page != null)
            {
                OneNoteApplication.Instance.InteropApplication.NavigateTo(page.ID);
            }
        }

        private void Topic_OnClick(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("TopicPages", MethodBase.GetCurrentMethod().Name);
            RefreshTopicLruMenu();
        }

        private void CurrentPerson_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            HideTextBoxAndShowComboBox();
        }
        private void PeopleList_OnDropDownClosed(object sender, EventArgs e)
        {
            HideComboBoxAndShowTextBox();
        }

        private void HideTextBoxAndShowComboBox()
        {
            this.CurrentPerson.Visibility = Visibility.Hidden;
            this.PeopleList.Visibility = Visibility.Visible;
        }

        private void HideComboBoxAndShowTextBox()
        {
            this.PeopleList.Visibility = Visibility.Hidden;
            this.CurrentPerson.Visibility = Visibility.Visible;
        }

        private void MinimizeClicked(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("Windows", MethodBase.GetCurrentMethod().Name);
            this.WindowState = WindowState.Minimized;
        }

        private void AboutClicked(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Check out https://github.com/idvorkin/onom");
        }

        private void ExitClicked(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("Windows", MethodBase.GetCurrentMethod().Name);
            Application.Current.Shutdown();
        }

        private void OnTopClicked(object sender, RoutedEventArgs e)
        {
            AutoMeasurement.Client.TrackEvent("Windows", MethodBase.GetCurrentMethod().Name);
            this.Topmost = true;
            // Update the Opacity so we can see edges when on top.
            var backGroundBrush = (this.Background as SolidColorBrush);
            backGroundBrush.Opacity = 0.1;
            backGroundBrush.Color = Colors.DarkGray;
        }

    }
}
