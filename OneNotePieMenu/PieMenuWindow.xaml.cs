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
    }
}
