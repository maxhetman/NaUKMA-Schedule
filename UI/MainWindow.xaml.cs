using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MYSchedule.DataAccess;
using MYSchedule.ExcelExport;

namespace UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            FillInfo();
            InitializeListeners();
        }

        private void InitializeListeners()
        {
            MethodistListeners();
            searchBtn.Click += OnSearchBtnClick;
            exportBtn.Click += OnExportBtnClick;
        }

        private void MethodistListeners()
        {
            mquery1.Selected += MethodistQuery1Selected;
            showAllClassroms.Checked += OnShowAllClassroomToggle;
            showAllClassroms.Unchecked += OnShowAllClassroomToggle;

            mquery2.Selected += MethodistQuery2Selected;
        }

        private void OnShowAllClassroomToggle(object sender, RoutedEventArgs e)
        {
            var isChecked = (bool) showAllClassroms.IsChecked;
            showComputerClassrooms.IsEnabled = !isChecked;
            classRoomNumbers.IsEnabled = !isChecked;
            buildings.IsEnabled = !isChecked;
        }

        private void MethodistQuery2Selected(object sender, RoutedEventArgs e)
        {
            methoditsParamsQuery1.Visibility = Visibility.Collapsed; ;
            methoditsParamsQuery2.Visibility = Visibility.Visible; ;
        }

        private void MethodistQuery1Selected(object sender, RoutedEventArgs e)
        {
            methoditsParamsQuery1.Visibility = Visibility.Visible; ;
            methoditsParamsQuery2.Visibility = Visibility.Collapsed; ;
        }

        private void OnExportBtnClick(object sender, RoutedEventArgs e)
        {
            if (mquery1.IsSelected)
            {
                ExcelExportManager.ShowAllClassRooms(GetCurrentData());
            }
        }

        private void OnSearchBtnClick(object sender, RoutedEventArgs e)
        {
            if (mquery1.IsSelected)
            {
                var dataTable = GetCurrentData();
                SetDataView(dataTable);
            }
        }

        private DataTable GetCurrentData()
        {
            if (mquery1.IsSelected)
            {
                return GetAvailableClassRooms();
            }
            else
            {
                return null;
            }
        }

        private DataTable GetAvailableClassRooms()
        {
            DataTable classRooms = null;
            var isShowAll = showAllClassroms.IsChecked;
            int? building = string.IsNullOrEmpty(buildings.Text) ? (int?)null : int.Parse(buildings.Text);
            var isComputer = showComputerClassrooms.IsChecked;
            var classRoomNumber = classRoomNumbers.Text;

            if (isComputer != true)
            {
                isComputer = null;
            }

            if ((bool) isShowAll)
            {
                return QueryManager.GetClassRoomsAvailability();
            }

            if (building == null && isComputer != true && string.IsNullOrEmpty(classRoomNumber))
            {
                Debug.WriteLine("return everything");
                return QueryManager.GetClassRoomsAvailability();
            }

            if (string.IsNullOrEmpty(classRoomNumber) == false)
            {
                return QueryManager.GetClassRoomsAvailability(classroomNumber:classRoomNumber);
            }

            return QueryManager.GetClassRoomsAvailability(buildingNumber: building, isComputer: isComputer);
        }

        private void SetDataView(DataTable dataTable)
        {
            dataView.DataContext = dataTable;
        }

        private void FillInfo()
        {
            buildings.ItemsSource = ClassRoomsDao.GetAllBuildings();
            classRoomNumbers.ItemsSource = ClassRoomsDao.GetAllNumbers();
        }
    }
}
