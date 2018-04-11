﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using MYSchedule.DataAccess;
using MYSchedule.DTO;
using MYSchedule.ExcelExport;
using MYSchedule.Parser;

namespace UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly BackgroundWorker worker = new BackgroundWorker();
        private string[] fileNames;

        public MainWindow()
        {
            InitializeComponent();
            InitLoader();
            FillInfo();
            InitializeListeners();
        }

        private void InitLoader()
        {
            worker.DoWork += worker_DoWork;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
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
            classRoomNumbers.SelectionChanged += OnClassRoomNumbersSelectionChanged;
            buildings.SelectionChanged += OnBuildingsSelectionChanged;
            mquery2.Selected += MethodistQuery2Selected;
        }

        private void OnBuildingsSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (buildings.SelectedIndex != -1)
                classRoomNumbers.SelectedIndex = -1;
        }

        private void OnClassRoomNumbersSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (classRoomNumbers.SelectedIndex != -1)
            {
                buildings.SelectedIndex = -1;
                showComputerClassrooms.IsChecked = false;
            }
        }

        private void OnShowAllClassroomToggle(object sender, RoutedEventArgs e)
        {
            var isAllClassRoomsChecked = (bool) showAllClassroms.IsChecked;
            classRoomNumbers.IsEnabled = !isAllClassRoomsChecked;
            buildings.IsEnabled = !isAllClassRoomsChecked;

            if (isAllClassRoomsChecked)
            {
                classRoomNumbers.SelectedIndex = -1;
                buildings.SelectedIndex = -1;
            }
        }

        private void MethodistQuery2Selected(object sender, RoutedEventArgs e)
        {
            methoditsParamsQuery1.Visibility = Visibility.Collapsed; ;
            methoditsParamsQuery2.Visibility = Visibility.Visible; ;
        }

        private void SetUiState(bool state)
        {
            mainGrid.IsEnabled = state;
            if (state == false)
            {
                mainGrid.Opacity = 50f;
            }
            else
            {
                mainGrid.Opacity = 100f;
            }
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
                if (dataTable.Rows.Count < 1)
                    ShowPopup("По заданим даним немає інформації");
                else
                    SetDataView(dataTable);
            }
        }

        private void ShowPopup(string message)
        {
            MessageBox.Show(message, "Розклад", MessageBoxButton.OK);
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

            int? building = string.IsNullOrEmpty(buildings.Text) ? (int?) null : int.Parse(buildings.Text);
            var isComputer = showComputerClassrooms.IsChecked;
            var classRoomNumber = classRoomNumbers.Text;            

            if (isComputer != true)
            {
                isComputer = null;
            }

            if (classRoomNumber == string.Empty)
            {
                classRoomNumber = null;
            }

            return QueryManager.GetClassRoomsAvailability(classroomNumber:classRoomNumber, buildingNumber:building, isComputer:isComputer);
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

        private void addExcelBtn_Click(object sender, RoutedEventArgs e)
        {
                // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;

                // Set filter for file extension and default file extension 
         //   dlg.Filter = "Excel (*.xls, *.xlt, *.xlm)";
            dlg.Filter = "Excel Files(*.xls;*.xlt;*.xlm;*.xlsx;*.xlsm)|*.xls;*.xlt;*.xlm;*.xlsx;*.xlsm";

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                loader.Visibility = Visibility.Visible;
                SetUiState(false);
                fileNames = dlg.FileNames;
                worker.RunWorkerAsync();
            }
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            // run all background tasks here

            foreach (var name in fileNames)
            {
                var schedule = ExcelParser.GetScheduleFromExcel(name);
                foreach (KeyValuePair<ScheduleRecordDto, List<int>> entry in schedule)
                {
                    bool isAdded = ScheduleRecordDao.AddIfNotExists(entry.Key);

                    if (!isAdded)
                    {
                        continue;
                    }

                    foreach (var weekNumber in entry.Value)
                    {
                        WeekScheduleDao.AddWeekSchedule(weekNumber: weekNumber, scheduleRecordId: entry.Key.Id);
                    }
                }
            }
        }

        private void worker_RunWorkerCompleted(object sender,
            RunWorkerCompletedEventArgs e)
        {
            loader.Visibility = Visibility.Hidden;
            SetUiState(true);
        }

        private static void ParseExcelFiles(string[] fileNames)
        {
           
        }
    }
}
