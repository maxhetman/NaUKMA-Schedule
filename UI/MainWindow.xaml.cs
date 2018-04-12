using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using MYSchedule.DataAccess;
using MYSchedule.DTO;
using MYSchedule.ExcelExport;
using MYSchedule.Parser;
using MYSchedule.Utils;

namespace UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly BackgroundWorker ImportExcelWorker = new BackgroundWorker();
        private readonly BackgroundWorker ExportExcelWorker = new BackgroundWorker();
        private string[] fileNames;

        public MainWindow()
        {
            InitializeComponent();
            InitWorkers();
            FillDropDownsInfo();
            InitializeListeners();
        }

        private void InitWorkers()
        {
            ImportExcelWorker.DoWork += ImportExcelWorkerDoWork;
            ImportExcelWorker.RunWorkerCompleted += ImportExcelWorkerRunImportExcelWorkerCompleted;

            ExportExcelWorker.DoWork += ExportExcelWorkerDoWork;
            ExportExcelWorker.RunWorkerCompleted += ExportExcelWorkerCompleted;
        }

        private void InitializeListeners()
        {
            MethodistListeners();
            SettingsListeners();
            searchBtn.Click += OnSearchBtnClick;
            exportBtn.Click += OnExportBtnClick;
            clearDbButton.Click += OnClearDbClick;
        }

        private void SettingsListeners()
        {
            weekDateSelector.SelectedDateChanged += OnFirstWeekDateChanged;
        }

        private void OnFirstWeekDateChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedDate = weekDateSelector.SelectedDate;

            if (selectedDate.Value.DayOfWeek != DayOfWeek.Monday)
            {
                ShowPopup("Перший день має бути понеділок!");
                return;
            }

            WeeksDao.SetFirstWeekDate((DateTime) weekDateSelector.SelectedDate);
        }

        private void OnClearDbClick(object sender, RoutedEventArgs e)
        {
            DBAccessManager.ClearDataBase();
        }

        private void MethodistListeners()
        {
            mquery1.Selected += MethodistQuery1Selected;
            mquery1reset.Click += OnMQuery1Reset;
            showAllClassroms.Checked += OnShowAllClassroomToggle;
            showAllClassroms.Unchecked += OnShowAllClassroomToggle;
            classRoomNumbers.SelectionChanged += OnClassRoomNumbersSelectionChanged;
            buildings.SelectionChanged += OnBuildingsSelectionChanged;
            mquery2.Selected += MethodistQuery2Selected;

            //consistensy check
            classroomConsistensy.Selected += CheckClassRoomConsistensy;
            teacherConsistensy.Selected += CheckTeacherConsistensy;
        }

        private void CheckTeacherConsistensy(object sender, RoutedEventArgs e)
        {
            mQueries.SelectedIndex = -1;
            var inconsistentData = DBAccessManager.GetInconsistentTeachers();
            SetDataView("Перевірка вчителів на несуперечливість", inconsistentData);
        }

        private void CheckClassRoomConsistensy(object sender, RoutedEventArgs e)
        {
            mQueries.SelectedIndex = -1;
            var inconsistentData = DBAccessManager.GetInconsistentClassrooms();
            SetDataView("Перевірка аудиторій на несуперечливість", inconsistentData);
        }

        private void OnMQuery1Reset(object sender, RoutedEventArgs e)
        {
            classRoomNumbers.SelectedIndex = -1;
            showComputerClassrooms.IsChecked = false;
            buildings.SelectedIndex = -1;
            showAllClassroms.IsChecked = false;
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
            methoditsParamsQuery1.Visibility = Visibility.Collapsed; 
            methoditsParamsQuery2.Visibility = Visibility.Visible;
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
            methoditsParamsQuery1.Visibility = Visibility.Visible; 
            methoditsParamsQuery2.Visibility = Visibility.Collapsed;
        }

        private void OnExportBtnClick(object sender, RoutedEventArgs e)
        {
            if (CurrentData == null || CurrentData.Rows.Count < 1)
            {
                ShowPopup("Немає даних для експорту\nНатисніть кнопку \"Виконати\"");
                return;
            }

            Action<string, DataTable> exportDelegate = null;

            if (methodistTab.IsSelected && mquery1.IsSelected)
            {
                exportDelegate = ExcelExportManager.ShowAllClassRooms;
            }

            if (teacherTab.IsSelected)
            {
                if (teacherSubjectScheduleQuery.IsSelected)
                {
                    exportDelegate = ExcelExportLessonByCourseAndSpecialty.LessonScheduleByCourseAndSpecialty;
                }
            }

            if (exportDelegate == null)
            {
                exportDelegate = GenericExcelExport.Export;
            }

            loader.Visibility = Visibility.Visible;
            SetUiState(false);
            ExportExcelWorker.RunWorkerAsync(new object[]{exportDelegate, CurrentData, dataViewHeader.Content});
        }

        private DataTable CurrentData
        {
            get
            {
                var itemsSource = (DataView) dataView.ItemsSource;
                return itemsSource?.ToTable();
            }
        }

        private void OnSearchBtnClick(object sender, RoutedEventArgs e)
        {
            var header = string.Empty;

            if (methodistTab.IsSelected)
            {
                if (mquery1.IsSelected)
                {
                    var dataTable = GetAvailableClassRooms();
                    if (dataTable.Rows.Count < 1)
                    {
                        ShowPopup("По заданим даним немає інформації");
                    }
                    else
                    {
                        header = GetMquery1Header();
                    }
                    SetDataView(header, dataTable);
                } else if (mquery2.IsSelected)
                {
                    var dataTable = GetScheduleForWeek();
                    if (dataTable.Rows.Count > 1)
                    {
                        header = GetMquery2Header();
                    }
                    SetDataView(header, dataTable);
                }
            } else if (teacherTab.IsSelected)
            {
                if (teacherSubjectScheduleQuery.IsSelected)
                {
                    var dataTable = GetSubjectSchedule();
                    if (dataTable.Rows.Count < 1)
                    {
                        ShowPopup("По заданим даним немає інформації");
                    }
                    else
                    {
                        header = GetSubjectScheduleHeader();
                    }
                    SetDataView(header, dataTable);
                }
            }

        }

        private string GetSubjectScheduleHeader()
        {
            var specialty = teacherSpecialtyCb.Text;
            int? yearOfStudying = string.IsNullOrEmpty(teacherYearOfStudyingCb.Text)
                ? (int?)null
                : int.Parse(teacherYearOfStudyingCb.Text);
            var subject = teacherSubjectCb.Text;

            return string.Format("{0}, {1} курс, {2}", specialty, yearOfStudying, subject);
        }

        private string GetMquery2Header()
        {
            return "Розклад на " + mquery2Weeks.Text;
        }

        private string GetMquery1Header()
        {
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

            if (!string.IsNullOrEmpty(classRoomNumber))
            {
                return string.Format("Зайнятість аудиторії {0}", classRoomNumber);
            }

            var classroomType = isComputer == null ? " " : " комп`ютерних ";
            var header = string.Format("Зайнятість{0}аудиторій ", classroomType);

            if (building != null)
            {
                header += string.Format("{0} корпусу ", building);
            }

            return header;
        }

        private DataTable GetScheduleForWeek()
        {
            string chosenWeeks = mquery2Weeks.Text;

            if (string.IsNullOrEmpty(chosenWeeks))
            {
                ShowPopup("Виберіть номер тижня");
                return new DataTable();
            }

            var weekNumber = int.Parse(chosenWeeks.Substring(0, chosenWeeks.IndexOf(" ")));
            var dataTable =  QueryManager.GetScheduleForWeek(weekNumber);

            if (!string.IsNullOrEmpty(chosenWeeks) && dataTable.Rows.Count < 1)
            {
                ShowPopup("По заданим даним немає інформації");
            }
            
            return dataTable;
        }

        private void SetDataView(string header, DataTable dataTable)
        {
            dataViewHeader.Content = header;
            dataView.DataContext = dataTable;
        }

        private  DataTable GetSubjectSchedule()
        {
            var specialty = teacherSpecialtyCb.Text;
            int? yearOfStudying = string.IsNullOrEmpty(teacherYearOfStudyingCb.Text)
                ? (int?) null
                : int.Parse(teacherYearOfStudyingCb.Text);
            var subject = teacherSubjectCb.Text;

            if (string.IsNullOrEmpty(specialty) || string.IsNullOrEmpty(subject) || yearOfStudying == null)
            {
                ShowPopup("Вкажіть всі фільтри");
                return new DataTable();
            }
            else
            {
                //sorry for this
                specialty = specialty.Replace("\"", "\"\"");
                return QueryManager.GetScheduleBySubjectSpecialtyAndCourse("\"" + specialty + "\"", (int) yearOfStudying , "\"" + subject + "\"");
            }


        }
    

        private void ShowPopup(string message)
        {
            MessageBox.Show(message, "Розклад", MessageBoxButton.OK);
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

        private void FillDropDownsInfo()
        {
            buildings.ItemsSource = ClassRoomsDao.GetAllBuildings();
            classRoomNumbers.ItemsSource = ClassRoomsDao.GetAllNumbers();
            teacherSpecialtyCb.ItemsSource = SpecialtyDao.GetAllSpecialties();
            teacherSubjectCb.ItemsSource = ScheduleRecordDao.GetAllSubjects();
            teacherYearOfStudyingCb.ItemsSource = ScheduleRecordDao.GetAllYears();
            mquery2Weeks.ItemsSource = WeeksDao.GetFormattedWeeks();
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
                ImportExcelWorker.RunWorkerAsync();
            }
        }

        private void ImportExcelWorkerDoWork(object sender, DoWorkEventArgs e)
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
                        //TSR
                        if (weekNumber == 8)
                            continue;

                        WeekScheduleDao.AddWeekSchedule(weekNumber: weekNumber, scheduleRecordId: entry.Key.Id);
                    }
                }
            }
        }

        private void ImportExcelWorkerRunImportExcelWorkerCompleted(object sender,
            RunWorkerCompletedEventArgs e)
        {
            FillDropDownsInfo();
            loader.Visibility = Visibility.Hidden;
            SetUiState(true);
        }

        private void ExportExcelWorkerDoWork(object sender, DoWorkEventArgs e)
        {
            var parameters = (object[]) e.Argument;
            var exportDelegate = (Action<string, DataTable>) parameters[0];
            var currData = (DataTable) parameters[1];
            var header = (string) parameters[2];
            exportDelegate.Invoke(header, currData);
        }

        private void ExportExcelWorkerCompleted(object sender,RunWorkerCompletedEventArgs e)
        {
            loader.Visibility = Visibility.Hidden;
            SetUiState(true);
        }

    }
}
