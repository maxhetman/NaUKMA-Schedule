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
            TeachersListeners();
            SettingsListeners();
            StudentsListeners();
            searchBtn.Click += OnSearchBtnClick;
            exportBtn.Click += OnExportBtnClick;
            clearDbButton.Click += OnClearDbClick;
        }

        private void StudentsListeners()
        {
            studentSubjectScheduleQuery.Selected += studentSubjectScheduleQuerySelected;
            studentScheduleQuery.Selected += studentScheduleQuerySelected;
        }

        private void TeachersListeners()
        {
            teacherSubjectScheduleQuery.Selected += teacherSubjectScheduleQuerySelected;
            teacherScheduleQuery.Selected += teacherScheduleQuerySelected;
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
            var resut = MessageBox.Show("Ви впевнені, що хочете видалити дані?", "Розклад", MessageBoxButton.YesNo);
            if (resut == MessageBoxResult.Yes)
            {
                DBAccessManager.ClearDataBase();
                ShowPopup("Дані видалено");
                FillDropDownsInfo();
            }
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

        private void MethodistQuery2Selected(object sender, RoutedEventArgs e)
        {
            methoditsParamsQuery1.Visibility = Visibility.Collapsed;
            methoditsParamsQuery2.Visibility = Visibility.Visible;
        }

        private void teacherSubjectScheduleQuerySelected(object sender, RoutedEventArgs e)
        {
            teacherSubjectScheduleParams.Visibility = Visibility.Visible;
            teacherScheduleParams.Visibility = Visibility.Collapsed;
        }

        private void studentSubjectScheduleQuerySelected(object sender, RoutedEventArgs e)
        {
            studentSubjectScheduleParams.Visibility = Visibility.Visible;
            studentScheduleParams.Visibility = Visibility.Collapsed;
        }

        private void teacherScheduleQuerySelected(object sender, RoutedEventArgs e)
        {
            teacherScheduleParams.Visibility = Visibility.Visible;
            teacherSubjectScheduleParams.Visibility = Visibility.Collapsed;         
        }


        private void studentScheduleQuerySelected(object sender, RoutedEventArgs e)
        {
            studentScheduleParams.Visibility = Visibility.Visible;
            studentSubjectScheduleParams.Visibility = Visibility.Collapsed;
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

            if (studentTab.IsSelected)
            {
                if (studentSubjectScheduleQuery.IsSelected)
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
                    if (dataTable == null)
                    {
                        return;
                    }
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
                    if (dataTable == null)
                    {
                        return;
                    }
                    if (dataTable.Rows.Count < 1)
                    {
                        ShowPopup("По заданим даним немає інформації");
                    } else 
                    {
                        header = GetMquery2Header();

                    }
                    SetDataView(header, dataTable);
                }
            } else if (teacherTab.IsSelected)
            {
                if (teacherSubjectScheduleQuery.IsSelected)
                {
                    var dataTable = GetSubjectScheduleTeacherTab();
                    if (dataTable == null)
                        return;
                    if (dataTable.Rows.Count < 1)
                    {
                        ShowPopup("По заданим даним немає інформації");
                    }
                    else
                    {
                        header = GetSubjectScheduleHeaderTeacherTab();
                    }
                    SetDataView(header, dataTable);
                }else if (teacherScheduleQuery.IsSelected)
                {
                    var dataTable = GetTeacherSchedule();
                    if (dataTable == null)
                        return;
                    if (dataTable.Rows.Count < 1)
                    {
                        ShowPopup("По заданим даним немає інформації");
                    }
                    else
                    {
                        header = GetTeacherScheduleHeader();
                    }
                    SetDataView(header, dataTable);
                }
            } else if (studentTab.IsSelected)
            {
                if (studentSubjectScheduleQuery.IsSelected)
                {
                    var dataTable = GetSubjectScheduleStudentsTab();
                    if (dataTable == null)
                        return;
                    if (dataTable.Rows.Count < 1)
                    {
                        ShowPopup("По заданим даним немає інформації");
                    }
                    else
                    {
                        header = GetSubjectScheduleHeaderStudentTab();
                    }
                    SetDataView(header, dataTable);
                } else if (studentScheduleQuery.IsSelected)
                {
                    var dataTable = GetStudentSchedule();
                    if (dataTable == null)
                        return;
                    if (dataTable.Rows.Count < 1)
                    {
                        ShowPopup("По заданим даним немає інформації");
                    }
                    else
                    {
                        header = GetStudentScheduleHeader();
                    }
                    SetDataView(header, dataTable);
                }
            }

        }

        private string GetTeacherScheduleHeader()
        {
            var teacher = teacherNameSelect.Text;
            var weekNumberStr = teacherWeekSelect.Text;

            return string.Format("Розклад для {0} на {1}", teacher, weekNumberStr);
        }

        private string GetStudentScheduleHeader()
        {
            var specialty = studentSpecialtySelect.Text;
            var weekNumberStr = studentWeekSelect.Text;

            return string.Format("{0}, {1}", specialty, weekNumberStr);
        }

        private DataTable GetStudentSchedule()
        {
            var weekNumberStr = studentWeekSelect.Text;
            var specialty = studentSpecialtySelect.Text;

            if (string.IsNullOrEmpty(weekNumberStr) || string.IsNullOrEmpty(specialty))
            {
                ShowPopup("Виберіть всі параметри");
                return null;
            }

            specialty = specialty.Replace("\"", "\"\"");

            if (weekNumberStr == "Всі тижні")
            {
                return QueryManager.GetStudentScheduleForAllWeeks(specialty);
            }
            var weekNumber = int.Parse(weekNumberStr.Substring(0, weekNumberStr.IndexOf(" ")));
            return QueryManager.GetStudentScheduleForSelectedWeek(specialty, weekNumber);
        }

        private DataTable GetTeacherSchedule()
        {
            var weekNumberStr = teacherWeekSelect.Text;
            var teacher = teacherNameSelect.Text;

            if (string.IsNullOrEmpty(weekNumberStr) || string.IsNullOrEmpty(teacher))
            {
                ShowPopup("Виберіть всі параметри");
                return null;
            }

            var spaceIndex = teacher.IndexOf(" ");
            var lastName = teacher.Substring(0, spaceIndex);
            var initials = teacher.Substring(spaceIndex + 1, teacher.Length - spaceIndex - 1);

            if (weekNumberStr == "Всі тижні")
            {
                return QueryManager.GetTeacherScheduleForAllWeeks(lastName, initials);
            }

            var weekNumber = int.Parse(weekNumberStr.Substring(0, weekNumberStr.IndexOf(" ")));
            return QueryManager.GetTeacherScheduleForSelectedWeek(lastName, initials, weekNumber);
        }

        private string GetSubjectScheduleHeaderTeacherTab()
        {
            var specialty = teacherSpecialtyCb.Text;
            int? yearOfStudying = string.IsNullOrEmpty(teacherYearOfStudyingCb.Text)
                ? (int?)null
                : int.Parse(teacherYearOfStudyingCb.Text);
            var subject = teacherSubjectCb.Text;

            return string.Format("{0}, {1} курс, {2}", specialty, yearOfStudying, subject);
        }

        private string GetSubjectScheduleHeaderStudentTab()
        {
            var specialty = studentSpecialtyCb.Text;
            int? yearOfStudying = string.IsNullOrEmpty(studentYearOfStudyingCb.Text)
                ? (int?)null
                : int.Parse(studentYearOfStudyingCb.Text);
            var subject = studentSubjectCb.Text;

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
                return null;
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

        private  DataTable GetSubjectScheduleTeacherTab()
        {
            var specialty = teacherSpecialtyCb.Text;
            int? yearOfStudying = string.IsNullOrEmpty(teacherYearOfStudyingCb.Text)
                ? (int?) null
                : int.Parse(teacherYearOfStudyingCb.Text);
            var subject = teacherSubjectCb.Text;

            if (string.IsNullOrEmpty(specialty) || string.IsNullOrEmpty(subject) || yearOfStudying == null)
            {
                ShowPopup("Виберіть всі параметри");
                return null;
            }
            else
            {
                //sorry for this
                specialty = specialty.Replace("\"", "\"\"");
                return QueryManager.GetScheduleBySubjectSpecialtyAndCourse("\"" + specialty + "\"", (int) yearOfStudying , "\"" + subject + "\"");
            }

        }

        private DataTable GetSubjectScheduleStudentsTab()
        {
            var specialty = studentSpecialtyCb.Text;
            int? yearOfStudying = string.IsNullOrEmpty(studentYearOfStudyingCb.Text)
                ? (int?)null
                : int.Parse(studentYearOfStudyingCb.Text);
            var subject = studentSubjectCb.Text;

            if (string.IsNullOrEmpty(specialty) || string.IsNullOrEmpty(subject) || yearOfStudying == null)
            {
                ShowPopup("Виберіть всі параметри");
                return null;
            }
            else
            {
                //sorry for this
                specialty = specialty.Replace("\"", "\"\"");
                return QueryManager.GetScheduleBySubjectSpecialtyAndCourse("\"" + specialty + "\"", (int)yearOfStudying, "\"" + subject + "\"");
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
            var isShowAllClassroms = showAllClassroms.IsChecked;

            if (isComputer != true)
            {
                isComputer = null;
            }

            if (classRoomNumber == string.Empty)
            {
                classRoomNumber = null;
            }

            if (building == null && isComputer == null && classRoomNumber == null && isShowAllClassroms != true)
            {
                ShowPopup("Виберіть хоча б один параметр пошуку");
                return null;
            }

            return QueryManager.GetClassRoomsAvailability(classroomNumber:classRoomNumber, buildingNumber:building, isComputer:isComputer);
        }

        private void FillDropDownsInfo()
        {
            buildings.ItemsSource = ClassRoomsDao.GetAllBuildings();
            classRoomNumbers.ItemsSource = ClassRoomsDao.GetAllNumbers();
            
            var allSpecialties = SpecialtyDao.GetAllSpecialties();
            var allSubjects = ScheduleRecordDao.GetAllSubjects();
            var allYears = ScheduleRecordDao.GetAllYears();

            teacherSpecialtyCb.ItemsSource = allSpecialties;
            teacherSubjectCb.ItemsSource = allSubjects;
            teacherYearOfStudyingCb.ItemsSource = allYears;

            studentSpecialtyCb.ItemsSource = allSpecialties;
            studentSubjectCb.ItemsSource = allSubjects;
            studentYearOfStudyingCb.ItemsSource = allYears;

            var allWeeks = WeeksDao.GetFormattedWeeks();
            mquery2Weeks.ItemsSource = allWeeks;

            var selectWeeks = new string[allWeeks.Length + 1];
            selectWeeks[0] = "Всі тижні";
            Array.Copy(allWeeks,0, selectWeeks, 1, allWeeks.Length);
            teacherWeekSelect.ItemsSource = selectWeeks;
            teacherWeekSelect.SelectedIndex = 0;

            teacherNameSelect.ItemsSource = TeacherDao.GetFormattedTeachers();

            studentWeekSelect.ItemsSource = selectWeeks;
            studentWeekSelect.SelectedIndex = 0;
            studentSpecialtySelect.ItemsSource = allSpecialties;
        }

        private void addExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;

            dlg.Filter = "Excel Files(*.xls;*.xlt;*.xlm;*.xlsx;*.xlsm)|*.xls;*.xlt;*.xlm;*.xlsx;*.xlsm";

            Nullable<bool> result = dlg.ShowDialog();

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
