using GCGHS.Workers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using GCGHS.Properties;
using System.Net.Mail;
using System.Net;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace GCGHS
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string connectionString; //переменная со строкой подключения к нужной нам БД

        public ObservableCollection<WorkerCenter> WorkersCenter { get; set; }       //обновляемая коллекция с работниками центра
        public ObservableCollection<WorkerRegion> WorkersRegion { get; set; }       //обновляемая коллекция с работниками регионов
        public ObservableCollection<Otdel> Otdels { get; set; }                     //с отделами
        List<string> okrug = new List<string>() { "ЦАО", "ВАО", "ЗАО", "САО", "ЮАО", "СВАО", "СЗАО", "ЮЗАО", "ЮВАО", "ТАО", "НАО", "ЗелАО" };

        public MainWindow()
        {
            InitializeComponent();

            tb_ConnectServerName.Text = Settings.Default.ServerName;
            tb_ConnectBaseName.Text = Settings.Default.BaseName;
            rb_SendMailOutlook.IsChecked = Settings.Default.MailOutlook;
            rb_SendMailVnutr.IsChecked = !Settings.Default.MailOutlook;
            tb_MailLogin.Text = Settings.Default.MailLogin;
            tb_MailPass.Text = Settings.Default.MailPass;
            tb_MailTheme.Text = Settings.Default.MailTheme;
            tb_MailTo.Text = Settings.Default.MailTo;
            tb_SMTPserver.Text = Settings.Default.SMTPserver;
            tb_RadminPath.Text = Settings.Default.RadminPath;
            tb_ScannerPath.Text = Settings.Default.ScannerPath;

            connectionString = @"Data Source=" + tb_ConnectServerName.Text + ";Initial Catalog=" + tb_ConnectBaseName.Text + ";Integrated Security=True"; //генерим строку подключения используя данные на странице настроек

            WorkersCenter = new ObservableCollection<WorkerCenter>();       //инициализация в памяти коллекция с работниками и отделами
            WorkersRegion = new ObservableCollection<WorkerRegion>();
            Otdels = new ObservableCollection<Otdel>();
        }

        private void bt_Search_Click(object sender, RoutedEventArgs e) //обработка нажатия кнопки поиск
        {
            SearchWorker();             //запускаем метод поиска
        }

        private void bt_Write_Click(object sender, RoutedEventArgs e)
        {
            if (WorkersGrid.SelectedIndex != -1)
            {
                WriteData();
                RefreshDataInGrid();
            }
            else { MessageBox.Show("Не выбран сотрудник для записи."); }
        }

        public void SearchWorker()          //что логично, метод поиска работника
        {
            WorkersCenter.Clear();      //чистим сначала обе коллекции с работниками, чтобы не копился хлам после прошлых запросов
            WorkersRegion.Clear();
            SqlConnection connection = new SqlConnection(connectionString);         //создаем подключение
            SqlCommand command = new SqlCommand("search_WorkerByName", connection); //выбираем объект из базы
            command.CommandType = CommandType.StoredProcedure;                      //определяем тип, что это хранимая процедура
            command.Parameters.AddWithValue("Fio", tb_Fio.Text);                    //передаем хранимой процедуре переменную Fio
            connection.Open();                                                      //открываем соединение
            try
            {
                SqlDataReader reader = command.ExecuteReader();  //чтец
                while (reader.Read()) //построчно читает таблицу
                {
                    WorkersCenter.Add(new WorkerCenter(Convert.ToInt32(reader["id"]), Convert.ToString(reader["Fio"]), Convert.ToString(reader["User"]), Convert.ToString(reader["Login"]), Convert.ToString(reader["Pass"]), Convert.ToString(reader["Otdel"]), Convert.ToString(reader["NumberOtdel"]), Convert.ToString(reader["TelOtdel"]), Convert.ToString(reader["TelVnutr"]), Convert.ToString(reader["IP"]), Convert.ToString(reader["Mail"]), Convert.ToString(reader["Komment"]), Convert.ToString(reader["KabNumb"])));
                }
                if (WorkersCenter.Count == 0)       //если по запросу пусто, и в коллекцию ничего не записалось, выдаем сообщение
                {
                    MessageBox.Show("Данные не найдены");
                }
            }
            catch (InvalidOperationException) { MessageBox.Show("Подключение закрыто.", "ОШИБКА", MessageBoxButton.OK);}
            catch (SqlException) { MessageBox.Show("SQL сервер вернул ошибку", "ОШИБКА", MessageBoxButton.OK);}
            catch {MessageBox.Show("Неизвестная ошибка", "ОШИБКА", MessageBoxButton.OK);}

            connection.Close(); //закрываем соединение

            if (rb_Center.IsChecked == true)
            {
                WorkersGrid.ItemsSource = WorkersCenter;
            }
            else WorkersGrid.ItemsSource = WorkersRegion;
            
        }

        private void bt_NewWorker_Click(object sender, RoutedEventArgs e)
        {
            NewWorkerWrite();
            RefreshDataInGrid();
        }

        public void GetOtdelList()          //метод загрузки в программу списка отделов
        {
            cb_Okrug.ItemsSource = okrug;

            SqlConnection connection = new SqlConnection(connectionString); //создаем подключение
            SqlCommand command;

            if (rb_Center.IsChecked == true) {command = new SqlCommand("GetOtdel", connection);} //выбираем объект из базы
            else {command = new SqlCommand("GetOtdelRegion", connection); }

            command.CommandType = CommandType.StoredProcedure; //определяем тип, что это хранимая процедура
            connection.Open(); //открываем соединение
            try
            {
                SqlDataReader reader = command.ExecuteReader();  //чтец
                while (reader.Read()) //построчно читает таблицу
                {
                    Otdels.Add(new Otdel(Convert.ToString(reader["NumberOtdel"]), Convert.ToString(reader["NameOtdel"]), Convert.ToString(reader["Okrug"])));
                }
            }
            catch (InvalidOperationException) { MessageBox.Show("Подключение закрыто.", "ОШИБКА", MessageBoxButton.OK); }
            catch (SqlException) { MessageBox.Show("SQL сервер вернул ошибку", "ОШИБКА", MessageBoxButton.OK); }
            catch { MessageBox.Show("Неизвестная ошибка", "ОШИБКА", MessageBoxButton.OK); }
            connection.Close(); //закрываем соединение

            for (int i = 0; i < Otdels.Count; i++)
            {
                cb_OtdelName.Items.Insert(i, Otdels[i].OtdelName);
                cb_OtdelNumb.Items.Insert(i, Otdels[i].OtdelNumber);
            }
        }

        public string GetPass(int lenght)   //метод генерации пароля
        {
            string symbols = "qwertyupasdfghjkzxvbnmQWERTYUPASDFGHJKZXVBNM123456789"; //набор символов
            string result = "";

            Random rnd = new Random();
            int lng = symbols.Length;
            for (int i = 0; i < lenght; i++)
                result += symbols[rnd.Next(lng)];
            return result;
        }

        private void cb_OtdelNumb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int x = cb_OtdelNumb.SelectedIndex;
            cb_OtdelName.SelectedIndex = x;
            // = okrug[x];
            if (cb_OtdelNumb.SelectedIndex != -1)
            {
                SearchOtdel();
            }           
        }

        private void bt_SaveSettings_Click(object sender, RoutedEventArgs e)
        {
            Settings.Default.ServerName = tb_ConnectServerName.Text;
            Settings.Default.BaseName = tb_ConnectBaseName.Text;
            Settings.Default.MailLogin = tb_MailLogin.Text;
            Settings.Default.MailPass = tb_MailPass.Text;
            Settings.Default.MailTheme = tb_MailTheme.Text;
            Settings.Default.MailTo = tb_MailTo.Text;
            Settings.Default.SMTPserver = tb_SMTPserver.Text;
            Settings.Default.RadminPath = tb_RadminPath.Text;
            Settings.Default.ScannerPath = tb_ScannerPath.Text;
            if (rb_SendMailOutlook.IsChecked == true)
            {
                Settings.Default.MailOutlook = true;
            }
            else Settings.Default.MailOutlook = false;


            Settings.Default.Save();
            MessageBox.Show("Настройки сохранены", "Message", MessageBoxButton.OK);
        }

        private void cb_Okrug_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void cb_OtdelName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int x = cb_OtdelName.SelectedIndex;
            cb_OtdelNumb.SelectedIndex = x;
        }

        private void rb_Center_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                cb_OtdelName.Items.Clear();
                cb_OtdelNumb.Items.Clear();
                Otdels.Clear();
                cb_Okrug.IsEnabled = false;
            }
            catch { }
            
            try
            {
                GetOtdelList();             //загрузка из базы отделов (целиком с именами, номерами и округами)
            }
            catch { MessageBox.Show("Ошибка подключения к базе", "ОШИБКА", MessageBoxButton.OK); } //ловим ошибку если что-то пошло не так            
        }

        private void rb_Region_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                cb_OtdelName.Items.Clear();
                cb_OtdelNumb.Items.Clear();
                Otdels.Clear();
                cb_Okrug.IsEnabled = true;
            }
            catch { }

            try
            {
                GetOtdelList();             //загрузка из базы отделов (целиком с именами, номерами и округами)
            }
            catch { MessageBox.Show("Ошибка подключения к базе", "ОШИБКА", MessageBoxButton.OK); } //ловим ошибку если что-то пошло не так
        }

        public void SearchOtdel()
        {
            WorkersCenter.Clear();      //чистим сначала обе коллекции с работниками, чтобы не копился хлам после прошлых запросов
            WorkersRegion.Clear();
            SqlConnection connection = new SqlConnection(connectionString);         //создаем подключение
            SqlCommand command;
            if (rb_Center.IsChecked == true)
            {
                command = new SqlCommand("search_WorkersByOtdelInCenter", connection); //выбираем объект из базы
            }
            else
            {
                command = new SqlCommand("search_WorkersByOtdelInRegion", connection); //выбираем объект из базы
            }
            command.CommandType = CommandType.StoredProcedure;                      //определяем тип, что это хранимая процедура
            command.Parameters.AddWithValue("NumberOtdel", cb_OtdelNumb.SelectedValue);                    //передаем хранимой процедуре переменную Fio
            
            connection.Open();                                                      //открываем соединение
            try
            {
                SqlDataReader reader = command.ExecuteReader();  //чтец
                if (rb_Center.IsChecked == true)
                {
                    while (reader.Read()) //построчно читает таблицу
                    {
                        WorkersCenter.Add(new WorkerCenter(Convert.ToInt32(reader["id"]), Convert.ToString(reader["Fio"]), Convert.ToString(reader["User"]), Convert.ToString(reader["Login"]), Convert.ToString(reader["Pass"]), Convert.ToString(reader["Otdel"]), Convert.ToString(reader["NumberOtdel"]), Convert.ToString(reader["TelOtdel"]), Convert.ToString(reader["TelVnutr"]), Convert.ToString(reader["IP"]), Convert.ToString(reader["Mail"]), Convert.ToString(reader["Komment"]), Convert.ToString(reader["KabNumb"])));
                    }
                    if (WorkersCenter.Count == 0)       //если по запросу пусто, и в коллекцию ничего не записалось, выдаем сообщение
                    {
                        MessageBox.Show("Данные не найдены");
                    }
                }
                else
                {
                    while (reader.Read()) //построчно читает таблицу
                    {
                        WorkersRegion.Add(new WorkerRegion(Convert.ToInt32(reader["id"]), Convert.ToString(reader["Fio"]), Convert.ToString(reader["User"]), Convert.ToString(reader["Login"]), Convert.ToString(reader["Pass"]), Convert.ToString(reader["Otdel"]), Convert.ToString(reader["NumberOtdel"]), Convert.ToString(reader["TelOtdel"]), Convert.ToString(reader["IP"]), Convert.ToString(reader["Mail"]), Convert.ToString(reader["Komment"])));
                    }
                    if (WorkersRegion.Count == 0)       //если по запросу пусто, и в коллекцию ничего не записалось, выдаем сообщение
                    {
                        MessageBox.Show("Данные не найдены");
                    }
                }
            }
            catch (InvalidOperationException) { MessageBox.Show("Подключение закрыто.", "ОШИБКА", MessageBoxButton.OK); }
            catch (SqlException) { MessageBox.Show("SQL сервер вернул ошибку", "ОШИБКА", MessageBoxButton.OK); }
            catch { MessageBox.Show("Неизвестная ошибка", "ОШИБКА", MessageBoxButton.OK); }

            connection.Close(); //закрываем соединение

            if (rb_Center.IsChecked == true)
            {
                WorkersGrid.ItemsSource = WorkersCenter;
            }
            else WorkersGrid.ItemsSource = WorkersRegion;
        }

        public void SetData()
        {

        }

        private void WorkersGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (rb_Center.IsChecked == true)
            {
                try
                {
                    WorkerCenter selected = WorkersGrid.SelectedItem as WorkerCenter;
                    tb_Fio.Text = selected.Name;
                    tb_IP.Text = selected.Ip;
                    tb_Login.Text = selected.Login;
                    tb_TelOtdel.Text = selected.TelOtdel;
                    tb_TelVnutr.Text = selected.TelVnutr;
                    tb_Pass.Text = selected.Pass;
                    tb_User.Text = selected.User;
                    tb_roomNumber.Text = selected.RoomNumber.ToString();
                    tb_Comment.Text = selected.Komment;
                    tb_MailWorker.Text = selected.MailWorker;
                    //cb_OtdelName.SelectedItem = selected.OtdelName;
                }
                catch { }
                
            }
            else
            {
                try
                {
                    WorkerRegion selected = WorkersGrid.SelectedItem as WorkerRegion;
                    tb_Fio.Text = selected.Name;
                    tb_IP.Text = selected.Ip;
                    tb_Login.Text = selected.Login;
                    tb_TelOtdel.Text = selected.TelOtdel;
                    tb_Pass.Text = selected.Pass;
                    tb_User.Text = selected.User;
                    //cb_OtdelName.SelectedItem = selected.OtdelName;
                }
                catch { } 
            }
        }

        private void Image_MouseUp(object sender, MouseButtonEventArgs e)
        {
            string body = "1. Наименование подразделения (номер отдела): Отдел ГЦЖС №" + cb_OtdelNumb.SelectedItem + "\r\n" + "2. ФИО: " + tb_Fio.Text + "\r\n" +
            "3. Контактный телефон: " + tb_TelOtdel.Text + "\r\n" +
            "4. Адрес электронной почты: " + tb_MailWorker.Text + "\r\n" +
            "5. Описание проблемы: "; //это тело письма, записанное в переменную, ибо длинное

            if (rb_SendMailOutlook.IsChecked == true)
            {
                SendByOutlook(body, tb_MailTheme.Text, tb_MailTo.Text);
            }
            else
            {
                try
                {
                    SendMail(tb_SMTPserver.Text, tb_MailLogin.Text, tb_MailPass.Text, tb_MailTo.Text, tb_MailTheme.Text, body);
                }
                catch { MessageBox.Show("Проверьте настройки почты. Поля внутренней отправки и получателя не должны быть пустыми"); } 
            }
        }

        private void tb_Fio_KeyDown(object sender, KeyEventArgs e) //задает поиск в текстбоксе ФИО по нажатию энтера
        {
            if (e.Key == Key.Return)
            {
                SearchWorker();
            }
        }

        public void SendMail(string smtpServer, string from, string password, //это все целиком метод, используемый для отправки почты из программы
            string mailTo, string theme, string body, string attachFile = null)
        {
            try
            {
                MailMessage mail = new MailMessage();   //создаем объект письма
                mail.From = new MailAddress(from);      //задаем объекту письма адрес отправителя
                mail.To.Add(new MailAddress(mailTo));   //и получателя
                mail.Subject = theme;                   //тему письма
                mail.Body = body;                    //тело письма
                if (!string.IsNullOrEmpty(attachFile))
                    mail.Attachments.Add(new System.Net.Mail.Attachment(attachFile)); //если строка вложения не пуста, то добавляем в письмо вложение
                SmtpClient client = new SmtpClient();   //тут создаем объект подключения видимо
                client.Host = smtpServer;
                client.Port = 587;
                client.EnableSsl = true;
                client.Credentials = new NetworkCredential(from.Split('@')[0], password);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Send(mail);
                mail.Dispose();
            }
            catch (System.Exception o)
            {
                throw new System.Exception("Mail.Send: " + o.Message);
            }

            MessageBox.Show("Письмо отправлено");
        }

        public void SendByOutlook(string body, string theme, string mailTo)
        {
            Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application(); //Запуск приложения Outlook
            MailItem oMsg = (MailItem)oApp.CreateItem(OlItemType.olMailItem);           //Создание письма
            oMsg.Body = body;                                                           //тело письма
            //Outlook.Attachment oAttach = oMsg.Attachments.Add(@"D:\\мой*файл.txt");   //если нужно прикреплять файл
            oMsg.Subject = theme;                                                       //Тема письма
            Recipients oRecips = (Recipients)oMsg.Recipients;                           //даем объект получатели
            try
            {
                Recipient oRecip = (Recipient)oRecips.Add(mailTo);           // Change the recipient in the next line if necessary.
                oRecip.Resolve();
                oMsg.Display(); //отправка формы в письмо Outlook
                                // Clean up:
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }
            catch { MessageBox.Show("Ошибка передачи данных в Outlook. Проверьте корректность введенных данных."); }

            MessageBox.Show("Данные перенаправлены в Outlook");
        }

        private void bt_Radmin_Click(object sender, RoutedEventArgs e)
        {
            if (tb_IP.Text != "")
            {
                try
                {
                    Process.Start(tb_RadminPath.Text, "/connect:" + tb_IP.Text + ":4899");
                }
                catch (System.Exception)
                {

                    throw;
                }
            }
            else { MessageBox.Show("Не указан IP в поле ввода.", "Нет данных", MessageBoxButton.OK); }           
        }

        private void bt_IpScanner_Click(object sender, RoutedEventArgs e)
        {           
            try
            {
                Process.Start(tb_ScannerPath.Text);
            }
            catch (System.Exception)
            {

                throw;
            }
        }

        private void bt_RadminPath_Click(object sender, RoutedEventArgs e)      //Открывает диалог для выбор экзешника радмина в настройках
        {
            Microsoft.Win32.OpenFileDialog Radmin = new Microsoft.Win32.OpenFileDialog();
            Radmin.ShowDialog();
            tb_RadminPath.Text = Radmin.FileName;
        }

        private void bt_ScannerPath_Click(object sender, RoutedEventArgs e)     //Открывает диалог для выбор экзешника сканера
        {
            Microsoft.Win32.OpenFileDialog Scanner = new Microsoft.Win32.OpenFileDialog();
            Scanner.ShowDialog();
            tb_ScannerPath.Text = Scanner.FileName;
        }

        private void bt_Zimbra_Click(object sender, RoutedEventArgs e)
        {
            try {Process.Start("explorer", @"\\win2k8fs\11.3. Отдел систем связи\ZIMBRA");}
            catch{ MessageBox.Show("Что-то пошло не так. Может сеть легла? А может папка больше не доступна. В общем либо попробуй еще раз, либо больше не тыкай.", "Неведомая фигня"); }
        }

        private void bt_Delete_Click(object sender, RoutedEventArgs e)      //кнопка удаления сотрудникарев
        {
            if (WorkersGrid.SelectedIndex != -1)
            {
                DeleteWorker();
            }
            else { MessageBox.Show("Не выбран сотрудник для удаления."); }            
        }

        private void bt_ClearFields_Click(object sender, RoutedEventArgs e) //кнопка "Очистить форму"
        {
            cb_OtdelName.SelectedIndex = -1;
            tb_Fio.Clear();
            tb_Login.Clear();
            tb_Pass.Clear();
            tb_TelOtdel.Clear();
            tb_User.Clear();
            tb_TelVnutr.Clear();
            tb_IP.Clear();
            tb_MailWorker.Clear();
            tb_Comment.Clear();
            tb_roomNumber.Clear();
        }

        public void NewWorkerWrite()                //метод создания нового сотрудника в базе
        {
            if (rb_Center.IsChecked == true)
            {
                WorkerCenter newWorker = new WorkerCenter(0, tb_Fio.Text, tb_User.Text, tb_Login.Text, tb_Pass.Text, cb_OtdelName.SelectedItem.ToString(), null, tb_TelOtdel.Text, tb_TelVnutr.Text, tb_IP.Text, tb_MailWorker.Text, tb_Comment.Text, Convert.ToString(tb_roomNumber.Text));
                if (MessageBox.Show("Создать новую запись в базе:\nИмя - " + newWorker.Name + "\nUser - " + newWorker.User + "\nОтдел - " + newWorker.OtdelName, "Новая запись", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    try
                    {
                        SqlConnection connection = new SqlConnection(connectionString);
                        SqlCommand command = new SqlCommand("create_NewWorkerCenter", connection);
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("Fio", newWorker.Name);
                        command.Parameters.AddWithValue("Otdel", newWorker.OtdelName);
                        command.Parameters.AddWithValue("User", newWorker.User);
                        command.Parameters.AddWithValue("Login", newWorker.Login);
                        command.Parameters.AddWithValue("Pass", newWorker.Pass);
                        command.Parameters.AddWithValue("TelOtdel", newWorker.TelOtdel);
                        command.Parameters.AddWithValue("TelVnutr", newWorker.TelVnutr);
                        command.Parameters.AddWithValue("IP", newWorker.Ip);
                        command.Parameters.AddWithValue("Mail", tb_MailWorker.Text);
                        command.Parameters.AddWithValue("Komment", tb_Comment.Text);
                        command.Parameters.AddWithValue("KabNumb", tb_roomNumber.Text);

                        connection.Open();
                        command.ExecuteNonQuery();
                        connection.Close();
                        MessageBox.Show("Сотрудник создан");
                    }
                    catch
                    {
                        MessageBox.Show("Траблы");
                    }
                }                                             
            }
        }

        public void DeleteWorker()                  //метод удаления работника из базы
        {
            if (rb_Center.IsChecked == true)
            {
                WorkerCenter deleting = WorkersGrid.SelectedItem as WorkerCenter;
                if (MessageBox.Show("Вы уверены что хотите удалить запись: " + deleting.Name + "?", "Удаление записи из базы", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    try
                    {
                        WorkersCenter.Remove(deleting);
                        SqlConnection connection = new SqlConnection(connectionString); //создаем подключение
                        SqlCommand command = new SqlCommand("delete_WorkerCenter", connection); //выбираем объект из базы
                        command.CommandType = CommandType.StoredProcedure; //определяем тип, что это хранимая процедура
                        command.Parameters.AddWithValue("id", deleting.Id); //первое куда, второе что из кода
                        connection.Open(); //открываем соединение
                        command.ExecuteNonQuery(); //выполнение хранимой процедуры
                        connection.Close();
                        MessageBox.Show("Запись удалена");
                    }
                    catch{ MessageBox.Show("И еще какие-то проблемы"); }                    
                }
            }
        }

        public void WriteData()                     //запись данных в базу
        {
            if (rb_Center.IsChecked == true)
            {
                var rewritered = WorkersGrid.SelectedItem as WorkerCenter;
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand command = new SqlCommand("rewrite_DataInCenter", connection);
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("id", rewritered.Id);
                    command.Parameters.AddWithValue("Fio", tb_Fio.Text);
                    command.Parameters.AddWithValue("User", tb_User.Text);
                    command.Parameters.AddWithValue("Login", tb_Login.Text);
                    command.Parameters.AddWithValue("Pass", tb_Pass.Text);
                    command.Parameters.AddWithValue("TelOtdel", tb_TelOtdel.Text);
                    command.Parameters.AddWithValue("TelVnutr", tb_TelVnutr.Text);
                    command.Parameters.AddWithValue("IP", tb_IP.Text);
                    command.Parameters.AddWithValue("Mail", tb_MailWorker.Text);
                    command.Parameters.AddWithValue("Komment", tb_Comment.Text);
                    command.Parameters.AddWithValue("KabNumb", tb_roomNumber.Text);
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
            }
            else
            {
                var rewritered = WorkersGrid.SelectedItem as WorkerRegion;
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand command = new SqlCommand("rewrite_DataInRegion", connection);
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("id", rewritered.Id);
                    command.Parameters.AddWithValue("Fio", tb_Fio.Text);
                    command.Parameters.AddWithValue("User", tb_User.Text);
                    command.Parameters.AddWithValue("Login", tb_Login.Text);
                    command.Parameters.AddWithValue("Pass", tb_Pass.Text);
                    command.Parameters.AddWithValue("TelOtdel", tb_TelOtdel.Text);
                    command.Parameters.AddWithValue("IP", tb_IP.Text);
                    command.Parameters.AddWithValue("Mail", tb_MailWorker.Text);
                    command.Parameters.AddWithValue("Komment", tb_Comment.Text);
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
            }
            MessageBox.Show("Данные записаны в базу.");
        }

        public void RefreshDataInGrid()             //костыль обновления данных
        {
            var x = cb_OtdelNumb.SelectedIndex;     //корявое решение проблемы с обновлением данных в WorkersGrid
            var y = WorkersGrid.SelectedIndex;      //не делайте так!!!!
            cb_OtdelNumb.SelectedIndex = -1;
            cb_OtdelNumb.SelectedIndex = x;
            WorkersGrid.SelectedIndex = -1;
            WorkersGrid.SelectedIndex = y;
        }
    }
}
