using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
using System.Text.RegularExpressions;
using System.Collections;
using System.Collections.Specialized;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace OutlookMeetingAdd
{
    public partial class MeetingSchedule : Form
    {
        public MeetingSchedule()
        {
            InitializeComponent();   
        }

        public DateTime start;
        public DateTime end;
        private int Global = 0;
        private List<integer> list=new List<integer>(); 




        public Outlook.AppointmentItem item;
        public ExchangeService service;


        ///////////////////////////////////////Paint datagridview1 index//////////////////////////////////////////////////////////////////////////////////
        #region
        /* private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView1.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 4);
            }
        }
        */
        #endregion
        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            DataGridView dg = (DataGridView)sender;
            Outlook.Application app = new Outlook.Application();


            if (app.ActiveWindow() is Outlook._Inspector)
            {
                Outlook.Inspector inspector = app.ActiveInspector();
                if (inspector.CurrentItem is Outlook.AppointmentItem)
                {
                    Outlook.AppointmentItem str = inspector.CurrentItem;
                    if (dg.Columns[e.ColumnIndex].Name == "MeetingRoomLocation")
                    {
                        if (str != null)
                        {
                            
                            str.MeetingStatus = Microsoft.Office.Interop.Outlook.OlMeetingStatus.olMeeting;
                            str.Start = Convert.ToDateTime(dg.Rows[e.RowIndex].Cells[0].Value);
                            str.End = str.Start.AddHours(end.Subtract(start).Hours).AddMinutes(end.Subtract(start).Minutes);
                            str.Location = dg.Rows[e.RowIndex].Cells[1].Value + "@ericsson.com";
                            Outlook.Recipient recipient = str.Recipients.Add(dg.Rows[e.RowIndex].Cells[1].Value+"@ericsson.com");
                        }
                        this.Close();
                    }
                }            
            }
           
        }
        

        private void MeetingSchedule_Load(object sender, EventArgs e)
        {
            int index=0;

            comboBox2.SelectedIndex = 2;
            comboBox2.Enabled = false;

            if (DateTime.Compare(DateTime.Now, new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, 0, 0)) > 0 && DateTime.Compare(DateTime.Now, new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, 30, 0)) < 0)
            {
                index = 2 * DateTime.Now.Hour + 1;
                this.comboBox1.SelectedIndex =index;
            }
            if (DateTime.Compare(DateTime.Now, new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, 30, 0)) > 0 && DateTime.Compare(DateTime.Now, new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.AddHours(1).Hour, 0, 0)) < 0)
            {
                index = 2 * DateTime.Now.Hour + 2;
                this.comboBox1.SelectedIndex = index;
            
            }
            this.comboBox3.SelectedIndex = index + 1;


            /////////////////////////////////连接exchange服务器/////////////////////////////////////////////////////////////////
            #region
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
            //以windows账户用户名和密码登陆
            //service.Credentials = new NetworkCredential("ezhgyon", "zyc&900916", "ericsson");
            //默认以window用户名密码登陆
            service.UseDefaultCredentials = true;
            //自动获取邮箱URL
            //service.AutodiscoverUrl("yongchan.zhang@ericsson.com", RedirectionUrlValidationCallback);
            //手动设置exchange服务器地址
            service.Url = new Uri("https://mail-ao.internal.ericsson.com/EWS/Exchange.asmx");
            #endregion
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = comboBox1.SelectedIndex;
            this.comboBox3.SelectedIndex = index + 1;
         }


        public int Select_MeetingRoom(DateTime Start, DateTime End)
        {

            
            ////////////////////////////////Get Shanghai MeetingRoom Location///////////////////////////////////////////////////
            #region
            string myRoomList = "PDLRDCMEET@ex1.eapac.ericsson.se";
            ExpandGroupResults myRoomLists = service.ExpandGroup(myRoomList);
            System.Collections.ObjectModel.Collection<EmailAddress> roomAddresses = myRoomLists.Members;
            List<AttendeeInfo> attendees = new List<AttendeeInfo>();
            foreach (EmailAddress address in roomAddresses)
            {
                attendees.Add(new AttendeeInfo()
                {
                    SmtpAddress = address.Address,
                });
            }
            #endregion
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////




            /////////////////////////get the MeetingRoom Location information in a specified period////////////////////////////
            #region
            /*
                foreach (EmailAddress address in roomAddresses)
                {
                 CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
                 CalendarView calendarView = new CalendarView(DateTime.Now, DateTime.Now.AddHours(8));

                 /////////////////////////// 决定对象是否加载(very important)
                 calendarView.PropertySet = new PropertySet(AppointmentSchema.IsMeeting, AppointmentSchema.Start, AppointmentSchema.End);

                 FolderId folderID = new FolderId(WellKnownFolderName.Calendar, address.Address);
                 FindItemsResults<Appointment> roomAppts = service.FindAppointments(folderID, calendarView);
                
                 if (roomAppts.Items.Count > 0)
                 {
                     foreach (Appointment appt in roomAppts)
                     {
                        // Console.WriteLine("{0} - {1} : {2}", appt.Start, appt.End, appt.IsMeeting);
                      textBox2.Text += ("Appointments for Room"+address.Address+"\r\n"
                      + "start:" + appt.Start + "-" + "end:" + appt.End + "IsMeeting" + appt.IsMeeting + "\r\n" + "\r\n");
                     }
                 }
                }
                 */
            #endregion
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////




            //////////////////////display the meetingroom free/busy status//////////////////////////////////////////////////////  
            #region
            AvailabilityOptions myOptions = new AvailabilityOptions();
            myOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusyMerged;
            // Return a set of free/busy times.
            GetUserAvailabilityResults freeBusyResults = service.GetUserAvailability(attendees,
                                                                            new TimeWindow(DateTime.Now, DateTime.Now.AddDays(7)),
                                                                            AvailabilityData.FreeBusy, myOptions);


            start = Start;
            end = End;

        

            int count = 0;
            int test = 0;             /////调试计数
            bool flag = true;        //////////////////////////判断输入时间是否处于会议时间中
            bool Flag = true;         /////////////////////////判断年月日是否相等
            List<string> lst = new List<string>();
            foreach (AttendeeAvailability availability in freeBusyResults.AttendeesAvailability)
            {

                foreach (CalendarEvent calendarItem in availability.CalendarEvents)
                {
                    if (DateTime.Compare(start, end) != 0)    ////////////////////时间不相等
                    {
                        if (DateTime.Compare(end.Date, start.Date) == 0)  /////////////日期相等
                        {
                            if ((DateTime.Compare(calendarItem.StartTime, start) <= 0 && DateTime.Compare(calendarItem.EndTime, end) >= 0) || (DateTime.Compare(calendarItem.EndTime, end) < 0 && DateTime.Compare(calendarItem.EndTime, start) > 0) || (DateTime.Compare(calendarItem.StartTime, end) < 0 && DateTime.Compare(calendarItem.StartTime, start) > 0))
                            {
                                flag = false;
                                break;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please select the time not more than a day");
                            Flag = false;
                            return 0;
                        }
                    } 
                    else
                    {
                        MessageBox.Show("Start time must be greater than the end time");
                        return 0;
                    }
                }
                if (!Flag)
                    break;
                if (flag)
                {
                    lst.Add(attendees[count].SmtpAddress);
                    test++;
                }
                flag = true;
                count++;
            }
            #endregion
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



            //////////////////////Find the MeetingRoom belongs to E//////////////////////////////////////////////////////////////
            #region
            string pattern = @"(?<=\D\.)\d[?=.\d+]";
            int count_LKE = 0;
            //////conf.cn.sh.0a.03.0308.09.aries@ericsson.com
            /////Conf.CN.SH.LKE.03.E312.20.TaiMountain@ericsson.com
            Regex rgx = new Regex(pattern);
            List<string> choosing_Buiding_location = new List<string>();

            foreach (string address in lst)
            {
                if (Regex.IsMatch(address, "LKE", RegexOptions.IgnoreCase))
                {
                    choosing_Buiding_location.Add(address);
                    count_LKE++;
                }
            }

            if (count_LKE == 0)
            {
                MessageBox.Show("No MeetingRooms avalibility,Please reselect the MeetingTime!");
                return 0;
            }

            #endregion 
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



            ////////////////////////////////////Sort the floor//////////////////////////////////////////////////////////////////
            #region
            NameValueCollection location       = new NameValueCollection();
            NameValueCollection location_filter= new NameValueCollection();

            foreach (string lc in choosing_Buiding_location)
            {
                foreach (Match match in rgx.Matches(lc))
                    location.Add(match.Value, lc);
            }


            int t=0;
            string temp;
            string pattern_1 = @"\d{1,}";
            Regex Rg = new Regex(pattern_1);
            string[] sortkeys = location.AllKeys;
            for (int i = 0; i < location.Count-1; i++)
            {
                t = i;
                for (int j = i + 1; j < location.Count; j++)
                {
                        if(Math.Abs((int.Parse(sortkeys[j])-int.Parse(Rg.Match(comboBox2.Text).Value)))<Math.Abs((int.Parse(sortkeys[t])-int.Parse(Rg.Match(comboBox2.Text).Value))))
                            t = j;
                }
                if (t != i)
                {
                    temp = sortkeys[i];
                    sortkeys[i] = sortkeys[t];
                    sortkeys[t] = temp;
                }
            }

            for (int i = 0; i < location.Count; i++)
            {
                location_filter.Add(sortkeys[i], location[sortkeys[i]]);   
            }

           
          
            int iter = 0;
            bool flag_T = true;
            string pattern_2 = @"^[^@]+";
            Regex st = new Regex(pattern_2);
            Regex reg = new Regex(",");


            
            
            foreach (string str in location_filter.Keys)
            {
                string[] svec = reg.Split(location_filter[str]);

                foreach (string s in svec)
                {

                    integer suggest_information = new integer(Start.ToString(), st.Match(s).Value, int.Parse(Rg.Match(s).Value).ToString());
                    list.Add(suggest_information);
                    iter++;
                    if (iter > 2)
                    {
                        flag_T = false;
                        break;
                    }
                }
                if (!flag_T)
                    break;
            }


            return 1;
        }
            #endregion
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////



            ///////////////////////Repaint Datagridview1///////////////////////////////////////////////////////////////////////
 
        private void MergeCellInOneColumn(DataGridView dgv, List<int> columnIndexList, DataGridViewCellPaintingEventArgs e)
        {
            if (columnIndexList.Contains(e.ColumnIndex) && e.RowIndex != -1)
            {
                Brush gridBrush = new SolidBrush(dgv.GridColor);
                Brush backBrush = new SolidBrush(e.CellStyle.BackColor);
                Pen gridLinePen = new Pen(gridBrush);

                //擦除
                e.Graphics.FillRectangle(backBrush, e.CellBounds);

                //画右边线
                e.Graphics.DrawLine(gridLinePen,
                   e.CellBounds.Right - 1,
                   e.CellBounds.Top,
                   e.CellBounds.Right - 1,
                   e.CellBounds.Bottom - 1);

                //画底边线
                if ((e.RowIndex < dgv.Rows.Count - 1 && dgv.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value.ToString() != e.Value.ToString()) ||
                    e.RowIndex == dgv.Rows.Count - 1)
                {
                    e.Graphics.DrawLine(gridLinePen,
                        e.CellBounds.Left,
                        e.CellBounds.Bottom - 1,
                        e.CellBounds.Right - 1,
                        e.CellBounds.Bottom - 1);
                }

                //写文本
                if (e.RowIndex == 0 || dgv.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.ToString() != e.Value.ToString())
                {
                    e.Graphics.DrawString((String)e.Value, e.CellStyle.Font,
                        Brushes.Black, e.CellBounds.X + 2,
                        e.CellBounds.Y +5, StringFormat.GenericDefault);
                }

                e.Handled = true;
            }
        }
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            List<int> indexs = new List<int>() { 0, 1 };
            MergeCellInOneColumn(dataGridView1, indexs, e);
        }
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////



        private void button2_Click(object sender, EventArgs e)
        {
            ////////////////////////////////////get attendees Email////////////////////////////////////////////////////////////////
            #region
            Outlook.Application app = new Outlook.Application();
            List<AttendeeInfo> attendees = new List<AttendeeInfo>();
            string pattern_3 = " ";
            string replacement = ".";
            Regex cheat = new Regex(pattern_3);

            if (app.ActiveWindow() is Outlook._Inspector)
            {
                Outlook.Inspector inspector = app.ActiveInspector();
                if (inspector.CurrentItem is Outlook.AppointmentItem)
                {
                    item = inspector.CurrentItem;
                    foreach (Outlook.Recipient acquird in item.Recipients)
                    {
                        attendees.Add(new AttendeeInfo()
                        {
                            SmtpAddress = cheat.Replace(acquird.Name, replacement) + "@ericsson.com",
                            // AttendeeType = MeetingAttendeeType.Required
                        });
                    }


                }
                //////////////////////测试邮箱名是否被正确加载////////////////////////////
                /*EmailMessage message = new EmailMessage(service);

                // Set properties on the email message.
                message.Subject = "Company Soccer Team";
                message.Body = "Are you interested in joining?";
                message.ToRecipients.Add(attendees[0].SmtpAddress);

                // Send the email message and save a copy.
                // This method call results in a CreateItem call to EWS.
                message.Send();*/
            }
            #endregion
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



            ////////////////////////////////////datagrid1 initial////////////////////////////////////////////////////////////////
            #region
            if (Global != 0)
            {
                dataGridView1.Rows.Clear();
            }
            else
            {
                ///////////////////////////////////////datagridview1设置///////////////////////////////////////////////////////////////////
                DataGridViewTextBoxColumn ct = new DataGridViewTextBoxColumn();
                ct.Name = "Suggested Meeting Time";
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                ct.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridView1.Columns.Add(ct);


                DataGridViewButtonColumn column_dg1 = new DataGridViewButtonColumn();
                column_dg1.Name = "MeetingRoomLocation";
                column_dg1.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                column_dg1.UseColumnTextForButtonValue = true;
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
              //  dataGridView1.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
                dataGridView1.Columns.Add(column_dg1);

                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                col.Name = "Floor";
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns.Add(col);

                dataGridView1.RowHeadersVisible = false;
               
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            }
           
            #endregion
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


            ///////////////////////////////////////get the attendes's information////////////////////////////////////////////////
            #region
            AvailabilityOptions availabilityOptions = new AvailabilityOptions();
            if (Global == 0)
            {
                var dateSpan = item.End.Subtract(item.Start);
                availabilityOptions.MeetingDuration = dateSpan.Hours * 60 + dateSpan.Minutes;
                start = item.Start;
                end = item.End;
              
            }
            else
            {
                var dateSpan = new DateTime(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day, Convert.ToDateTime(comboBox3.SelectedItem).Hour, Convert.ToDateTime(comboBox3.SelectedItem).Minute, 0).Subtract(new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day, Convert.ToDateTime(comboBox1.SelectedItem).Hour, Convert.ToDateTime(comboBox1.SelectedItem).Minute, 0));
               // MessageBox.Show(dateSpan.Minutes.ToString());
                availabilityOptions.MeetingDuration = dateSpan.Hours * 60 + dateSpan.Minutes;

                start = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day, Convert.ToDateTime(comboBox1.SelectedItem).Hour, Convert.ToDateTime(comboBox1.SelectedItem).Minute, 0);
                end = start.Add(TimeSpan.FromMinutes(availabilityOptions.MeetingDuration));
            }
            
            availabilityOptions.MaximumNonWorkHoursSuggestionsPerDay = 0;
            availabilityOptions.MaximumSuggestionsPerDay = 20;
            availabilityOptions.GoodSuggestionThreshold = 49;
            availabilityOptions.MinimumSuggestionQuality = SuggestionQuality.Excellent;
            availabilityOptions.DetailedSuggestionsWindow = new TimeWindow(item.Start, item.Start.AddDays(1));
            availabilityOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;

            GetUserAvailabilityResults results = service.GetUserAvailability(attendees,
                                                                   availabilityOptions.DetailedSuggestionsWindow,
                                                                   AvailabilityData.FreeBusyAndSuggestions,
                                                                   availabilityOptions);



            bool suggestion_flag=true;
            bool conflict_flag = true;
            List<DateTime> str = new List<DateTime>();
            int counts=0;
            string conflict_number=null;
           
            Regex Rg = new Regex("^[^@]+");
            foreach (AttendeeAvailability availability in results.AttendeesAvailability)
            {
                foreach (CalendarEvent calEvent in availability.CalendarEvents)
                {
                    
                    if ((DateTime.Compare(calEvent.StartTime, start) <= 0 && DateTime.Compare(calEvent.EndTime, end) >= 0) || (DateTime.Compare(calEvent.EndTime, end) < 0 && DateTime.Compare(calEvent.EndTime,start) > 0) || (DateTime.Compare(calEvent.StartTime, end) < 0 && DateTime.Compare(calEvent.StartTime, start) > 0))
                    {
                       // MessageBox.Show("Conflict!");
                        conflict_number += attendees[counts].SmtpAddress+" ";
                        conflict_flag = false;
                    }      
                }
                counts++;
            }

            if (!conflict_flag)
            {
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                string message = "Conflict:" + Rg.Match(conflict_number).Value.ToString();
                string caption = "Warning";
                DialogResult result;


                // Displays the MessageBox.
                result = MessageBox.Show(message, caption, buttons);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    suggestion_flag = false;
                }
            }

            if (!suggestion_flag)
            {
                TimeSpan different;
                foreach (Suggestion suggestion in results.Suggestions)
                {
                    foreach (TimeSuggestion timeSuggestion in suggestion.TimeSuggestions)
                    {
                        if (timeSuggestion.MeetingTime.Hour >= 9)
                        {
                            str.Add(timeSuggestion.MeetingTime);
                        }
                    }
                }

                int k = 0;
                DateTime temp;
                for (int s = 0; s < str.Count - 1; s++)
                {
                    k = s;
                    for (int j = s + 1; j < str.Count; j++)
                    {
                        different = str[j].Subtract(item.Start);
                        if (Math.Abs(different.Hours * 60 + different.Minutes) < Math.Abs(str[k].Subtract(item.Start).Hours * 60 + str[k].Subtract(item.Start).Minutes))
                            k = j;
                    }
                    if (k != s)
                    {
                        temp = str[k];
                        str[k] = str[s];
                        str[s] = temp;
                    }
                }
            }
            else
            {
                if (Global == 0)
                    str.Add(item.Start);
                else
                {
                    str.Add(new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day, Convert.ToDateTime(comboBox1.SelectedItem).Hour, Convert.ToDateTime(comboBox1.SelectedItem).Minute, 0));
                }
            }



            for (int i = 0; i < str.Count; i++)
            {
                if (Select_MeetingRoom(str[i], str[i].Add(TimeSpan.FromMinutes(availabilityOptions.MeetingDuration))) == 0)
                    break;
                if (i > 1)
                    break;
            }


            foreach (integer information in list)
            {
                DataGridViewRow row = new DataGridViewRow();

                DataGridViewTextBoxCell textedit = new DataGridViewTextBoxCell();
                textedit.Value = information.MeetingTime;
                row.Cells.Add(textedit);

                DataGridViewButtonCell btnEdit = new DataGridViewButtonCell();
                btnEdit.Value =information.MeetingRoomLocation;
               // btnEdit.Tag = information.MeetingRoomLocation + "@ericsson.com";
                btnEdit.FlatStyle = FlatStyle.Popup;
                row.Cells.Add(btnEdit);

                DataGridViewTextBoxCell textEdit = new DataGridViewTextBoxCell();
                textEdit.Value = information.Floor;
                row.Cells.Add(textEdit);


                dataGridView1.Rows.Add(row);
            }
            list.Clear();
            Global++;


            #endregion
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker1.Value;
        }
    }
}