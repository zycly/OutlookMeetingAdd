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

        private DateTime start;
        private DateTime end;
        private int Global_1 = 0;
        private int Global_2 = 0;
        public Outlook.AppointmentItem item;
        public ExchangeService service;

 
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView1.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 4);
            }
        }

        #region
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
                            str.Location = dg.Rows[e.RowIndex].Cells[0].Tag.ToString();
                            str.Start = new DateTime(start.Year, start.Month, start.Day, start.Hour, start.Minute, 0);
                            str.End = new DateTime(end.Year, end.Month, end.Day, end.Hour, end.Minute, 0);
                            Outlook.Recipient recipient = str.Recipients.Add(dg.Rows[e.RowIndex].Cells[0].Tag.ToString().Trim());
                        }
                        this.Close();
                    }
                }            
            }
           
        }
        #endregion
     


        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView2.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 4);
            }
        }


        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

                ////////////////////////////////获取上海地区会议室的信息////////////////////////////////////////////////////////////
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
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                ////////////////////////////////datagrid1初始化////////////////////////////////////////////////////////////////////
                #region
                if (Global_1 != 0)
                {
                    dataGridView1.Rows.Clear();
                    //dataGridView2.Rows.Clear();
                }
                else
                {
                    ///////////////////////////////////////datagridview1设置///////////////////////////////////////////////////////////////////
                    DataGridViewButtonColumn column_dg1 = new DataGridViewButtonColumn();
                    column_dg1.Name = "MeetingRoomLocation";
                    column_dg1.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    column_dg1.UseColumnTextForButtonValue = true;
                    dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView1.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
                    dataGridView1.Columns.Add(column_dg1);
                    DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                    col.Name = "Floor";
                    dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView1.Columns.Add(col);
                    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                }
                Global_1++;
                #endregion
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                /////////////////////////获取某一时间段内的所有会议数据信息/////////////////////////////////////////////////////////
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


                //////////////////////获取会议室信息/////////////////////////////////////////////////////////////////////////////

                AvailabilityOptions myOptions = new AvailabilityOptions();
                myOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusyMerged;
                // Return a set of free/busy times.
                GetUserAvailabilityResults freeBusyResults = service.GetUserAvailability(attendees,
                                                                                new TimeWindow(DateTime.Now, DateTime.Now.AddDays(7)),
                                                                                AvailabilityData.FreeBusy, myOptions);


                DataGridView dg = (DataGridView)sender;

                if (dg.Columns[e.ColumnIndex].Name == "Suggested Meeting Time")
                {
                    start =Convert.ToDateTime(dg.Rows[e.RowIndex].Cells[0].Value);
                    end = Convert.ToDateTime(dg.Rows[e.RowIndex].Cells[0].Tag);
                }
                MessageBox.Show(start.ToString()+"\r\n");
                MessageBox.Show(end.ToString());
                #region



                ///////////////////////显示会议室free/busy////////////////////////////////////////////////   

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
                                break;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Start time must be greater than the end time");
                            return;
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


                ///////////////////////////////////匹配办公地点/////////////////////////////////////
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
                ////////////////////////////////////匹配楼层/////////////////////////////////////////


                if (count_LKE == 0)
                {
                    MessageBox.Show("No MeetingRooms avalibility,Please reselect the MeetingTime!");
                    return;
                }

                NameValueCollection location = new NameValueCollection();
                NameValueCollection locat = new NameValueCollection();
                foreach (string lc in choosing_Buiding_location)
                {
                    foreach (Match match in rgx.Matches(lc))
                        location.Add(match.Value, lc);

                }


                int diff;
                int j = 0;
                string[] sortkeys = location.AllKeys;
                int[] sort = new int[location.Count];
                int[] flag_l = new int[location.Count];




                /*Array.Sort(sortkeys);
                foreach (string key in sortkeys)
                {
                    locat.Add(key, location[key]);
                }
                foreach (string str in locat.Keys)
                {
                    textBox3.Text += (str.ToString() + "\r\n");
                    string[] svec = reg.Split(locat[str]);
                    foreach (string s in svec)
                    {
                        textBox3.Text += (s.ToString() + "\r\n");
                    }
                    textBox3.Text+=("\r\n");
                }*/


                string pattern_1 = @"\d{1,}";
                Regex Rg = new Regex(pattern_1);
                try
                {
                    foreach (string t in sortkeys)
                    {
                        diff = int.Parse(t) - int.Parse(Rg.Match(comboBox2.Text).Value);
                        if (diff < 0)
                        {
                            diff = diff * (-1);
                            flag_l[j] = 1;
                        }
                        sort[j] = diff;
                        j++;
                    }
                    int temp;
                    for (int i = 0; i < location.Count - 1; i++)
                    {
                        for (int k = 0; k < location.Count - 1 - i; k++)
                        {
                            if (sort[k] > sort[k + 1])
                            {
                                temp = sort[k];
                                sort[k] = sort[k + 1];
                                sort[k + 1] = temp;


                                temp = flag_l[k];
                                flag_l[k] = flag_l[k + 1];
                                flag_l[k + 1] = temp;
                            }
                        }
                    }
                }
                catch (Exception e3)
                {
                    MessageBox.Show("Warning:" + "\r\n" + "Please check your location columm and then try again!");
                    return;
                }
                /////////////////////////////////////////////////////////////////////

                int num;
                string change;
                for (int i = 0; i < location.Count; i++)
                {
                    if (flag_l[i] == 0)
                        num = sort[i] + int.Parse(Rg.Match(comboBox2.SelectedItem.ToString()).Value);
                    else
                    {
                        num = sort[i] * (-1) + int.Parse(Rg.Match(comboBox2.SelectedItem.ToString()).Value);
                        flag_l[i] = 0;
                    }

                    if (num > 0 && num < 10)
                    {
                        change = "0" + num.ToString();
                    }
                    else
                        change = num.ToString();

                    locat.Add(change, location[change]);
                }
                ////////////////////////////////////////////////////////////////////////////////////////////////


                int iter = 0;
                string pattern_2 = @"^[^@]+";
                Regex st = new Regex(pattern_2);
                Regex reg = new Regex(",");


                foreach (string str in locat.Keys)
                {
                    string[] svec = reg.Split(locat[str]);
                    foreach (string s in svec)
                    {
                        DataGridViewRow row = new DataGridViewRow();

                        DataGridViewButtonCell btnEdit = new DataGridViewButtonCell();
                        btnEdit.Value = st.Match(s).Value;
                        btnEdit.Tag = s;

                        btnEdit.FlatStyle = FlatStyle.Popup;
                        row.Cells.Add(btnEdit);

                        DataGridViewTextBoxCell textedit = new DataGridViewTextBoxCell();
                        textedit.Value = int.Parse(Rg.Match(s).Value).ToString();
                        row.Cells.Add(textedit);

                        dataGridView1.Rows.Add(row);
                        iter++;
                    }
                }
                /////////////////////////////////////////////////////////////////////////////////////////////////

        }



        private void MeetingSchedule_Load(object sender, EventArgs e)
        {
            int index=0;

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






        
        private void button2_Click(object sender, EventArgs e)
        {
           
            ////////////////////////////////////获取与会人时间信息//////////////////////////////////////////////////////////////
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



            ///////////////////////////////////////datagridview2设置////////////////////////////////////////////////////////////
            #region

            if (Global_2 != 0)
            {
                dataGridView1.Rows.Clear();
                dataGridView2.Rows.Clear();
            }
            else
            {
                ///////////////////////////////////////datagridview1设置///////////////////////////////////////////////////////////////////
          
                DataGridViewButtonColumn column_dg2 = new DataGridViewButtonColumn();
                column_dg2.Name = "Suggested Meeting Time";
                column_dg2.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                column_dg2.UseColumnTextForButtonValue = true;
                dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView2_RowPostPaint);
                dataGridView2.Columns.Add(column_dg2);
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            }
            Global_2++;
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            #endregion
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



            #region
            AvailabilityOptions availabilityOptions = new AvailabilityOptions();
            var dateSpan = item.End.Subtract(item.Start);
            availabilityOptions.MeetingDuration = dateSpan.Hours * 60 + dateSpan.Minutes;
            availabilityOptions.MaximumNonWorkHoursSuggestionsPerDay = 0;
            availabilityOptions.MaximumSuggestionsPerDay = 10;
            availabilityOptions.GoodSuggestionThreshold = 49;
            availabilityOptions.MinimumSuggestionQuality = SuggestionQuality.Good;
            availabilityOptions.DetailedSuggestionsWindow = new TimeWindow(item.Start, item.Start.AddDays(1));
            availabilityOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;

            GetUserAvailabilityResults results = service.GetUserAvailability(attendees,
                                                                   availabilityOptions.DetailedSuggestionsWindow,
                                                                   AvailabilityData.FreeBusyAndSuggestions,
                                                                   availabilityOptions);




            TimeSpan different;
            List<DateTime> str = new List<DateTime>();

            foreach (Suggestion suggestion in results.Suggestions)
            {
                foreach (TimeSuggestion timeSuggestion in suggestion.TimeSuggestions)
                {
                    if (timeSuggestion.MeetingTime.Hour >= 9)
                        str.Add(timeSuggestion.MeetingTime);
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
            #endregion


        
            for (int i = 0; i < 5; i++)
			{
			     DataGridViewRow row = new DataGridViewRow();
                 DataGridViewButtonCell btn = new DataGridViewButtonCell();
                 btn.Value =str[i].ToString();
                 btn.Tag = str[i].Add(TimeSpan.FromMinutes(availabilityOptions.MeetingDuration)).ToString();
                 btn.FlatStyle = FlatStyle.Popup;
                 row.Cells.Add(btn);
                 dataGridView2.Rows.Add(row);
			}

            str.Clear();

        }
    }
}