using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public const double Invalid = 1E+100;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                int IntStart;
                int IntEnd;
                int IntDatePos;

                const int IntMA1 = 5;
                const int IntMA2 = 20;
                const int IntMA3 = 60;
                const int IntMA4 = 100;
                const int IntMA5 = 300;

                var price = new ActiveMarket.Prices();
                var cal = new ActiveMarket.Calendar();

                IntStart = cal.DatePosition(DateTime.Parse(txtStart.Text),1);
                IntEnd = cal.DatePosition(DateTime.Parse(txtEnd.Text), -1);

                price.ReadBegin = IntStart;
                price.ReadEnd = IntEnd;

                price.Read(txtMeigara.Text);
                label1.Text = price.Name();

                if (IntStart < price.ReadBegin)
                {
                    IntStart = price.ReadBegin;
                }
                if(IntEnd > price.ReadEnd)
                {
                    IntEnd = price.ReadEnd;
                }

                DataTable dt = new DataTable("tname");
                dt.Columns.Add("Date", typeof(String));
                dt.Columns.Add("WeekNum", typeof(Decimal));
                dt.Columns.Add("Youbi", typeof(Decimal));
                dt.Columns.Add("Open",typeof(Decimal));
                dt.Columns.Add("High", typeof(Decimal));
                dt.Columns.Add("Low", typeof(Decimal));
                dt.Columns.Add("Close", typeof(Decimal));
                dt.Columns.Add("MA" + IntMA1.ToString(), typeof(Decimal));
                dt.Columns.Add("MA" + IntMA2.ToString(), typeof(Decimal));
                dt.Columns.Add("MA" + IntMA3.ToString(), typeof(Decimal));
                dt.Columns.Add("MA" + IntMA4.ToString(), typeof(Decimal));
                dt.Columns.Add("MA" + IntMA5.ToString(), typeof(Decimal));
                dt.Columns.Add("MA" + IntMA1.ToString() + "乖", typeof(Decimal));
                dt.Columns.Add("MA" + IntMA2.ToString() + "乖", typeof(Decimal));
                dt.Columns.Add("MA" + IntMA3.ToString() + "乖", typeof(Decimal));
                dt.Columns.Add("MA" + IntMA4.ToString() + "乖", typeof(Decimal));
                dt.Columns.Add("MA" + IntMA5.ToString() + "乖", typeof(Decimal));
                dt.Columns.Add("ZenSa", typeof(Decimal));
                dt.Columns.Add("ZenHi", typeof(Decimal));

                dt.AcceptChanges();

                for (IntDatePos  = IntStart; IntDatePos <= IntEnd; IntDatePos++)
                {
                    if (price.IsClosed(IntDatePos) != 0)
                    {
                        //dataGridView1.Rows.Add(cal.Date(IntDatePos).ToString("yyyy/MM/dd"),"","","","");
                    }
                    else if(price.Close(IntDatePos) == 0 | price.Close(IntDatePos) == Invalid)
                    {
                        //dataGridView1.Rows.Add(cal.Date(IntDatePos).ToString("yyyy/MM/dd"), "", "", "", "");
                    }
                    else
                    {
                        //dataGridView1.Rows.Add(cal.Date(IntDatePos).ToString("yyyy/MM/dd")
                        //    , 
                        //                        price.Open(IntDatePos),
                        //                        price.High(IntDatePos),
                        //                        price.Low(IntDatePos),
                        //                        price.Close(IntDatePos));

                        DataRow drNewRow = dt.NewRow();
                        drNewRow["Date"] = cal.Date(IntDatePos).ToString("yyyy/MM/dd");

                        DateTime calDate = new DateTime(cal.Date(IntDatePos).Year,cal.Date(IntDatePos).Month,cal.Date(IntDatePos).Day);
                        Calendar calender = CultureInfo.CurrentCulture.Calendar;

                        drNewRow["Youbi"] = (int)calDate.DayOfWeek;
                        drNewRow["WeekNum"] = calender.GetWeekOfYear(calDate,CalendarWeekRule.FirstDay,DayOfWeek.Sunday);
             
                        drNewRow["Open"] = price.Open(IntDatePos);
                        drNewRow["High"] = price.High(IntDatePos);
                        drNewRow["Low"] = price.Low(IntDatePos);
                        drNewRow["Close"] = price.Close(IntDatePos);

                        drNewRow["MA" + IntMA1.ToString()] = 0;
                        drNewRow["MA" + IntMA2.ToString()] = 0;
                        drNewRow["MA" + IntMA3.ToString()] = 0;
                        drNewRow["MA" + IntMA4.ToString()] = 0;
                        drNewRow["MA" + IntMA5.ToString()] = 0;

                        drNewRow["MA" + IntMA1.ToString() + "乖"] = 0;
                        drNewRow["MA" + IntMA2.ToString() + "乖"] = 0;
                        drNewRow["MA" + IntMA3.ToString() + "乖"] = 0;
                        drNewRow["MA" + IntMA4.ToString() + "乖"] = 0;
                        drNewRow["MA" + IntMA5.ToString() + "乖"] = 0;

                        drNewRow["ZenSa"] = 0;
                        drNewRow["ZenHi"] = 0;

                        dt.Rows.Add(drNewRow);
                    }
                }
                int i = 0;
                int j = 0;
                decimal decMA1 = 0;
                decimal decMA2 = 0;
                decimal decMA3 = 0;
                decimal decMA4 = 0;
                decimal decMA5 = 0;

                decimal decZenSa = 0;
                decimal decZenHi = 0;

                for (i=0;i < dt.Rows.Count; i++)
                {
                    if(i > 0)
                    {
                        decZenSa = decimal.Parse(dt.Rows[i]["Close"].ToString());
                        decZenSa -= decimal.Parse(dt.Rows[i-1]["Close"].ToString());
                        dt.Rows[i]["ZenSa"] = decZenSa;

                        decZenHi = decimal.Parse(dt.Rows[i]["Close"].ToString());
                        decZenHi = decZenHi / decimal.Parse(dt.Rows[i - 1]["Close"].ToString());
                        dt.Rows[i]["ZenHi"] = decZenHi;
                    }
                    if(i >= IntMA1)
                    {
                        decMA1 = 0;
                        for(j = 0;j < IntMA1; j++)
                        {
                            decMA1 += decimal.Parse(dt.Rows[i-j]["Close"].ToString());
                        }
                        decMA1 = decMA1 / IntMA1;
                        dt.Rows[i]["MA" + IntMA1.ToString()] = decMA1;

                        decMA1 = decimal.Parse(dt.Rows[i]["Close"].ToString()) - decMA1;
                        decMA1 = decMA1 / decimal.Parse(dt.Rows[i]["Close"].ToString());
                        dt.Rows[i]["MA" + IntMA1.ToString() + "乖"] = decMA1 * 100;
                    }
                    if (i >= IntMA2)
                    {
                        decMA2 = 0;
                        for (j = 0; j < IntMA2; j++)
                        {
                            decMA2 += decimal.Parse(dt.Rows[i-j]["Close"].ToString());
                        }
                        decMA2 = decMA2 / IntMA2;
                        dt.Rows[i]["MA" + IntMA2.ToString()] = decMA2;

                        decMA2 = decimal.Parse(dt.Rows[i]["Close"].ToString()) - decMA2;
                        decMA2 = decMA2 / decimal.Parse(dt.Rows[i]["Close"].ToString());
                        dt.Rows[i]["MA" + IntMA2.ToString() + "乖"] = decMA2 * 100;
                    }
                    if (i >= IntMA3)
                    {
                        decMA3 = 0;
                        for (j = 0; j < IntMA3; j++)
                        {
                            decMA3 += decimal.Parse(dt.Rows[i-j]["Close"].ToString());
                        }
                        decMA3 = decMA3 / IntMA3;
                        dt.Rows[i]["MA" + IntMA3.ToString()] = decMA3;

                        decMA3 = decimal.Parse(dt.Rows[i]["Close"].ToString()) - decMA3;
                        decMA3 = decMA3 / decimal.Parse(dt.Rows[i]["Close"].ToString());
                        dt.Rows[i]["MA" + IntMA3.ToString() + "乖"] = decMA3 * 100;
                    }
                    if (i >= IntMA4)
                    {
                        decMA4 = 0;
                        for (j = 0; j < IntMA4; j++)
                        {
                            decMA4 += decimal.Parse(dt.Rows[i-j]["Close"].ToString());
                        }
                        decMA4 = decMA4 / IntMA4;
                        dt.Rows[i]["MA" + IntMA4.ToString()] = decMA4;

                        decMA4 = decimal.Parse(dt.Rows[i]["Close"].ToString()) - decMA4;
                        decMA4 = decMA4 / decimal.Parse(dt.Rows[i]["Close"].ToString());
                        dt.Rows[i]["MA" + IntMA4.ToString() + "乖"] = decMA4 * 100;
                    }
                    if (i >= IntMA5)
                    {
                        decMA5 = 0;
                        for (j = 0; j < IntMA5; j++)
                        {
                            decMA5 += decimal.Parse(dt.Rows[i-j]["Close"].ToString());
                        }
                        decMA5 = decMA5 / IntMA5;
                        dt.Rows[i]["MA" + IntMA5.ToString()] = decMA5;

                        decMA5 = decimal.Parse(dt.Rows[i]["Close"].ToString()) - decMA5;
                        decMA5 = decMA5 / decimal.Parse(dt.Rows[i]["Close"].ToString());
                        dt.Rows[i]["MA" + IntMA5.ToString() + "乖"] = decMA5 * 100;
                    }
                }
                dataGridView1.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }




        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(dataGridView1.GetClipboardContent());
        }
    }
}
