using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Timers;
using System.Drawing.Printing;
using System.Drawing.Text;
using Novacode;

namespace LoanerTracking
{
    public partial class Form1 : Form
    {
        private ListViewColumnSorter lvwColumnSorter;

        List<string> listA = new List<string>();
        List<string> listB = new List<string>();
        List<string> listC = new List<string>();
        List<string> listD = new List<string>();
        List<string> listE = new List<string>();
        List<string> listF = new List<string>();
        List<string> listG = new List<string>();
        List<string> listH = new List<string>();
        List<string> listI = new List<string>();
        List<string> listJ = new List<string>();
        List<string> listK = new List<string>();
        List<string> listL = new List<string>();
        List<string> listM = new List<string>();
        List<string> listN = new List<string>();
        List<string> listO = new List<string>();

        List<string> listNumber = new List<string>();
        List<string> listAccount = new List<string>();
        List<string> listCustName = new List<string>();
        List<string> listContact = new List<string>();
        List<string> listPhone = new List<string>();
        List<string> listEmail = new List<string>();

        List<string> Users = new List<string>();
        List<string> PW = new List<string>();
        List<string> Access = new List<string>();

        
        List<List<string>> MasterList = new List<List<string>>();
        List<List<string>> ContactList = new List<List<string>>();

        int filter = 0;
        int login = 0;
        int currentindex = 0;
        bool logbuttpushed = false;
        string currentuser = "";
        int permanent = 0;

        public Form1()
        {
            InitializeComponent();

            // Create an instance of a ListView column sorter and assign it 
            // to the ListView control.
            lvwColumnSorter = new ListViewColumnSorter();
            this.listFields.ListViewItemSorter = lvwColumnSorter;
        }

        

        private void Form1_Load(object sender, EventArgs e)
        {
            UpdateCurrentStatus();
            UpdateCustomerList();
            LoadUsers();
            

            MasterList.Clear();
            MasterList.Add(listA);
            MasterList.Add(listB);
            MasterList.Add(listC);
            MasterList.Add(listD);
            MasterList.Add(listE);
            MasterList.Add(listF);
            MasterList.Add(listG);
            MasterList.Add(listH);
            MasterList.Add(listI);
            MasterList.Add(listJ);
            MasterList.Add(listK);
            MasterList.Add(listL);
            MasterList.Add(listM);
            MasterList.Add(listN);
            MasterList.Add(listO);

            ContactList.Add(listNumber);
            ContactList.Add(listAccount);
            ContactList.Add(listCustName);
            ContactList.Add(listContact);
            ContactList.Add(listPhone);
            ContactList.Add(listEmail);

            AutoLogin();

            PopulateLists();

            foreach (ColumnHeader ch in this.listFields.Columns)
            {
                ch.Width = -2;
            }
        }

        private void AutoLogin()
        {
            try
            {
                string loginfile = @"C:\Users\Public\loginfile";
                var reader = new StreamReader(loginfile);

                string username = reader.ReadLine();
                string password = reader.ReadLine();
            reader.Close();

                if (username == "")
                {
                    return;
                }

                UserLogin(username, password);
            }
            catch
            {
                //do nothing
            }
        }

        private void PopulateLists()
        {
            listFields.Items.Clear();

            for (int i = 0; i < MasterList[0].Count; i++)
            {
                ListViewItem lvi = new ListViewItem(MasterList[0][i]);
                for (int j = 1; j < MasterList.Count; j++)
                {
                    
                    if (j < 6 || j > 8)
                    {
                        lvi.SubItems.Add(MasterList[j][i]);
                    }
                    else
                    {
                        if (MasterList[j][i] != "")
                        {
                            DateTime parsedDate;
                            string[] formatstring = {"d/m/yyyy", "dd/m/yyyy", "d/mm/yyyy","dd/mm/yyyy"};

                            DateTime.TryParseExact(MasterList[j][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                            lvi.SubItems.Add(parsedDate.ToString("yyyy/mm/dd"));
                        }
                        else
                        {
                            lvi.SubItems.Add("");
                        }
                    }
                    
                    //lvi.SubItems.Add(MasterList[j][i]);
                }
                listFields.Items.Add(lvi);
            }

            filter = 0;
        }

        private void UpdateCustomerList()
        {
                var reader = new StreamReader(File.OpenRead(@"T:\\Databases\\c_list.dbs"));
            
               //var reader = new StreamReader(File.OpenRead("c_list.dbs"));
            
            foreach (List<string> i in ContactList)
            {
                i.Clear();
            }

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(',');
                
                    listNumber.Add(values[0]);
                    listAccount.Add(values[1]);
                    listCustName.Add(values[2]);
                    listContact.Add(values[3]);
                    listPhone.Add(values[4]);
                    listEmail.Add(values[5]);
            }
            reader.Close();
        }

        private void ClearFields()
        {
            listFields.Items.Clear();
        }

        private void UpdateCurrentStatus()
        {
            permanent = 0;
            var reader = new StreamReader(File.OpenRead(@"T:\Databases\current.dbs"));
            //var reader = new StreamReader(File.OpenRead("current.dbs"));

            foreach (List<string> i in MasterList)
            {
                i.Clear();
            }

            //bool first = true;
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(',');
                
                    listA.Add(values[0]);
                    listB.Add(values[1]);
                    listC.Add(values[2]);
                    if (values[2] == "PERMA")
                    {
                        permanent++;
                    }
                    listD.Add(values[3]);
                    listE.Add(values[4]);
                    listF.Add(values[5]);
                    listG.Add(values[6]);
                    listH.Add(values[7]);
                    listI.Add(values[8]);
                    listJ.Add(values[9]);
                    listK.Add(values[10]);
                    listL.Add(values[11]);
                    listM.Add(values[12]);
                    listN.Add(values[13]);
                    listO.Add(values[14]);
            }

            reader.Close();

            string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };
            List<DateTime> listDate2 = new List<DateTime>();
            DateTime parsedDate2;
            List<DateTime> listDate = new List<DateTime>();
            DateTime parsedDate;

            string[] dateStrings = listI.ToArray();
            int ready = 0;
            int sent = 0;
            int inhouse = 0;
            int overdue = 0;
            int repair = 0;

            for (int i = 0; i < listI.Count; i++)
            {
                DateTime.TryParseExact(listI[i], formatstring,null,DateTimeStyles.AllowWhiteSpaces|DateTimeStyles.AdjustToUniversal,out parsedDate2);

                if (listJ[i] == "X-SERIES" || listJ[i] == "PROPAQ")
                {
                    if (DateTime.Today.AddMonths(-12).CompareTo(parsedDate2) > 0)
                    {
                        listL[i] = "N";
                    }
                }
                else
                {
                    if (DateTime.Today.AddMonths(-6).CompareTo(parsedDate2) > 0)
                    {
                        listL[i] = "N";
                    }
                }
            }

            for (int i = 0; i < listC.Count(); i++)
            {
                
                if (listC[i] == "OUT") 
                {
                    sent++;
                    DateTime.TryParseExact(listG[i], formatstring,null,DateTimeStyles.AllowWhiteSpaces|DateTimeStyles.AdjustToUniversal,out parsedDate);
                    if (DateTime.Today.AddMonths(-1).CompareTo(parsedDate) > 0) 
                    {
                        overdue++;
                    }

                }
                if (listC[i] == "IN")
                {
                    inhouse++;
                    if (listL[i] == "Y") { ready++; }
                    if (listM[i] == "Y") { repair++; }
                }
            }


            btnPool.Text = listA.Count() + " total Loaners";
            btnReady.Text = ready + " ready to ship";
            btnInHouse.Text = inhouse + " in-house";
            btnOut.Text = sent + " out in field";
            btnSent.Text = sent - overdue+" out within 30 days";
            btnNotReady.Text = inhouse - ready -repair+ " not ready";
            btnRepair.Text = repair+" require repairs";
            btnOverdue.Text=overdue+" loaners overdue";


        }

        private void btnPool_Click(object sender, EventArgs e)
        {
            UpdateCurrentStatus();
            PopulateLists();
            SearchData();
        }

        private void ShowInHouse()
        {
            ClearFields();
            filter = 1;

            for (int i = 0; i < listC.Count(); i++)
            {
                if (listC[i] == "IN")
                {
                    ListViewItem lvi = new ListViewItem(MasterList[0][i]);

                    for (int j = 1; j < MasterList.Count; j++)
                    {
                        if (j < 6 || j > 8)
                        {
                            lvi.SubItems.Add(MasterList[j][i]);
                        }
                        else
                        {
                            if (MasterList[j][i] != "")
                            {
                                DateTime parsedDate;
                                string[] formatstring = { "d/m/yyyy", "dd/m/yyyy", "d/mm/yyyy", "dd/mm/yyyy" };

                                DateTime.TryParseExact(MasterList[j][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                                lvi.SubItems.Add(parsedDate.ToString("yyyy/mm/dd"));
                            }
                            else
                            {
                                lvi.SubItems.Add("");
                            }
                        }
                        //lvi.SubItems.Add(MasterList[j][i]);
                    }

                    listFields.Items.Add(lvi);

                }
            }
        }

        private void btnInHouse_Click(object sender, EventArgs e)
        {
            ShowInHouse();
            SearchData();
        }
        

        private void ShowSent()
        {

            ClearFields();
            filter = 4;
            DateTime parsedDate;
            string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };

            for (int i = 0; i < listG.Count(); i++)
            {
                DateTime.TryParseExact(listG[i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                if (listC[i] == "OUT" && DateTime.Today.AddMonths(-1).CompareTo(parsedDate) < 0)
                {
                    
                    ListViewItem lvi = new ListViewItem(MasterList[0][i]);

                    for (int j = 1; j < MasterList.Count; j++)
                    {
                        if (j < 6 || j > 8)
                        {
                            lvi.SubItems.Add(MasterList[j][i]);
                        }
                        else
                        {
                            if (MasterList[j][i] != "")
                            {
                                DateTime parsedDate2;
                                string[] formatstring2 = { "d/m/yyyy", "dd/m/yyyy", "d/mm/yyyy", "dd/mm/yyyy" };

                                DateTime.TryParseExact(MasterList[j][i], formatstring2, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate2);

                                lvi.SubItems.Add(parsedDate2.ToString("yyyy/mm/dd"));
                            }
                            else
                            {
                                lvi.SubItems.Add("");
                            }
                        }
                        //lvi.SubItems.Add(MasterList[j][i]);
                    }
                    listFields.Items.Add(lvi);
                }
            }
        }

        private void btnSent_Click(object sender, EventArgs e)
        {
            ShowSent();
            SearchData();
        }

        private void ShowReady()
        {
            ClearFields();
            filter = 2;

            for (int i = 0; i < listL.Count(); i++)
            {
                if (listL[i] == "Y" && listC[i] == "IN")
                {
                    ListViewItem lvi = new ListViewItem(MasterList[0][i]);

                    for (int j = 1; j < MasterList.Count; j++)
                    {
                        if (j < 6 || j > 8)
                        {
                            lvi.SubItems.Add(MasterList[j][i]);
                        }
                        else
                        {
                            if (MasterList[j][i] != "")
                            {
                                DateTime parsedDate;
                                string[] formatstring = { "d/m/yyyy", "dd/m/yyyy", "d/mm/yyyy", "dd/mm/yyyy" };

                                DateTime.TryParseExact(MasterList[j][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                                lvi.SubItems.Add(parsedDate.ToString("yyyy/mm/dd"));
                            }
                            else
                            {
                                lvi.SubItems.Add("");
                            }
                        }

                        //lvi.SubItems.Add(MasterList[j][i]);
                    }
                    listFields.Items.Add(lvi);
                }
            }
        }

        private void btnReady_Click(object sender, EventArgs e)
        {
            ShowReady();
            SearchData();
        }

        private void ShowNotReady()
        {
            ClearFields();
            filter = 3;

            for (int i = 0; i < listL.Count(); i++)
            {
                if (listL[i] == "N" && listC[i] == "IN")
                {
                    ListViewItem lvi = new ListViewItem(MasterList[0][i]);

                    for (int j = 1; j < MasterList.Count; j++)
                    {
                        if (j < 6 || j > 8)
                        {
                            lvi.SubItems.Add(MasterList[j][i]);
                        }
                        else
                        {
                            if (MasterList[j][i] != "")
                            {
                                DateTime parsedDate;
                                string[] formatstring = { "d/m/yyyy", "dd/m/yyyy", "d/mm/yyyy", "dd/mm/yyyy" };

                                DateTime.TryParseExact(MasterList[j][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                                lvi.SubItems.Add(parsedDate.ToString("yyyy/mm/dd"));
                            }
                            else
                            {
                                lvi.SubItems.Add("");
                            }
                        }

                        //lvi.SubItems.Add(MasterList[j][i]);
                    }
                    listFields.Items.Add(lvi);
                }
            }
        }

        private void btnNotReady_Click(object sender, EventArgs e)
        {
            ShowNotReady();
            SearchData();
        }

        private void SearchList()
        {
            
            RunFilter();
                        List<string> search = new List<string>();

                        List<string> list1 = new List<string>();
                        List<string> list2 = new List<string>();
                        List<string> list3 = new List<string>();
                        List<string> list4 = new List<string>();
                        List<string> list5 = new List<string>();
                        List<string> list6 = new List<string>();
                        List<string> list7 = new List<string>();
                        List<string> list8 = new List<string>();
                        List<string> list9 = new List<string>();
                        List<string> list10 = new List<string>();
                        List<string> list11 = new List<string>();
                        List<string> list12 = new List<string>();
                        List<string> list13 = new List<string>();
                        List<string> list14 = new List<string>();

                        List<string> alist1 = new List<string>();
                        List<string> alist2 = new List<string>();
                        List<string> alist3 = new List<string>();
                        List<string> alist4 = new List<string>();
                        List<string> alist5 = new List<string>();
                        List<string> alist6 = new List<string>();
                        List<string> alist7 = new List<string>();
                        List<string> alist8 = new List<string>();
                        List<string> alist9 = new List<string>();
                        List<string> alist10 = new List<string>();
                        List<string> alist11 = new List<string>();
                        List<string> alist12 = new List<string>();
                        List<string> alist13 = new List<string>();
                        List<string> alist14 = new List<string>();

                        List<List<string>> ResultsList = new List<List<string>>();
                        List<List<string>> ResultsList2 = new List<List<string>>();
                        List<List<string>> ShownList = new List<List<string>>();

                        ResultsList.Add(list1);
                        ResultsList.Add(list2);
                        ResultsList.Add(list3);
                        ResultsList.Add(list4);
                        ResultsList.Add(list5);
                        ResultsList.Add(list6);
                        ResultsList.Add(list7);
                        ResultsList.Add(list8);
                        ResultsList.Add(list9);
                        ResultsList.Add(list10);
                        ResultsList.Add(list11);
                        ResultsList.Add(list12);
                        ResultsList.Add(list13);
                        ResultsList.Add(list14);

                        ResultsList2.Add(alist1);
                        ResultsList2.Add(alist2);
                        ResultsList2.Add(alist3);
                        ResultsList2.Add(alist4);
                        ResultsList2.Add(alist5);
                        ResultsList2.Add(alist6);
                        ResultsList2.Add(alist7);
                        ResultsList2.Add(alist8);
                        ResultsList2.Add(alist9);
                        ResultsList2.Add(alist10);
                        ResultsList2.Add(alist11);
                        ResultsList2.Add(alist12);
                        ResultsList2.Add(alist13);
                        ResultsList2.Add(alist14);

                        
            List<string> temp = new List<string>();

            for (int i = 0; i < 14; i++)
            {
                foreach (ListViewItem lvi in listFields.Items)
                {
                    temp.Add(lvi.SubItems[i].Text);
                }
                
                ShownList.Add(temp.ToList());
                temp.Clear();
            }
            
                        /*
                        ShownList.Add(listSN.Items.Cast<String>().ToList());
                        ShownList.Add(listPN.Items.Cast<String>().ToList());
                        ShownList.Add(listStatus.Items.Cast<String>().ToList());
                        ShownList.Add(listCustomer.Items.Cast<String>().ToList());
                        ShownList.Add(listLoanerSR.Items.Cast<String>().ToList());
                        ShownList.Add(listRepairSR.Items.Cast<String>().ToList());
                        ShownList.Add(listDateSent.Items.Cast<String>().ToList());
                        ShownList.Add(listDateReturn.Items.Cast<String>().ToList());
                        ShownList.Add(listLastPM.Items.Cast<String>().ToList());
                        ShownList.Add(listModel.Items.Cast<String>().ToList());
                        ShownList.Add(listSW.Items.Cast<String>().ToList());
                        ShownList.Add(listReady.Items.Cast<String>().ToList());
                        ShownList.Add(listRepair.Items.Cast<String>().ToList());
                        ShownList.Add(listNotes.Items.Cast<String>().ToList());
                        */
                        ClearFields();
                        

                        search.AddRange(txtSearch.Text.ToUpper().Split(' ').ToArray());

                        foreach (string q in search)
                        {
                            //If it is the first search term
                            if (search.IndexOf(q) == 0)
                            {
                                for (int i = 0; i < ShownList.Count; i++)
                                {
                                    for (int j = 0; j < ShownList[i].Count; j++)
                                    {
                                        if (ShownList[i][j].Contains(q))
                                        {
                                            if (ResultsList[0].Contains(ShownList[0][j]) == false)
                                            {
                                                for (int k = 0; k < ShownList.Count(); k++)
                                                {

                                                    ResultsList[k].Add(ShownList[k][j]);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                //Add Results to Results2
                                for (int i = 0; i < ResultsList.Count(); i++)
                                {
                                    ResultsList2[i].AddRange(ResultsList[i]);
                                    ResultsList[i].Clear();
                                }

                                //Search Results2 for keywords, add them to Results
                                for (int i = 0; i < ResultsList2.Count; i++)
                                {
                                    for (int j = 0; j < ResultsList2[i].Count; j++)
                                    {
                                        if (ResultsList2[i][j].Contains(q))
                                        {
                                            if (ResultsList[0].Contains(ResultsList2[0][j]) == false)
                                            {
                                                for (int k = 0; k < ResultsList2.Count; k++)
                                                {
                                                    ResultsList[k].Add(ResultsList2[k][j]);
                                                }
                                            }
                                        }
                                    }

                                }
                                foreach(List<string> item in ResultsList2)
                                {
                                    item.Clear();
                                }
                            }
                        }

            listFields.Items.Clear();

            for (int i = 0; i < ResultsList[0].Count; i++)
            {
                ListViewItem lvi = new ListViewItem(ResultsList[0][i]);
                for (int j = 1; j < ResultsList.Count; j++)
                {
                    if (j < 6 || j > 8)
                    {
                        lvi.SubItems.Add(ResultsList[j][i]);
                    }
                    else
                    {
                        if (ResultsList[j][i] != "")
                        {
                            DateTime parsedDate;
                            string[] formatstring = { "yyyy/mm/dd","dd/mm/yyyy", "d/mm/yyyy", "d/m/yyyy", "dd/m/yyyy" };

                            DateTime.TryParseExact(ResultsList[j][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                            lvi.SubItems.Add(parsedDate.ToString("yyyy/mm/dd"));
                        }
                        else
                        {
                            lvi.SubItems.Add("");
                        }
                    }
                    //lvi.SubItems.Add(ResultsList[j][i]);
                }
                listFields.Items.Add(lvi);
            }

                lblResults.Text = listFields.Items.Count.ToString() + " results found.";
        }

        private void SearchData()
        {
            if (txtSearch.Text == "")
            {
                RunFilter();
                return;
            }

            if (txtSearch.Text != "" && txtSearch.Text.EndsWith(" ") == false)
            {
                SearchList();
            }
            else
            {
                lblResults.Text = "Search Error";
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {

        }
        
        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void RunFilter()
        {
            switch (filter)
            {
                case 0:
                    PopulateLists();
                    lblResults.Text = listFields.Items.Count + " results";
                    lblSearch.Text = "Search in 'All'";
                    break;
                case 1:
                    ShowInHouse();
                    lblResults.Text = listFields.Items.Count + " results";
                    lblSearch.Text = "Search in 'In-House'";
                    break;
                case 2:
                    ShowReady();
                    lblResults.Text = listFields.Items.Count + " results";
                    lblSearch.Text = "Search in 'Ready'";
                    break;
                case 3:
                    ShowNotReady();
                    lblResults.Text = listFields.Items.Count + " results";
                    lblSearch.Text = "Search in 'Not Ready'";
                    break;
                case 4:
                    ShowSent();
                    lblResults.Text = listFields.Items.Count + " results";
                    lblSearch.Text = "Search in 'Out within 30'";
                    break;
                case 5:
                    ShowOverdue();
                    lblResults.Text = listFields.Items.Count + " results";
                    lblSearch.Text = "Search in 'Overdue'";
                    break;
                case 6:
                    ShowRepair();
                    lblResults.Text = listFields.Items.Count + " results";
                    lblSearch.Text = "Search in 'Require repair'";
                    break;
                case 7:
                    ShowOut();
                    lblResults.Text = listFields.Items.Count + " results";
                    lblSearch.Text = "Search in 'Out'";
                    break;
                default:
                    MessageBox.Show("RunFilter method fucked up");
                    break;
            }
        }

        private void ClearNewEntryForm()
        {
            txtNewSN.Clear();
            txtNewItem.Clear();
            cmbNewStatus.SelectedIndex=-1;
            txtNewCustomer.Clear();
            txtNewLoanerSR.Clear();
            txtNewRepairSR.Clear();
            txtNewDateOut.Clear();
            txtNewDateReturn.Clear();
            txtNewDatePM.Clear();
            cmbNewModel.SelectedIndex=-1;
            txtNewSW.Clear();
            cmbNewReady.SelectedIndex=-1;
            cmbNewRepair.SelectedIndex=-1;
            txtNewNotes.Clear() ;
            txtNewCustNum.Clear();
        }

        private void CommaNewEntryForm()
        {
            txtNewSN.Text = txtNewSN.Text.Replace(",", " ");
            txtNewItem.Text = txtNewItem.Text.Replace(",", " ");
            txtNewCustomer.Text = txtNewCustomer.Text.Replace(",", " ");
            txtNewLoanerSR.Text = txtNewLoanerSR.Text.Replace(",", " ");
            txtNewRepairSR.Text = txtNewRepairSR.Text.Replace(",", " ");
            txtNewDateOut.Text = txtNewDateOut.Text.Replace(",", " ");
            txtNewDateReturn.Text = txtNewDateReturn.Text.Replace(",", " ");
            txtNewDatePM.Text = txtNewDatePM.Text.Replace(",", " ");
            txtNewSW.Text = txtNewSW.Text.Replace(",", " ");
            txtNewNotes.Text = txtNewNotes.Text.Replace(",", " ");
            txtNewCustNum.Text = txtNewCustNum.Text.Replace(",", " ");
        }

        private void btnNewSubmit_Click(object sender, EventArgs e)
        {
            CommaNewEntryForm();
            try
            {
                UpdateCurrentStatus();
            }
            catch
            {
                MessageBox.Show("Getting latest database list failed, please try again in a bit.");
                return;
            }

            if (txtNewSN.Text == "" || txtNewItem.Text == "" || cmbNewStatus.SelectedIndex == -1 || txtNewCustNum.Text==""||txtNewCustomer.Text == "" || txtNewLoanerSR.Text == "" || cmbNewModel.SelectedIndex == -1 || txtNewSW.Text == "" || cmbNewReady.SelectedIndex == -1 || cmbNewRepair.SelectedIndex == -1)
            {
                MessageBox.Show("Please fill all mandatory fields");
                return;
            }

            //Mode 1: New
            if (radNewEntry.Checked==true)
            {
                try
                {
                    var NewEntry = new StreamWriter(@"T:\Databases\current.dbs", true);

                    string Entry = "";
                    if (txtNewRepairSR.Text == "") { txtNewRepairSR.Text = "000000"; }
                    if (txtNewDateOut.Text == "DD/MM/YYYY") { txtNewDateOut.Text = "01/01/2001"; }
                    if (txtNewDateReturn.Text == "DD/MM/YYYY") { txtNewDateReturn.Text = DateTime.Now.ToString("dd/MM/yyyy"); ; }
                    if (txtNewDatePM.Text == "DD/MM/YYYY") { txtNewDatePM.Text = "01/01/2001"; }
                    if (txtNewNotes.Text == "") { txtNewLoanerSR.Text = "N/A"; }

                    Entry = txtNewSN.Text + "," + txtNewItem.Text + "," + cmbNewStatus.Text + "," + txtNewCustomer.Text + "," + txtNewLoanerSR.Text + "," + txtNewRepairSR.Text + "," + txtNewDateOut.Text + "," + txtNewDateReturn.Text + "," + txtNewDatePM.Text + "," + cmbNewModel.Text + "," + txtNewSW.Text + "," + cmbNewReady.Text + "," + cmbNewRepair.Text + "," + txtNewNotes.Text + "," + txtNewCustNum.Text;
                    
                    try
                    {
                        UpdateRecords("CREATE", txtNewSN.Text,txtNewCustNum.Text,txtNewCustomer.Text, txtNewLoanerSR.Text, txtNewRepairSR.Text, txtNewNotes.Text,"",cmbNewModel.Text);
                    }
                    catch
                    {
                        MessageBox.Show("There was an error during recording. Records were not updated.");
                    }

                    NewEntry.WriteLine(Entry);

                    NewEntry.Close();

                    ClearNewEntryForm();

                    lblSuccess.Text = "New Entry successful";
                    Util.Animate(lblSuccess, Util.Effect.Slide, 90, 90);
                    timer1.Enabled = true;
                    timer1.Start();

                    
                    UpdateCurrentStatus();
                    txtSearch.Text = txtSearch.Text;
                }
                catch
                {
                    MessageBox.Show("Loaners list is currently open. Try again in a bit.");
                    return;
                }
            }

            //Mode 2: Edit
            if (radEditEntry.Checked==true)
            {
                try {
                    if (txtNewRepairSR.Text == "") { txtNewRepairSR.Text = "000000"; }
                    if (txtNewDateOut.Text == "DD/MM/YYYY") { txtNewDateOut.Text = "01/01/2001"; }
                    if (txtNewDateReturn.Text == "DD/MM/YYYY") { txtNewDateReturn.Text = DateTime.Now.ToString("dd/MM/yyyy"); ; }
                    if (txtNewDatePM.Text == "DD/MM/YYYY") { txtNewDatePM.Text = "01/01/2001"; }
                    if (txtNewNotes.Text == "") { txtNewLoanerSR.Text = "N/A"; }

                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        if (txtNewSN.Text == MasterList[0][i])
                        {
                            MasterList[1][i] = txtNewItem.Text;
                            MasterList[2][i] = cmbNewStatus.Text;
                            MasterList[3][i] = txtNewCustomer.Text;
                            MasterList[4][i] = txtNewLoanerSR.Text;
                            MasterList[5][i] = txtNewRepairSR.Text;
                            MasterList[6][i] = txtNewDateOut.Text;
                            MasterList[7][i] = txtNewDateReturn.Text;
                            MasterList[8][i] = txtNewDatePM.Text;
                            MasterList[9][i] = cmbNewModel.Text;
                            MasterList[10][i] = txtNewSW.Text;
                            MasterList[11][i] = cmbNewReady.Text;
                            MasterList[12][i] = cmbNewRepair.Text;
                            MasterList[13][i] = txtNewNotes.Text;
                            MasterList[14][i] = txtNewCustNum.Text;
                            break;
                        }
                    }

                    try
                    {
                        UpdateRecords("EDIT", txtNewSN.Text, txtNewCustNum.Text, txtNewCustomer.Text, txtNewLoanerSR.Text, txtNewRepairSR.Text, txtNewNotes.Text, "",cmbNewModel.Text);
                    }
                    catch
                    {
                        MessageBox.Show("There was an error during recording. Records were not updated.");
                    }

                    var database = new StreamWriter(@"T:\Databases\current.dbs");

                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        string entry = "";

                        entry = MasterList[0][i] + "," + MasterList[1][i] + "," + MasterList[2][i] + "," + MasterList[3][i] + "," + MasterList[4][i] + "," + MasterList[5][i] + "," + MasterList[6][i] + "," + MasterList[7][i] + "," + MasterList[8][i] + "," + MasterList[9][i] + "," + MasterList[10][i] + "," + MasterList[11][i] + "," + MasterList[12][i] + "," + MasterList[13][i] + "," + MasterList[14][i];
                        database.WriteLine(entry);
                    }
                    database.Close();

                    ClearNewEntryForm();
                    UpdateCurrentStatus();
                    //MessageBox.Show("Edit Successful");
                    lblSuccess.Text = "Edit Entry successful";
                    Util.Animate(lblSuccess, Util.Effect.Slide, 90, 90);
                    timer1.Enabled = true;
                    timer1.Start();
                }
                catch
                {
                    MessageBox.Show("Loaners list is currently open. Try again in a bit.");
                    return;
                }
            }

            //Mode 3: Delete
            if (radDelete.Checked==true)
            {
                try {
                    var database = new StreamWriter(@"T:\Databases\current.dbs");

                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        string entry = "";
                        if (MasterList[0][i] != txtNewSN.Text)
                        {
                            entry = MasterList[0][i] + "," + MasterList[1][i] + "," + MasterList[2][i] + "," + MasterList[3][i] + "," + MasterList[4][i] + "," + MasterList[5][i] + "," + MasterList[6][i] + "," + MasterList[7][i] + "," + MasterList[8][i] + "," + MasterList[9][i] + "," + MasterList[10][i] + "," + MasterList[11][i] + "," + MasterList[12][i] + "," + MasterList[13][i] +"," + MasterList[14][i];
                            database.WriteLine(entry);
                        }
                    }

                    try
                    {
                        UpdateRecords("DELETE", txtNewSN.Text, txtNewCustNum.Text, txtNewCustomer.Text, txtNewLoanerSR.Text, txtNewRepairSR.Text, txtNewNotes.Text, "",cmbNewModel.Text);
                    }
                    catch
                    {
                        MessageBox.Show("There was an error during recording. Records were not updated.");
                    }

                    database.Close();
                    ClearNewEntryForm();
                    UpdateCurrentStatus();
                    //MessageBox.Show("Delete Successful");
                    lblSuccess.Text = "Delete Entry successful";
                    Util.Animate(lblSuccess, Util.Effect.Slide, 90, 90);
                    timer1.Enabled = true;
                    timer1.Start();
                }
                catch
                {
                    MessageBox.Show("Loaners list is currently open. Try again in a bit.");
                    return;
                }
            }
        }

        private void btnNewClear_Click(object sender, EventArgs e)
        {
            ClearNewEntryForm();
        }

        private void txtNewSN_TextChanged(object sender, EventArgs e)
        {
            if (txtNewSN.Text.Length >= 5)
            {
                if (radEditEntry.Checked==true||radDelete.Checked==true)
                {
                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        if (txtNewSN.Text == MasterList[0][i])
                        {
                            txtNewItem.Text = MasterList[1][i];
                            cmbNewStatus.Text = MasterList[2][i];
                            txtNewCustomer.Text = MasterList[3][i];
                            txtNewLoanerSR.Text = MasterList[4][i];
                            txtNewRepairSR.Text = MasterList[5][i];
                            txtNewDateOut.Text = MasterList[6][i];
                            txtNewDateReturn.Text = MasterList[7][i];
                            txtNewDatePM.Text = MasterList[8][i];
                            cmbNewModel.Text = MasterList[9][i];
                            txtNewSW.Text = MasterList[10][i];
                            cmbNewReady.Text = MasterList[11][i];
                            cmbNewRepair.Text = MasterList[12][i];
                            txtNewNotes.Text = MasterList[13][i];
                            txtNewCustNum.Text = MasterList[14][i];
                            btnNewSubmit.Enabled = true;
                            return;
                        }
                    }
                    txtNewItem.Clear();
                    cmbNewStatus.SelectedIndex = -1;
                    txtNewCustomer.Clear();
                    txtNewLoanerSR.Clear();
                    txtNewRepairSR.Clear();
                    txtNewDateOut.Clear();
                    txtNewDateReturn.Clear();
                    txtNewDatePM.Clear();
                    cmbNewModel.SelectedIndex = -1;
                    txtNewSW.Clear();
                    cmbNewReady.SelectedIndex = -1;
                    cmbNewRepair.SelectedIndex = -1;
                    txtNewNotes.Clear();
                    txtNewCustNum.Clear();
                    btnNewSubmit.Enabled = false;
                }
                else
                {
                    btnNewSubmit.Enabled = true;
                }
            }
            else
            {
                txtNewItem.Clear();
                cmbNewStatus.SelectedIndex = -1;
                txtNewCustomer.Clear();
                txtNewLoanerSR.Clear();
                txtNewRepairSR.Clear();
                txtNewDateOut.Clear();
                txtNewDateReturn.Clear();
                txtNewDatePM.Clear();
                cmbNewModel.SelectedIndex = -1;
                txtNewSW.Clear();
                cmbNewReady.SelectedIndex = -1;
                cmbNewRepair.SelectedIndex = -1;
                txtNewNotes.Clear();
                txtNewCustNum.Clear();
                btnNewSubmit.Enabled = false;
            }
            
        }

        private void txtNewSN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                CheckSNBox();
            }
            
        }
        
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (tabControl1.SelectedIndex == 0)
            {
                PopulateLists();    
                if (txtSearch.Text != "")
                {
                    SearchList();
                }
            }

            if (tabControl1.SelectedIndex == 1)
            {
                txtInfoSN.Focus();
                string info = txtInfoSN.Text;
                txtInfoSN.Clear();
                txtInfoSN.Text = info;
            }

            if (tabControl1.SelectedIndex == 3)
            {
                if (login == 0) txtNewLogin.Focus();
                if (logbuttpushed == false)
                {
                    currentindex = tabControl1.SelectedIndex;
                }
            }

            if (tabControl1.SelectedIndex == 2||tabControl1.SelectedIndex == 4)
            {
                UpdateCustomerList();
                currentindex = tabControl1.SelectedIndex;
                
            }

            if (tabControl1.SelectedIndex == 6)
            {
                RemoveOldRecords();
                calStart.MinDate = DateTime.Today.AddMonths(-24).AddDays(-1);
                grpMovement.Left = 254;
            }

            if (tabControl1.SelectedIndex == 7)
            {
                LoadUsers();
            }

            logbuttpushed = false;
        }

        private void CheckSNBox()
        {
            if (txtNewSN.Text == "")
            {
                return;
            }

            if (radNewEntry.Checked==true)
            {
                for (int i = 0; i < MasterList[0].Count; i++)
                {
                    if (txtNewSN.Text == MasterList[0][i])
                    {
                        MessageBox.Show("Entry already exists! Switching to Edit Mode");
                        //Mode 2 is Edit Entry
                        radEditEntry.Checked = true;

                        grpForms.Enabled = true;
                        lblEditHelp.Visible = true;
                        btnNewSubmit.Text = "Confirm Edit";
                    }
                }
            }
           
                if (radEditEntry.Checked==true || radDelete.Checked==true)
                {
                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        if (txtNewSN.Text == MasterList[0][i])
                        {
                            txtNewItem.Text = MasterList[1][i];
                            cmbNewStatus.Text = MasterList[2][i];
                            txtNewCustomer.Text = MasterList[3][i];
                            txtNewLoanerSR.Text = MasterList[4][i];
                            txtNewRepairSR.Text = MasterList[5][i];
                            txtNewDateOut.Text = MasterList[6][i];
                            txtNewDateReturn.Text = MasterList[7][i];
                            txtNewDatePM.Text = MasterList[8][i];
                            cmbNewModel.Text = MasterList[9][i];
                            txtNewSW.Text = MasterList[10][i];
                            cmbNewReady.Text = MasterList[11][i];
                            cmbNewRepair.Text = MasterList[12][i];
                            txtNewNotes.Text = MasterList[13][i];
                            txtNewCustNum.Text = MasterList[14][i];
                            break;
                        }
                    }
                }
            
        }

        private void txtNewSN_Leave(object sender, EventArgs e)
        {
            CheckSNBox();
        }
        
        private void txtNewDateOut_Enter(object sender, EventArgs e)
        {
            txtNewDateOut.Text = "";
        }

        private void txtNewDateReturn_Enter(object sender, EventArgs e)
        {
            txtNewDateReturn.Text = "";
        }

        private void txtNewDatePM_Enter(object sender, EventArgs e)
        {
            txtNewDatePM.Text = "";
        }

        private void ShowOverdue()
        {
            ClearFields();
            filter = 5;
            DateTime parsedDate;
            string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };

            for (int i = 0; i < listG.Count(); i++)
            {
                DateTime.TryParseExact(listG[i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                if (listC[i] == "OUT" && DateTime.Today.AddMonths(-1).CompareTo(parsedDate) > 0)
                {
                    
                    ListViewItem lvi = new ListViewItem(MasterList[0][i]);

                    for (int j = 1; j < MasterList.Count; j++)
                    {
                        if (j < 6 || j > 8)
                        {
                            lvi.SubItems.Add(MasterList[j][i]);
                        }
                        else
                        {
                            if (MasterList[j][i] != "")
                            {
                                DateTime parsedDate2;
                                string[] formatstring2 = { "d/m/yyyy", "dd/m/yyyy", "d/mm/yyyy", "dd/mm/yyyy" };

                                DateTime.TryParseExact(MasterList[j][i], formatstring2, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate2);

                                lvi.SubItems.Add(parsedDate2.ToString("yyyy/mm/dd"));
                            }
                            else
                            {
                                lvi.SubItems.Add("");
                            }
                        }
                        //lvi.SubItems.Add(MasterList[j][i]);
                    }
                    listFields.Items.Add(lvi);
                }
            }
        }

        private void btnOverdue_Click(object sender, EventArgs e)
        {
            ShowOverdue();
            SearchData();
        }

        /*private void radPM_CheckedChanged(object sender, EventArgs e)
        {
            if (radPM.Checked == true)
            {
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);
                lblOpMode.Text = "Finish PM";
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);

                grpOperations.Enabled = true;

                if (grpOpSR.Visible == true)
                    Util.Animate(grpOpSR, Util.Effect.Slide, 100, 90);
                grpOpSR.Visible = false;

                if (grpOpCustomer.Visible == true)
                Util.Animate(grpOpCustomer, Util.Effect.Slide, 100, 90);
                grpOpCustomer.Visible = false;
                
                

                chkOpRepair.Visible = true;
                chkOpRepair.Text = "Failed PM and Requires repair";
                chkOpRepair.Checked = false;
                txtOpSN.Focus();
            }
        }*/

       /* private void radSend_CheckedChanged(object sender, EventArgs e)
        {
            if (radSend.Checked == true)
            {
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);
                lblOpMode.Text = "Sending Loaner";
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);
                
                grpOperations.Enabled = true;

                if (grpOpCustomer.Visible == false)
                Util.Animate(grpOpCustomer, Util.Effect.Slide, 100, 90);
                grpOpCustomer.Visible = true;

                if (grpOpSR.Visible == false)
                Util.Animate(grpOpSR, Util.Effect.Slide, 100, 90);
                grpOpSR.Visible = true;

                chkOpRepair.Visible = false;
                chkOpRepair.Checked = false;
                txtOpSN.Focus();
            }
        }*/

        /*private void radReceiving_CheckedChanged(object sender, EventArgs e)
        {
            if (radReceiving.Checked == true)
            {
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);
                lblOpMode.Text = "Receiving Loaner";
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);
                grpOperations.Enabled = true;

                if(grpOpCustomer.Visible==false)
                Util.Animate(grpOpCustomer, Util.Effect.Slide, 100, 90);
                grpOpCustomer.Visible = true;
                
                if (grpOpSR.Visible == false)
                Util.Animate(grpOpSR, Util.Effect.Slide, 100, 90);
                grpOpSR.Visible = true;
                chkOpRepair.Visible = true;
                chkOpRepair.Checked = false;
                chkOpRepair.Text = "Repair required";
                txtOpSN.Focus();
            }
        }*/

        private void RemoveOldRecords()
        {
            string record = "record.dbs";
            List<string> RecA = new List<string>();
            List<string> RecB = new List<string>();
            List<string> RecC = new List<string>();
            List<string> RecD = new List<string>();
            List<string> RecE = new List<string>();
            List<string> RecF = new List<string>();
            List<string> RecG = new List<string>();
            List<string> RecH = new List<string>();
            List<string> RecI = new List<string>();
            List<string> RecJ = new List<string>();
            List<string> RecK = new List<string>();
            //List<List<string>> MasterRecord = new List<List<string>>();

            DateTime parsedDate;
            string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };
            var reader = new StreamReader(record);
            string dummy;
            bool first = true;

            while (!reader.EndOfStream)
            {
                if (first == true)
                {
                    dummy = reader.ReadLine();
                    first = false;
                }
                else 
                { 
                var line = reader.ReadLine();
                var values = line.Split(',');
                DateTime.TryParseExact(values[0], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);
                
                    if (DateTime.Today.AddMonths(-24).CompareTo(parsedDate) < 0)
                    {
                        RecA.Add(values[0]);
                        RecB.Add(values[1]);
                        RecC.Add(values[2]);
                        RecD.Add(values[3]);
                        RecE.Add(values[4]);
                        RecF.Add(values[5]);
                        RecG.Add(values[6]);
                        RecH.Add(values[7]);
                        RecI.Add(values[8]);
                        RecJ.Add(values[9]);
                        RecK.Add(values[10]);
                    }
                }
            }
            reader.Close();

            var fullrecord = new StreamWriter(record,false);
            fullrecord.WriteLine("Date,Operator,Serial,Customer Number,Customer,Loaner SR, Complaint SR,Notes,Duration Since Last Movement,User,Model");

            for (int i = 0; i < RecA.Count; i++)
            {
                fullrecord.WriteLine(RecA[i] + "," + RecB[i] + "," + RecC[i] + "," + RecD[i] + "," + RecE[i] + "," + RecF[i] + "," + RecG[i] + "," + RecH[i]+","+RecI[i]+","+RecJ[i]+","+RecK[i]);
            }

            fullrecord.Close();
        }

        private void UpdateRecords(string Operation,string SN, string custnum,string customer, string loanersr, string repairsr,string notes,string duration, string model)
        {
            string monthlymovement = "./Monthly Reports/"+DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MMMM") + ".csv";
            string record = "record.dbs";
            
            if (File.Exists(monthlymovement) == false)
            {
                var init = new StreamWriter(monthlymovement, true);
                init.WriteLine("Date,Operation,Serial,Customer Number,Customer,Loaner SR, Complaint SR,Notes,Duration Since Last Movement,User,Model");
                init.Close();
            }

            var monthrecord = new StreamWriter(monthlymovement, true);
            monthrecord.WriteLine(DateTime.Now.ToString("dd/MM/yyyy") + "," + Operation + "," + SN + ","+custnum+"," + customer + "," + loanersr + "," + repairsr + "," + notes+","+duration+","+currentuser+","+model);
            monthrecord.Close();
            
            var fullrecord = new StreamWriter(record,true);

            fullrecord.WriteLine(DateTime.Now.ToString("dd/MM/yyyy") + "," + Operation + "," + SN + "," + custnum + "," + customer + "," + loanersr + "," + repairsr + "," + notes + "," + duration + "," + currentuser + "," + model);
            fullrecord.Close();
        }

        private void btnOpConfirm_Click(object sender, EventArgs e)
        {
            CommaOpFields();

            try {
                UpdateCurrentStatus();
            }
            catch
            {
                MessageBox.Show("Getting latest database list failed, please try again in a bit.");
                return;
            }

            bool found = false;
            //1. PM Loaner
            if (radPM.Checked==true)
            {
                try
                {
                    if (txtOpNotes.Text=="") { txtOpNotes.Text = "N/A"; }
                    //if (txtNewDateReturn.Text == "DD/MM/YYYY") { txtNewDateReturn.Text = DateTime.Now.ToString("dd/MM/yyyy"); ; }
                    
                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        if (txtOpSN.Text == MasterList[0][i])
                        {
                            MasterList[8][i] = DateTime.Now.ToString("dd/MM/yyyy");

                            if (txtOpSW.Text != "")
                            {
                                MasterList[10][i] = txtOpSW.Text;
                            }

                            if (chkOpRepair.Checked == false)
                            {
                                MasterList[11][i] = "Y";
                                MasterList[12][i] = "N";
                            }
                            else
                            {
                                MasterList[11][i] = "N";
                                MasterList[12][i] = "Y";
                            }

                            MasterList[13][i] = txtOpNotes.Text;

                           try
                            {
                                UpdateRecords("PM", txtOpSN.Text, MasterList[14][i], txtOpCustomer.Text, MasterList[4][i], MasterList[5][i], txtOpNotes.Text, "", MasterList[9][i]);
                            }
                            catch
                            {
                                MessageBox.Show("There was an error during recording. Records were not updated.");
                            }
                            break;
                        }
                    }

                    var database = new StreamWriter(@"T:\Databases\current.dbs");

                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        string entry = "";

                        entry = MasterList[0][i] + "," + MasterList[1][i] + "," + MasterList[2][i] + "," + MasterList[3][i] + "," + MasterList[4][i] + "," + MasterList[5][i] + "," + MasterList[6][i] + "," + MasterList[7][i] + "," + MasterList[8][i] + "," + MasterList[9][i] + "," + MasterList[10][i] + "," + MasterList[11][i] + "," + MasterList[12][i] + "," + MasterList[13][i]+","+MasterList[14][i];
                        database.WriteLine(entry);
                    }
                    database.Close();
                    

                    UpdateCurrentStatus();
                    PopulateLists();
                    //MessageBox.Show("PM status change successful");
                    ClearOpFields();

                    lblSuccess3.Text = "PM successful";
                    Util.Animate(lblSuccess3, Util.Effect.Slide, 90, 90);

                    txtCopy.Text = "The unit passed the Final Test (PM) and was recertified for clinical use.";

                    timer1.Enabled = true;
                    timer1.Start();

                    
                }
                catch
                {
                   MessageBox.Show("Loaners list is currently open. Try again in a bit.");
                   return;
                }
            }

            //2. Send Loaner
            if (radSend.Checked == true)
            {
                try
                {
                    if (txtOpNotes.Text == "") { txtOpNotes.Text = "N/A"; }
                    //if (txtNewDateReturn.Text == "DD/MM/YYYY") { txtNewDateReturn.Text = DateTime.Now.ToString("dd/MM/yyyy"); ; }

                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        if (txtOpSN.Text == MasterList[0][i])
                        {
                            MasterList[2][i] = "OUT";
                            MasterList[3][i] = txtOpCustomer.Text;
                            MasterList[4][i] = txtOpLoaner.Text;
                            MasterList[5][i] = txtOpRepair.Text;
                            MasterList[6][i] = DateTime.Now.ToString("dd/MM/yyyy");

                            string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };
                            DateTime parsedDate;
                            DateTime.TryParseExact(MasterList[7][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);
                            string duration = (DateTime.Today - parsedDate).TotalDays.ToString();

                            MasterList[7][i] = "";
                            if (txtOpSW.Text != "")
                            {
                                MasterList[10][i] = txtOpSW.Text;
                            }
                            MasterList[11][i] = "N";
                            MasterList[13][i] = txtOpNotes.Text;
                            MasterList[14][i] = txtOpAccnt.Text;
                            
                            for (int j = 0; j < ContactList[0].Count; j++)
                            {
                                if (ContactList[0][j] == txtOpAccnt.Text)
                                {
                                    found = true;
                                    break;
                                }
                            }
                           
                                try
                                {
                                    UpdateRecords("SEND", txtOpSN.Text,txtOpAccnt.Text,txtOpCustomer.Text, MasterList[4][i], MasterList[5][i], txtOpNotes.Text,duration, MasterList[9][i]);
                                }
                                catch
                                {
                                    MessageBox.Show("There was an error during recording. Records were not updated.");
                                }

                            break;
                        }
                    }

                    var database = new StreamWriter(@"T:\Databases\current.dbs");

                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        string entry = "";

                        entry = MasterList[0][i] + "," + MasterList[1][i] + "," + MasterList[2][i] + "," + MasterList[3][i] + "," + MasterList[4][i] + "," + MasterList[5][i] + "," + MasterList[6][i] + "," + MasterList[7][i] + "," + MasterList[8][i] + "," + MasterList[9][i] + "," + MasterList[10][i] + "," + MasterList[11][i] + "," + MasterList[12][i] + "," + MasterList[13][i] + "," + MasterList[14][i];
                        database.WriteLine(entry);
                    }
                    database.Close();

                    DocX letter = DocX.Load(@"T:\\! LOANER SR FOLDERS\\" + txtOpLoaner.Text + "\\" + txtOpLoaner.Text + "_request.docx");
                    letter.ReplaceText("#loanersn#", txtOpSN.Text);

                    string[] acc_sn = new string[10];

                    foreach(DataGridViewRow dgv in dgvAccessories.Rows)
                    {
                        switch (dgv.Cells[0].Value.ToString())
                        {
                            case "SpO2 Cable":
                                letter.ReplaceText("#spo2cable#", "LOT# " + dgv.Cells[1].Value.ToString());
                                acc_sn[0] = dgv.Cells[1].Value.ToString();
                                break;

                            case "SpO2 Sensor":
                                letter.ReplaceText("#spo2sensor#", "LOT# " + dgv.Cells[1].Value.ToString());
                                acc_sn[1] = dgv.Cells[1].Value.ToString();
                                break;

                            case "EtCO2 Sensor":
                                letter.ReplaceText("#etco2sensor#", "LOT# " + dgv.Cells[1].Value.ToString());
                                acc_sn[2] = dgv.Cells[1].Value.ToString();
                                break;

                            case "Battery 1":
                                letter.ReplaceText("#battery1#", "Serial# " + dgv.Cells[1].Value.ToString());
                                acc_sn[3] = dgv.Cells[1].Value.ToString();
                                break;

                            case "Battery 2":
                                letter.ReplaceText("#battery2#", "Serial# " + dgv.Cells[1].Value.ToString());
                                acc_sn[4] = dgv.Cells[1].Value.ToString();
                                break;

                            case "ROC Adaptor":
                                letter.ReplaceText("#roc#", "Colour: " + dgv.Cells[1].Value.ToString());
                                acc_sn[5] = dgv.Cells[1].Value.ToString();
                                break;

                            case "Pads 1":
                                letter.ReplaceText("#pads1#", "LOT# " + dgv.Cells[1].Value.ToString());
                                acc_sn[6] = dgv.Cells[1].Value.ToString();
                                break;

                            case "Pads 2":
                                letter.ReplaceText("#pads2#", "LOT# " + dgv.Cells[1].Value.ToString());
                                acc_sn[7] = dgv.Cells[1].Value.ToString();
                                break;

                            case "Paddles":
                                letter.ReplaceText("#paddles#", "LOT# " + dgv.Cells[1].Value.ToString());
                                acc_sn[8] = dgv.Cells[1].Value.ToString();
                                break;

                            case "Data Card":
                                letter.ReplaceText("#datacard#", "Size: " + dgv.Cells[1].Value.ToString());
                                acc_sn[9] = dgv.Cells[1].Value.ToString();
                                break;

                            default:
                                //do nothing
                                break;
                        }
                    }

                    letter.Save();
                    letter.Dispose();

                    var pending_reader = new StreamReader(@"T:\\Databases\\PendingLoaners.dbs");
                    List<string> pendinglist = new List<string>();

                    while (!pending_reader.EndOfStream)
                    {
                        string line = pending_reader.ReadLine();
                        pendinglist.Add(line);
                    }
                    pending_reader.Close();

                    if (pendinglist.Contains(txtOpLoaner.Text))
                    {
                        pendinglist.Remove(txtOpLoaner.Text);
                    }

                    var pending_writer = new StreamWriter(@"T:\\Databases\\PendingLoaners.dbs");
                    foreach(string i in pendinglist)
                    {
                        pending_writer.WriteLine(i);
                    }
                    pending_writer.Close();

                    var preader = new StreamReader(@"T:\\Databases\\PendingLoaners.dbs");
                    lstPending.Items.Clear();

                    while (!preader.EndOfStream)
                    {
                        string line = preader.ReadLine();
                        lstPending.Items.Add(line);
                    }
                    preader.Close();


                    var reader = new StreamReader(@"T:\\! LOANER SR FOLDERS\\" + txtOpLoaner.Text + "\\" + txtOpLoaner.Text + "_OutgoingInventory.txt");
                    string line1 = reader.ReadLine();
                    line1 = txtOpSN.Text + line1;
                    string line2 = reader.ReadLine();
                    var dataval = line2.Split('<');
                    string outputline2 = dataval[0] + "<" + acc_sn[0] + "<" + acc_sn[1] + "<" + acc_sn[2] + "<" + acc_sn[3] + "<" + acc_sn[4] + "<" + acc_sn[5] + "<" + acc_sn[6] + "<" + acc_sn[7] + "<" + acc_sn[8] + "<" + acc_sn[9] + "<<<<";
                    reader.Close();

                    var writer = new StreamWriter(@"T:\\! LOANER SR FOLDERS\\" + txtOpLoaner.Text + "\\" + txtOpLoaner.Text + "_OutgoingInventory.txt");
                    writer.WriteLine(line1);
                    writer.WriteLine(outputline2);
                    writer.Close();

                    Process.Start(@"T:\\! LOANER SR FOLDERS\\" + txtOpLoaner.Text + "\\" + txtOpLoaner.Text + "_request.docx");

                    UpdateCurrentStatus();
                    PopulateLists();
                    //MessageBox.Show("Loaner sign out successful");
                    ClearOpFields();

                    lblSuccess3.Text = "Sending successful";
                    Util.Animate(lblSuccess3, Util.Effect.Slide, 90, 90);

                    txtCopy.Text = "";

                    timer1.Enabled = true;
                    timer1.Start();

                    if (found == false)
                    {
                        NewCustomer(txtOpAccnt.Text + ",UNK," + txtOpCustomer.Text + ",Contact Name,xxx-xxxx,example@zoll.com");
                        MessageBox.Show("Customer Entry created for Customer Number: " + txtOpAccnt.Text + ".\nPlease review in Customer Management.");
                        //tabControl1.SelectedIndex = 3;
                        //radCustEdit.Checked = true;
                        //txtCustNum.Text = txtOpAccnt.Text;
                    }

                    
                }
                catch
                {
                    MessageBox.Show("Loaners list is currently open. Try again in a bit.");
                    return;
                }
            }

            //3. Receiving Loaner
            if (radReceiving.Checked == true)
            {
                try
                {
                    if (txtOpNotes.Text == "") { txtOpNotes.Text = "N/A"; }
                    //if (txtNewDateReturn.Text == "DD/MM/YYYY") { txtNewDateReturn.Text = DateTime.Now.ToString("dd/MM/yyyy"); ; }

                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        if (txtOpSN.Text == MasterList[0][i])
                        {
                            MasterList[2][i] = "IN";
                            MasterList[3][i] = txtOpCustomer.Text;
                            MasterList[4][i] = txtOpLoaner.Text;
                            MasterList[5][i] = txtOpRepair.Text;

                            string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };
                            DateTime parsedDate;
                            DateTime.TryParseExact(MasterList[6][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);
                           
                            string duration = (DateTime.Today - parsedDate).TotalDays.ToString();

                            MasterList[7][i] = DateTime.Now.ToString("dd/MM/yyyy");
                            if (txtOpSW.Text != "")
                            {
                                MasterList[10][i] = txtOpSW.Text;
                            }

                            MasterList[11][i] = "N";

                            if (chkOpRepair.Checked == true)
                            {
                                MasterList[12][i] = "Y";
                                txtCopy.Text = "The unit was received back from the customer the complaint of ____.\n\nThe unit is scheduled for repair and recertification.";
                            }
                            else
                            {
                                txtCopy.Text = "The unit was received back from the customer with no complaints.\n\nThe unit is scheduled for recertification.";
                            }

                            MasterList[13][i] = txtOpNotes.Text;
                            MasterList[14][i] = txtOpAccnt.Text;

                            for (int j = 0; j < ContactList[0].Count; j++)
                            {
                                if (ContactList[0][j] == txtOpAccnt.Text)
                                {
                                    found = true;
                                    break;
                                }
                            }

                            try
                            {
                                UpdateRecords("RECEIVE", txtOpSN.Text,txtOpAccnt.Text, txtOpCustomer.Text, MasterList[4][i], MasterList[5][i], txtOpNotes.Text,duration, MasterList[9][i]);
                            }
                            catch
                            {
                                MessageBox.Show("There was an error during recording. Records were not updated.");
                            }
                            break;
                        }
                    }

                    var database = new StreamWriter(@"T:\Databases\current.dbs");

                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        string entry = "";

                        entry = MasterList[0][i] + "," + MasterList[1][i] + "," + MasterList[2][i] + "," + MasterList[3][i] + "," + MasterList[4][i] + "," + MasterList[5][i] + "," + MasterList[6][i] + "," + MasterList[7][i] + "," + MasterList[8][i] + "," + MasterList[9][i] + "," + MasterList[10][i] + "," + MasterList[11][i] + "," + MasterList[12][i] + "," + MasterList[13][i]+","+MasterList[14][i];
                        database.WriteLine(entry);
                    }
                    database.Close();

                    UpdateCurrentStatus();
                    PopulateLists();
                    //MessageBox.Show("Loaner sign in successful");
                    ClearOpFields();

                    lblSuccess3.Text = "Receiving successful";
                    Util.Animate(lblSuccess3, Util.Effect.Slide, 90, 90);


                    timer1.Enabled = true;
                    timer1.Start();

                    if (found == false)
                    {
                        NewCustomer(txtOpAccnt.Text + ",UNK," + txtOpCustomer.Text + ",Contact Name,xxx-xxxx,example@zoll.com");
                        MessageBox.Show("Customer Entry created for Customer Number: " + txtOpAccnt.Text + ".\nPlease review in Customer Management.");
                        //tabControl1.SelectedIndex = 3;
                        //radCustEdit.Checked = true;
                        //txtCustNum.Text = txtOpAccnt.Text;
                    }

                    
                }
                catch
                {
                    MessageBox.Show("Loaners list is currently open. Try again in a bit.");
                    return;
                }
            }
        }

        private void txtOpSN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
        }

        private void txtOpSN_TextChanged(object sender, EventArgs e)
        {
            if (radReceiving.Checked == true || radPM.Checked==true||radSend.Checked==true)
            {
                if (txtOpSN.Text.Length >= 5)
                {
                    for (int i = 0; i < MasterList[0].Count; i++)
                    {
                        if (txtOpSN.Text == MasterList[0][i])
                        {
                            btnOpConfirm.Enabled = true;
                            lblNoMatchWarning.Visible = false;

                            if (radSend.Checked != true)
                            {
                                txtOpCustomer.Text = MasterList[3][i];
                                txtOpAccnt.Text = MasterList[14][i];
                            }
                            

                            if (radReceiving.Checked == true)
                            {
                                txtOpRepair.Text = MasterList[5][i];
                                txtOpCustomer.Text = MasterList[3][i];
                                txtOpAccnt.Text = MasterList[14][i];
                                txtOpLoaner.Text = MasterList[4][i];
                            }
                            txtOpNotes.Text = MasterList[13][i];
                            return;
                        }
                    }

                    txtOpAccnt.Clear();
                    txtOpCustomer.Clear();
                    txtOpLoaner.Clear();
                    txtOpRepair.Clear();
                    txtOpSW.Clear();
                    txtOpNotes.Clear();
                    chkOpRepair.Checked = false;
                    btnOpConfirm.Enabled = false;
                    lblNoMatchWarning.Text = "No Loaner matches that serial number";
                    lblNoMatchWarning.Visible = true;
                }
                else
                {
                    btnOpConfirm.Enabled = false;
                    lblNoMatchWarning.Visible = true;
                    lblNoMatchWarning.Text = "Serial number too short";
                }
            }
        }

        private void txtOpSN_Leave(object sender, EventArgs e)
        {
            if (radReceiving.Checked == true)
            {
                for (int i = 0; i < MasterList[0].Count; i++)
                {
                    if (txtOpSN.Text == MasterList[0][i])
                    {
                        txtOpCustomer.Text = MasterList[3][i];
                        txtOpAccnt.Text = MasterList[14][i];
                        txtOpLoaner.Text = MasterList[4][i];
                        txtOpRepair.Text = MasterList[5][i];
                        txtOpNotes.Text = MasterList[13][i];
                        break;
                    }
                }
            }
        }

        private void ClearOpFields()
        {
            txtOpSN.Clear();
            txtOpAccnt.Clear();
            txtOpCustomer.Clear();
            txtOpLoaner.Clear();
            txtOpRepair.Clear();
            txtOpSW.Clear();
            txtOpNotes.Clear();
            chkOpRepair.Checked = false;
            //txtCopy.Clear();
        }

        private void CommaOpFields()
        {
            txtOpSN.Text = txtOpSN.Text.Replace(",", " ");
            txtOpAccnt.Text = txtOpAccnt.Text.Replace(",", " ");
            txtOpCustomer.Text = txtOpCustomer.Text.Replace(",", " ");
            txtOpLoaner.Text = txtOpLoaner.Text.Replace(",", " ");
            txtOpRepair.Text = txtOpRepair.Text.Replace(",", " ");
            txtOpSW.Text = txtOpSW.Text.Replace(",", " ");
            txtOpNotes.Text = txtOpNotes.Text.Replace(",", " ");
        }

        private void txtOpAccnt_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < ContactList[0].Count; i++)
            {
                if (txtOpAccnt.Text == ContactList[0][i])
                {
                    txtOpCustomer.Text = ContactList[2][i];
                    lblNoMatch.Visible = false;
                    return;
                }
            }
            txtOpCustomer.Text = "";
            lblNoMatch.Visible = true;
        }

        private void radCustNew_CheckedChanged(object sender, EventArgs e)
        {
            if (radCustNew.Checked == true)
            {
                Util.Animate(lblCustOp, Util.Effect.Slide, 80, 90);
                lblCustOp.Text = "New Customer";
                Util.Animate(lblCustOp, Util.Effect.Slide, 80, 90);

                ClearCustomerFields();
                grpCustFill.Enabled = true;

                if (lblCustHelp.Visible == true)
                {
                    Util.Animate(lblCustHelp, Util.Effect.Slide, 80, 0);
                    lblCustHelp.Visible = false;
                }

                txtCustNum.Focus();
                btnCustConfirm.Text = "Create";
            }
        }

        private void radCustEdit_CheckedChanged(object sender, EventArgs e)
        {
            if (radCustEdit.Checked == true)
            {
                Util.Animate(lblCustOp, Util.Effect.Slide, 80, 90);
                lblCustOp.Text = "Edit Customer";
                Util.Animate(lblCustOp, Util.Effect.Slide, 80, 90);
                chkCustAccnt.Checked = false;
                grpCustFill.Enabled = true;

                if (lblCustHelp.Visible == false)
                {
                    Util.Animate(lblCustHelp, Util.Effect.Slide, 80, 0);
                    lblCustHelp.Visible = true;
                }

                txtCustNum.Focus();
                btnCustConfirm.Text = "Confirm";
            }
        }

        private void radCustDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (radCustDelete.Checked == true)
            {
                Util.Animate(lblCustOp, Util.Effect.Slide, 80, 90);
                lblCustOp.Text = "Delete Customer";
                Util.Animate(lblCustOp, Util.Effect.Slide, 80, 90);
                chkCustAccnt.Checked = false;
                grpCustFill.Enabled = true;

                if (lblCustHelp.Visible == false)
                {
                    Util.Animate(lblCustHelp, Util.Effect.Slide, 80, 0);
                    lblCustHelp.Visible = true;
                }

                txtCustNum.Focus();
                btnCustConfirm.Text = "DELETE";
            }
        }

        private void ClearCustomerFields()
        {
            txtCustNum.Clear();
            txtCustAccount.Clear();
            chkCustAccnt.Checked = true;
            txtCustName.Clear();
            txtCustContact.Clear();
            txtCustPhone.Clear();
            txtCustEmail.Clear();
            //lblSuccess2.Visible = false;
        }

        private void CommaCustomerFields()
        {
            txtCustNum.Text = txtCustNum.Text.Replace(",", " ");
            txtCustAccount.Text = txtCustAccount.Text.Replace(",", " ");
            txtCustName.Text = txtCustName.Text.Replace(",", " ");
            txtCustContact.Text = txtCustContact.Text.Replace(",", " ");
            txtCustPhone.Text=txtCustPhone.Text.Replace(",", " ");
            txtCustEmail.Text=txtCustEmail.Text.Replace(",", " ");
        }

        private void NewCustomer(string entry)
        {
            try
            {
                var NewEntry = new StreamWriter(@"T:\Databases\c_list.dbs", true);
                NewEntry.WriteLine(entry);

                NewEntry.Close();
                
                UpdateCustomerList();
                ClearCustomerFields();
                txtCustNum.Focus();
            }
            catch
            {
                MessageBox.Show("Customers list is currently open. Try again in a bit.");
                return;
            }
        }



        private void btnCustConfirm_Click(object sender, EventArgs e)
        {
            CommaCustomerFields();
            try
            {
                UpdateCustomerList();
            }
            catch
            {
                MessageBox.Show("Getting latest database list failed, please try again in a bit.");
                return;
            }

            if (txtCustNum.Text==""||txtCustName.Text==""||txtCustAccount.Text=="")
            {
                MessageBox.Show("Please fill all mandatory fields");
                return;
            }

            if (radCustNew.Checked==true)
            {
                try
                {
                    NewCustomer(txtCustNum.Text + "," + txtCustAccount.Text + "," + txtCustName.Text + "," + txtCustContact.Text + "," + txtCustPhone.Text + "," + txtCustEmail.Text);
                    //MessageBox.Show("Customer Entry created");

                    ClearCustomerFields();

                    lblSuccess2.Text = "New Customer successful";
                    Util.Animate(lblSuccess2, Util.Effect.Slide, 90, 90);
                    timer1.Enabled = true;
                    timer1.Start();
                }
                catch
                {
                    MessageBox.Show("Customers list is currently open. Try again in a bit.");
                }
            }

            if (radCustEdit.Checked == true)
            {
                try
                {
                    for (int i = 0; i < ContactList[0].Count; i++)
                    {
                        if (txtCustNum.Text == ContactList[0][i])
                        {
                            ContactList[1][i] = txtCustAccount.Text;
                            ContactList[2][i] = txtCustName.Text;
                            ContactList[3][i] = txtCustContact.Text;
                            ContactList[4][i] = txtCustPhone.Text;
                            ContactList[5][i] = txtCustEmail.Text;
                            break;
                        }
                    }

                    var database = new StreamWriter(@"T:\Databases\c_list.dbs");

                    for (int i = 0; i < ContactList[0].Count; i++)
                    {
                        string entry = "";

                        entry = ContactList[0][i] + "," + ContactList[1][i] + "," + ContactList[2][i] + "," + ContactList[3][i] + "," + ContactList[4][i] + "," + ContactList[5][i];
                        database.WriteLine(entry);
                    }
                    database.Close();
                    ClearCustomerFields();
                    UpdateCustomerList();
                    //MessageBox.Show("Edit Successful");
                    lblSuccess2.Text = "Edit Customer successful";
                    Util.Animate(lblSuccess2, Util.Effect.Slide, 90, 90);
                    timer1.Enabled = true;
                    timer1.Start();
                    txtCustNum.Focus();
                }
                catch
                {
                    MessageBox.Show("Customers list is currently open. Try again in a bit.");
                    return;
                }
            }

            if (radCustDelete.Checked == true)
            {
                try { 
                    var database = new StreamWriter(@"T:\Databases\c_list.dbs");

                    for (int i = 0; i < ContactList[0].Count; i++)
                    {
                        string entry = "";
                        if (ContactList[0][i] != txtCustNum.Text)
                        {
                            entry = ContactList[0][i] + "," + ContactList[1][i] + "," + ContactList[2][i] + "," + ContactList[3][i] + "," + ContactList[4][i] + "," + ContactList[5][i];
                            database.WriteLine(entry);
                        }
                    }

                    database.Close();
                    ClearCustomerFields();
                    UpdateCustomerList();
                    //MessageBox.Show("Delete Successful");
                    lblSuccess2.Text = "Delete Customer successful";
                    Util.Animate(lblSuccess2, Util.Effect.Slide, 90, 90);
                    timer1.Enabled = true;
                    timer1.Start();
                    txtCustNum.Focus();
                }
                catch
                {
                    MessageBox.Show("Customers list is currently open. Try again in a bit.");
                    return;
                }
            }
            

        }

        private void txtCustNum_TextChanged(object sender, EventArgs e)
        {
            if (chkCustAccnt.Checked == true)
            {
                txtCustAccount.Text = txtCustNum.Text;
            }

            if (txtCustNum.Text != "")
            {
                btnCustConfirm.Enabled = true;
                if (radCustEdit.Checked == true || radCustDelete.Checked == true)
                {
                    for (int i = 0; i < ContactList[0].Count; i++)
                    {
                        if (txtCustNum.Text == ContactList[0][i])
                        {
                            chkCustAccnt.Checked = false;
                            txtCustAccount.Text = ContactList[1][i];
                            txtCustName.Text = ContactList[2][i];
                            txtCustContact.Text = ContactList[3][i];
                            txtCustPhone.Text = ContactList[4][i];
                            txtCustEmail.Text = ContactList[5][i];
                            return;
                        }
                       
                    }
                    txtCustAccount.Clear();
                    chkCustAccnt.Checked = false;
                    txtCustName.Clear();
                    txtCustContact.Clear();
                    txtCustPhone.Clear();
                    txtCustEmail.Clear();
                }
            }
            else
            {
                btnCustConfirm.Enabled = false;
            }
        }

        private void chkCustAccnt_CheckedChanged(object sender, EventArgs e)
        {
            if (chkCustAccnt.Checked == true)
            {
                txtCustAccount.Text = txtCustNum.Text;
            }
            else
            {
                txtCustAccount.Clear();
            }
        }

        private void txtCustNum_Leave(object sender, EventArgs e)
        {
            if (txtCustNum.Text == "")
            {
                return;
            }

            if (radCustNew.Checked==true)
            {
                for (int i = 0; i < ContactList[0].Count; i++)
                {
                    if (txtCustNum.Text == ContactList[0][i])
                    {
                        MessageBox.Show("Customer Number already exists! Switching to Edit Mode");
                        //lblCustOp.Text = "Edit Customer";
                        radCustEdit.Checked = true;
                        string temp =txtCustNum.Text;
                        txtCustNum.Text = "";
                        txtCustNum.Text = temp;
                    }
                }
            }

            
        }

        private void txtCustNum_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
        }

        private void btnCustClear_Click(object sender, EventArgs e)
        {
            ClearCustomerFields();
        }

        private void txtNewCustNum_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < ContactList[0].Count; i++)
            {
                if (txtNewCustNum.Text == ContactList[0][i])
                {
                    txtNewCustomer.Text = ContactList[2][i];
                    
                    return;
                }
            }
            txtNewCustomer.Text = "";
        }

        private void lblInfoStatus_Click(object sender, EventArgs e)
        {

        }

        private void ClearInfoFields()
        {
            lblInfoPN.Text = "...";
            lblInfoModel.Text = "...";
            lblInfoLocation.Text = "...";
            lblInfoCustNum.Text = "...";
            lblInfoCustName.Text = "...";
            lblInfoLoaner.Text = "...";
            lblInfoRepair.Text = "...";
            lblInfoDateSent.Text = "...";
            lblInfoDateReturn.Text = "...";
            lblInfoLastPM.Text = "...";
            lblInfoStatus.Text = "...";
            lblInfoNotes.Text = "...";
            lblInfoSW.Text = "...";
            lblInfoOut.Text = "...";
            lblInfoContact.Text = "...";
            lblInfoPhone.Text = "...";
            lblInfoEmail.Text = "...";
            lblInfoAccount.Text = "...";
            lblInfoStatus.ForeColor = System.Drawing.Color.Black;
            lblInfoStatusLabel.ForeColor = System.Drawing.Color.Black;
            btnConfig.Visible = false;
            btnRequestEmail.Visible = false;
            lblInfoPossesion.Text = "";
            listInfoPossess.Items.Clear();
            btnEmail.Visible = false;
            btnInfoEdit.Visible = false;
            btnPrintBarcode.Visible = false;
        }


        private void txtInfoSN_TextChanged(object sender, EventArgs e)
        {
            if (txtInfoSN.Text.Length >= 5)
            {
                ClearInfoFields();

                for (int i = 0; i < MasterList[0].Count; i++)
                {
                    if (txtInfoSN.Text == MasterList[0][i])
                    {
                        lblInfoPN.Text = MasterList[1][i];
                        btnConfig.Visible = true;
                        btnPrintBarcode.Visible = true;
                        lblInfoModel.Text = MasterList[9][i];

                        if (login == 0)
                        {
                            btnRequestEmail.Visible = true;
                            btnRequestEmail.Text = "Login to enable actions";
                        }

                        if (login == 1||login==3)
                        {
                            btnInfoEdit.Visible = true;
                            
                        }
                        else
                        {
                            btnInfoEdit.Visible = false;
                            
                        }

                        if (login == 2 || login == 3)
                        {
                            btnInfoEditCust.Visible = true;
                        }
                        else
                        {
                            btnInfoEditCust.Visible = false;
                        }

                        //lblInfoCustNum.Text = MasterList[14][i];
                        txtInfoAccount.Text = "";
                        txtInfoAccount.Text=MasterList[14][i];
                        //PopulateCustomerInfo();

                        /*
                        if (MasterList[14][i] != "" && MasterList[14][i] != "#N/A")
                        {
                            for (int j = 0; j < ContactList[0].Count; j++)
                            {
                                if (ContactList[0][j] == txtInfoAccount.Text)
                                {
                                    lblInfoAccount.Text = ContactList[1][j];
                                    lblInfoContact.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ContactList[3][j].ToLower());
                                    lblInfoPhone.Text = ContactList[4][j];
                                    lblInfoEmail.Text = ContactList[5][j].ToLower();
                                    if (login == 2 || login == 3)
                                    {
                                        if (lblInfoEmail.Text == "UNK" || lblInfoEmail.Text == "" || lblInfoEmail.Text == "...")
                                        {
                                            btnEmail.Visible = false;
                                        }
                                        else
                                        {
                                            btnEmail.Visible = true;
                                        }
                                    }
                                    break;
                                }
                            }
                        }


                        for (int k = 0; k < MasterList[14].Count; k++)
                        {
                            if (MasterList[14][k] == lblInfoCustNum.Text && MasterList[2][k] == "OUT")
                            {
                                if (MasterList[0][k] == txtInfoSN.Text)
                                {
                                    listInfoPossess.Items.Add(MasterList[0][k] + " - Currently Viewing");
                                }
                                else
                                {
                                    listInfoPossess.Items.Add(MasterList[0][k]);
                                }
                            }
                        }
                        lblInfoPossesion.Text = listInfoPossess.Items.Count.ToString() + " Loaner(s) in possession";
                        */
                            if (MasterList[2][i] == "IN")
                            {
                                lblInfoLocation.Text = "In-house";
                                if (MasterList[11][i] == "Y")
                                {
                                    lblInfoStatus.Text = "Ready to ship";
                                    if (login == 1 || login == 3)
                                    {
                                        btnRequestEmail.Visible = true;
                                        btnRequestEmail.Text = "Send out this loaner now";
                                    }
                                    lblInfoStatus.ForeColor = System.Drawing.Color.Green;
                                    lblInfoStatusLabel.ForeColor = System.Drawing.Color.Green;
                                }
                                if (MasterList[11][i] == "N")
                                {
                                    lblInfoStatus.Text = "Not prepared";
                                    if (login == 1 || login == 3)
                                    {
                                        btnRequestEmail.Visible = true;
                                        btnRequestEmail.Text = "Update PM status";
                                    }
                                    if (MasterList[12][i] == "Y")
                                    {
                                        lblInfoStatus.Text = "Requires a repair";
                                        btnRequestEmail.Visible = true;
                                        btnRequestEmail.Text = "Request a Repair SR";
                                        lblInfoStatus.ForeColor = System.Drawing.Color.Red;
                                        lblInfoStatusLabel.ForeColor = System.Drawing.Color.Red;
                                    }
                                }
                            }
                            else if(MasterList[2][i] == "OUT")
                            {
                                lblInfoLocation.Text = "Out";
                                DateTime parsedDate;
                                string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };

                                DateTime.TryParseExact(MasterList[6][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                                lblInfoOut.Text = (DateTime.Today - parsedDate).TotalDays.ToString();

                                if (DateTime.Today.AddMonths(-1).CompareTo(parsedDate) > 0)
                                {
                                    lblInfoStatus.Text = "OVERDUE";
                                    if (login == 2 || login == 3)
                                    {
                                        btnRequestEmail.Visible = true;
                                        btnRequestEmail.Text = "Send Overdue Reminder Email";
                                    }
                                    lblInfoStatus.ForeColor = System.Drawing.Color.Red;
                                    lblInfoStatusLabel.ForeColor = System.Drawing.Color.Red;
                                }
                                else
                                {
                                    lblInfoStatus.Text = "Out within the last 30 Days";
                                    lblInfoStatus.ForeColor = System.Drawing.Color.Black;
                                    lblInfoStatusLabel.ForeColor = System.Drawing.Color.Black;
                                }
                            }else if(MasterList[2][i] == "PERMA")
                        {
                            lblInfoLocation.Text = "Permanent/Long Term Loaner";
                            lblInfoStatus.Text = "Long Term";
                            lblInfoStatus.ForeColor = System.Drawing.Color.Black;
                            lblInfoStatusLabel.ForeColor = System.Drawing.Color.Black;

                            DateTime parsedDate;
                            string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };

                            DateTime.TryParseExact(MasterList[6][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                            lblInfoOut.Text = (DateTime.Today - parsedDate).TotalDays.ToString();
                        }

                        
                        lblInfoLoaner.Text = MasterList[4][i];
                        lblInfoRepair.Text = MasterList[5][i];
                        lblInfoDateSent.Text = MasterList[6][i];
                        lblInfoDateReturn.Text = MasterList[7][i];
                        lblInfoLastPM.Text = MasterList[8][i];
                        lblInfoNotes.Text = MasterList[13][i];
                        lblInfoSW.Text = MasterList[10][i];

                        return;
                    }
                }
                ClearInfoFields();                
            }
            else
            {
                ClearInfoFields();
            }
        }

        private void DisplayDetails()
        {
                txtInfoSN.Text = listFields.Items[listFields.SelectedIndices[0]].Text;
                Util.Animate(tabControl1, Util.Effect.Slide, 150, 0);
                tabControl1.SelectedIndex = 1;
                Util.Animate(tabControl1, Util.Effect.Slide, 150, 180);
        }

        private void listSN_DoubleClick(object sender, EventArgs e)
        {
            DisplayDetails();
        }

        private void SendOverdueEmail()
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(lblInfoEmail.Text);
                oRecip.Resolve();

                if (lblInfoStatus.Text != "OVERDUE")
                {
                    oMsg.Subject = "Inquiring about " + lblInfoModel.Text + " Loaner SN: " + txtInfoSN.Text;
                    //oMsg.Body = "Hi " + lblInfoContact.Text + ",\n\nYour " + lblInfoModel.Text + " Loaner with Serial Number: " + txtInfoSN.Text + " is now " + lblInfoOut.Text + " days overdue.\n\nPlease have this unit returned to ZOLL Canada as soon as you can.\n\nThanks,";
                }
                else
                {
                    oMsg.Subject = "OVERDUE " + lblInfoModel.Text + " Loaner SN: " + txtInfoSN.Text;
                    oMsg.Body = "Hi " + lblInfoContact.Text.Split(' ')[0] + ",\n\nYour " + lblInfoModel.Text + " Loaner with Serial Number: " + txtInfoSN.Text + " is now " + (Int32.Parse(lblInfoOut.Text) - 30).ToString() + " days overdue.\n\nPlease have this unit returned to ZOLL Canada as soon as you can.\n\nThanks,";
                }
                //oMsg.Attachments.Add("c:/temp/test.txt", Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                oMsg.Display(true);
            }
            catch
            {
                MessageBox.Show("Email function failed. Make sure Outlook is open");
            }
        }

        private void SendEmail()
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(lblInfoEmail.Text);
                oRecip.Resolve();

                oMsg.Subject = lblInfoModel.Text + " Loaner SN " + txtInfoSN.Text;
                oMsg.Body = "Hi " + lblInfoContact.Text.Split(' ')[0] + ",\n\n";
               
                oMsg.Display(true);
            }
            catch
            {
                MessageBox.Show("Email function failed. Make sure Outlook is open");
            }
        }

        private void btnEmail_Click(object sender, EventArgs e)
        {
            SendEmail();
        }

        private void ShowRepair()
        {
            ClearFields();
            filter = 6;

            for (int i = 0; i < listM.Count(); i++)
            {
                if (listM[i] == "Y")
                {
                    ListViewItem lvi = new ListViewItem(MasterList[0][i]);

                    for (int j = 1; j < MasterList.Count; j++)
                    {
                        if (j < 6 || j > 8)
                        {
                            lvi.SubItems.Add(MasterList[j][i]);
                        }
                        else
                        {
                            if (MasterList[j][i] != "")
                            {
                                DateTime parsedDate;
                                string[] formatstring = { "d/m/yyyy", "dd/m/yyyy", "d/mm/yyyy", "dd/mm/yyyy" };

                                DateTime.TryParseExact(MasterList[j][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                                lvi.SubItems.Add(parsedDate.ToString("yyyy/mm/dd"));
                            }
                            else
                            {
                                lvi.SubItems.Add("");
                            }
                        }
                        //lvi.SubItems.Add(MasterList[j][i]);
                    }
                    listFields.Items.Add(lvi);
                }
            }
        }

        private void btnRepair_Click(object sender, EventArgs e)
        {
            ShowRepair();
            SearchData();
        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            txtConfig.Text = lblInfoPN.Text;
            
            Util.Animate(tabControl1, Util.Effect.Slide, 150, 0);
            tabControl1.SelectedIndex = 2;
            FindConfig(txtConfig.Text);
            FindExactItemMatches(txtConfig.Text);
            Util.Animate(tabControl1, Util.Effect.Slide, 150, 180);
        }

        private void SendRepairSREmail()
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("canadatechsupport@zoll.com");
                oRecip.Resolve();

                oMsg.Subject = "Repair SR for " + lblInfoModel.Text + " Loaner SN: " + txtInfoSN.Text;
                oMsg.Body = "Hi CanadaTechSupport,\n\nThe " + lblInfoModel.Text + " Loaner with SN: " + txtInfoSN.Text + " will require a repair. The problem was found to be \"" + lblInfoNotes.Text + "\".\n\nCould a Repair SR please be created for this Loaner?\n\nThanks,";

                //oMsg.Attachments.Add("c:/temp/test.txt", Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                oMsg.Display(true);
            }
            catch
            {
                MessageBox.Show("Email function fucked up bad");
            }

        }

        private void btnRequestEmail_Click(object sender, EventArgs e)
        {
            if(btnRequestEmail.Text=="Login to enable actions")
            {
                if (tabControl1.SelectedIndex > 3)
                {
                    Util.Animate(tabControl1, Util.Effect.Slide, 150, 180);
                    currentindex = tabControl1.SelectedIndex;
                }
                if (tabControl1.SelectedIndex < 3)
                {
                    Util.Animate(tabControl1, Util.Effect.Slide, 150, 0);

                    currentindex = tabControl1.SelectedIndex;
                }
                logbuttpushed = true;
                tabControl1.SelectedIndex = 3;

                Util.Animate(tabControl1, Util.Effect.Slide, 150, 90);
                tabControl1.Visible = true;
                logbuttpushed = true;
                txtNewLogin.Focus();
                return;
            }

            if (lblInfoStatus.Text == "OVERDUE")
            {
                SendOverdueEmail();
            }
            if (lblInfoStatus.Text == "Requires a repair")
            {
                SendRepairSREmail();
            }
            if (lblInfoStatus.Text == "Ready to ship")
            {
                SwitchtoSend(txtInfoSN.Text, lblInfoLoaner.Text);
            }
            if (lblInfoStatus.Text == "Not prepared")
            {
                SwitchtoPM(txtInfoSN.Text);
            }
        }

        private void SwitchtoPM(string sn)
        {
            Util.Animate(tabControl1, Util.Effect.Slide, 150, 0);
            tabControl1.SelectedIndex = 4;
            radPM.Checked = true;
            txtOpSN.Text = sn;
            Util.Animate(tabControl1, Util.Effect.Slide, 150, 180);
        }

        private void SwitchtoSend(string sn, string loanersr)
        {
            Util.Animate(tabControl1, Util.Effect.Slide, 150, 0);
            tabControl1.SelectedIndex = 4;
            radSend.Checked = true;
            txtOpSN.Text = sn;
            txtOpLoaner.Text = loanersr;
            Util.Animate(tabControl1, Util.Effect.Slide, 150, 180);
        }

        private void listInfoPossess_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listInfoPossess.SelectedIndex >= 0)
            {
                var values = listInfoPossess.Items[listInfoPossess.SelectedIndex].ToString().Split(' ');

                ClearInfoFields();
                txtInfoSN.Clear();
                txtInfoSN.Text = values[0];
            }
        }
        
        private void cmbNewStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbNewStatus.SelectedIndex == 1|| cmbNewStatus.SelectedIndex == 2)
            {
                cmbNewReady.SelectedIndex = 1;
                cmbNewRepair.SelectedIndex = 1;
            }
            if (cmbNewStatus.SelectedIndex == 0)
            {
                cmbNewReady.SelectedIndex = -1;
                cmbNewRepair.SelectedIndex = -1;
            }
        }

        private void calStart_DateChanged(object sender, DateRangeEventArgs e)
        {
            txtStartDate.Text = calStart.SelectionStart.ToString("dd/MM/yyyy");
        }

        private void chkFull_CheckedChanged(object sender, EventArgs e)
        {
            if (chkFull.Checked == true)
            {
                chkReceive.Checked = true;
                chkSend.Checked = true;
                chkCreate.Checked = true;
                chkEdit.Checked = true;
                chkDelete.Checked = true;
                chkPM.Checked = true;
                grpSentReceive.Enabled = false;
            }
            else
            {
                chkReceive.Checked = false;
                chkSend.Checked = false;
                chkCreate.Checked = false;
                chkEdit.Checked = false;
                chkDelete.Checked = false;
                chkPM.Checked = false;
                grpSentReceive.Enabled = true ;
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (txtEndDate.Text == "" || txtStartDate.Text == "")
            {
                MessageBox.Show("Please input start and end dates to run the report.");
                return;
            }
            
            string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };
            DateTime StartDate;
            DateTime EndDate;
            DateTime parsedDate;

            DateTime.TryParseExact(txtStartDate.Text, formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out StartDate);
            DateTime.TryParseExact(txtEndDate.Text, formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out EndDate);

            string record = "record.dbs";
            List<string> RecA = new List<string>();
            List<string> RecB = new List<string>();
            List<string> RecC = new List<string>();
            List<string> RecD = new List<string>();
            List<string> RecE = new List<string>();
            List<string> RecF = new List<string>();
            List<string> RecG = new List<string>();
            List<string> RecH = new List<string>();
            List<string> RecI = new List<string>();
            List<string> RecJ = new List<string>();
            List<string> RecK = new List<string>();

            var reader = new StreamReader(record);
            string dummy;
            bool first = true;

            while (!reader.EndOfStream)
            {
                if (first == true)
                {
                    dummy = reader.ReadLine();
                    first = false;
                }
                else
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');
                    DateTime.TryParseExact(values[0], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                    if (StartDate.CompareTo(parsedDate) <= 0 && EndDate.CompareTo(parsedDate)>=0)
                    {
                        if (values[1] == "SEND" && chkSend.Checked==true)
                        {
                            RecA.Add(values[0]);
                            RecB.Add(values[1]);
                            RecC.Add(values[2]);
                            RecD.Add(values[3]);
                            RecE.Add(values[4]);
                            RecF.Add(values[5]);
                            RecG.Add(values[6]);
                            RecH.Add(values[7]);
                            RecI.Add(values[8]);
                            RecJ.Add(values[9]);
                            RecK.Add(values[10]);
                        }

                        if (values[1] == "PM" && chkPM.Checked == true)
                        {
                            RecA.Add(values[0]);
                            RecB.Add(values[1]);
                            RecC.Add(values[2]);
                            RecD.Add(values[3]);
                            RecE.Add(values[4]);
                            RecF.Add(values[5]);
                            RecG.Add(values[6]);
                            RecH.Add(values[7]);
                            RecI.Add(values[8]);
                            RecJ.Add(values[9]);
                            RecK.Add(values[10]);
                        }

                        if (values[1] == "RECEIVE" && chkReceive.Checked == true)
                        {
                            RecA.Add(values[0]);
                            RecB.Add(values[1]);
                            RecC.Add(values[2]);
                            RecD.Add(values[3]);
                            RecE.Add(values[4]);
                            RecF.Add(values[5]);
                            RecG.Add(values[6]);
                            RecH.Add(values[7]);
                            RecI.Add(values[8]);
                            RecJ.Add(values[9]);
                            RecK.Add(values[10]);
                        }

                        if (values[1] == "CREATE" && chkCreate.Checked == true)
                        {
                            RecA.Add(values[0]);
                            RecB.Add(values[1]);
                            RecC.Add(values[2]);
                            RecD.Add(values[3]);
                            RecE.Add(values[4]);
                            RecF.Add(values[5]);
                            RecG.Add(values[6]);
                            RecH.Add(values[7]);
                            RecI.Add(values[8]);
                            RecJ.Add(values[9]);
                            RecK.Add(values[10]);
                        }

                        if (values[1] == "EDIT" && chkEdit.Checked == true)
                        {
                            RecA.Add(values[0]);
                            RecB.Add(values[1]);
                            RecC.Add(values[2]);
                            RecD.Add(values[3]);
                            RecE.Add(values[4]);
                            RecF.Add(values[5]);
                            RecG.Add(values[6]);
                            RecH.Add(values[7]);
                            RecI.Add(values[8]);
                            RecJ.Add(values[9]);
                            RecK.Add(values[10]);
                        }

                        if (values[1] == "DELETE" && chkDelete.Checked == true)
                        {
                            RecA.Add(values[0]);
                            RecB.Add(values[1]);
                            RecC.Add(values[2]);
                            RecD.Add(values[3]);
                            RecE.Add(values[4]);
                            RecF.Add(values[5]);
                            RecG.Add(values[6]);
                            RecH.Add(values[7]);
                            RecI.Add(values[8]);
                            RecJ.Add(values[9]);
                            RecK.Add(values[10]);
                        }
                    }
                }
            }
            reader.Close();
            string reportname = "Movement Report.csv";

            try
            {
                var writer = new StreamWriter(reportname);
                writer.WriteLine("Date,User,Operation,Duration Since Last Movement,Model,Serial,Customer#,Customer Name,Loaner SR, Complaint SR,Notes");


                for (int i = 0; i < RecA.Count; i++)
                {
                    writer.WriteLine(RecA[i] + "," + RecJ[i] + "," + RecB[i] + "," + RecI[i] + ","+RecK[i]+"," + RecC[i] + "," + RecD[i] + "," + RecE[i] + "," + RecF[i] + "," + RecG[i] + "," + RecH[i]);
                }
                writer.Close();

                Process.Start(reportname);
            }
            catch
            {
                MessageBox.Show("Movement Report.csv is currently open elsewhere. Try again later.");
                return;
            }
        }

        private void calEnd_DateChanged(object sender, DateRangeEventArgs e)
        {
            txtEndDate.Text = calEnd.SelectionStart.ToString("dd/MM/yyyy");
        }

        private void lblInfoEmail_Click(object sender, EventArgs e)
        {
            
        }

        private void EnableManagement()
        {
            
            Util.Animate(grpLogin, Util.Effect.Slide, 150, 270);
            grpLogin.Visible = false;
            

            if (currentindex < tabControl1.SelectedIndex)
            {
                Util.Animate(tabControl1, Util.Effect.Slide, 150, 90);
                tabControl1.SelectedIndex = currentindex;
                Util.Animate(tabControl1, Util.Effect.Slide, 150, 0);
                tabControl1.Visible = true;
            }
            else if(currentindex>tabControl1.SelectedIndex)
            {
                Util.Animate(tabControl1, Util.Effect.Slide, 150, 90);
                tabControl1.SelectedIndex = currentindex;
                Util.Animate(tabControl1, Util.Effect.Slide, 150, 180);
                tabControl1.Visible = true;
            }
            else if (currentindex == tabControl1.SelectedIndex)
            {
                Util.Animate(tabControl2, Util.Effect.Slide, 150, 90);
                
                Util.Animate(grpLoanerOp, Util.Effect.Slide, 150, 90);
                
            }
            tabControl2.Visible = true;
            
            //tabControl2.Top = 11;
            //tabControl2.Left = 7;
            //grpLoanerOp.Top = 11;
            //grpLoanerOp.Left = 11;
            
            
            btnLogOut.Text = "Log Out";

            if (login == 1)
            {
                grpLoanerFunction.Enabled = true;
                grpOperation.Enabled = true;
                lblLocked.Visible = true;
                lblLocked2.Visible = false;
                lblLocked3.Visible = false;
                lblLocked4.Visible = true;
                lblLocked5.Visible = true;
                grpLoanerOp.Visible = true;
                grpLoanerFunction.Visible = true;
                grpForms.Visible = true;
                grpCustOp.Visible = false;
                grpCustFill.Visible = false;
                grpRequest.Visible = false;
            }

            if (login == 2)
            {
                grpLoanerFunction.Enabled = false;
                grpOperation.Enabled = false;
                grpOperations.Enabled = false;
                lblLocked.Visible = true;
                lblLocked2.Visible = true;
                lblLocked3.Visible = true;
                lblLocked4.Visible = false;
                lblLocked5.Visible = false;
                grpLoanerOp.Visible = false;
                grpLoanerFunction.Visible = false;
                grpForms.Visible = false;
                grpCustOp.Visible = true;
                grpCustFill.Visible = true;
                grpRequest.Visible =true;
            }

            if (login == 3)
            {
                grpLoanerFunction.Enabled = true;
                grpOperation.Enabled = true;
                grpUserManage.Visible = true;
                lblLocked.Visible = false;
                lblLocked2.Visible = false ;
                lblLocked3.Visible = false;
                lblLocked4.Visible = false;
                lblLocked5.Visible = false;
                grpLoanerOp.Visible = true;
                grpLoanerFunction.Visible = true;
                grpForms.Visible = true;
                grpCustOp.Visible = true;
                grpCustFill.Visible = true;
                grpRequest.Visible = true;
            }
        }

        private void UserLogin(string user,string password)
        {
            
            for(int i = 0; i < Users.Count; i++)
            {
                if(user==Users[i] && password == PW[i])
                {

                    var autowriter = new StreamWriter(@"C:\Users\Public\loginfile");
                    autowriter.WriteLine(user);
                    autowriter.WriteLine(password);
                    autowriter.Close();

                    login = Int32.Parse(Access[i]);
                    lblLogin.Text = "Currently logged in as \"" + Users[i]+"\"";
                    currentuser = Users[i];
                    EnableManagement();
                    btnChangePass.Visible =true;

                    txtNewLogin.Text = "";
                    txtNewPassword.Text = "";

                    return;
                }
            }
            MessageBox.Show("Wrong Username and Password");
            txtNewPassword.Text = "";
            txtNewPassword.Focus();

        }

        private void UserLogout()
        {
            var autowriter = new StreamWriter(@"C:\Users\Public\loginfile");
            autowriter.WriteLine("");
            autowriter.WriteLine("");
            autowriter.Close();

            login = 0;
            btnChangePass.Visible = false;
            txtChangePass.Clear();
            currentuser = "";
            string thing =txtInfoSN.Text;
            txtInfoSN.Clear();
            txtInfoSN.Text = thing;
            grpChangePass.Visible = false;
            btnChangePass.Visible = false;

            if(tabControl2.Visible==true)
                Util.Animate(tabControl2, Util.Effect.Slide, 150, 270);
            tabControl2.Visible = false;

            if (grpLoanerOp.Visible == true)
                Util.Animate(grpLoanerOp, Util.Effect.Slide, 150, 270);
            grpLoanerOp.Visible = false;

            if (grpLogin.Visible == true)
                Util.Animate(grpLogin, Util.Effect.Slide, 150, 90);
            grpLogin.Visible = true;

            if (grpUserManage.Visible == true)
                Util.Animate(grpUserManage, Util.Effect.Slide, 90, 0);
            grpUserManage.Visible = false;

            if (grpRequest.Visible == true)
                Util.Animate(grpRequest, Util.Effect.Slide, 90, 0);
            grpRequest.Visible = false;

            lblLogin.Text = "Not logged in";
            btnLogOut.Text = "Log In";
            

            if (lblLocked.Visible == false)
                Util.Animate(lblLocked, Util.Effect.Slide, 90, 0);
            lblLocked.Visible = true;

            if (lblLocked2.Visible == false)
                Util.Animate(lblLocked2, Util.Effect.Slide, 90, 0);
            lblLocked2.Visible = true;

            if (lblLocked3.Visible == false)
                Util.Animate(lblLocked3, Util.Effect.Slide, 90, 0);
            lblLocked3.Visible = true;

            if (lblLocked4.Visible == false)
                Util.Animate(lblLocked4, Util.Effect.Slide, 90, 0);
            lblLocked4.Visible = true;

            if (lblLocked5.Visible == false)
                Util.Animate(lblLocked5, Util.Effect.Slide, 90, 0);
            lblLocked5.Visible = true;

            ClearOpFields();
            ClearNewEntryForm();
        }

        private void btnNewLogin_Click(object sender, EventArgs e)
        {
            UserLogin(txtNewLogin.Text,txtNewPassword.Text);
        }

        private void txtNewPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                UserLogin(txtNewLogin.Text, txtNewPassword.Text);
            }
        }
        

        private void btnLogOut_Click(object sender, EventArgs e)
        {
            if (login == 0 && tabControl1.SelectedIndex!=3)
            {
                if (tabControl1.SelectedIndex > 3)
                {
                    Util.Animate(tabControl1, Util.Effect.Slide, 150, 180);
                    currentindex = tabControl1.SelectedIndex;
                }
                if (tabControl1.SelectedIndex < 3)
                {
                    Util.Animate(tabControl1, Util.Effect.Slide, 150, 0);
                    
                    currentindex = tabControl1.SelectedIndex;
                }
                logbuttpushed = true;
                tabControl1.SelectedIndex = 3;
                
                Util.Animate(tabControl1, Util.Effect.Slide, 150, 90);
                tabControl1.Visible = true;
                logbuttpushed = true;
                
                txtNewLogin.Focus();
            }
            else if(login!=0)
            {
                UserLogout();
            }
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                grpForms.Enabled = false;
                lblEditHelp.Visible = false;
            }
        }

        private void txtStartDate_TextChanged(object sender, EventArgs e)
        {
            if (txtStartDate.Text != "" && txtStartDate.Text != "dd/mm/yyyy" && txtEndDate.Text != "" && txtEndDate.Text != "dd/mm/yyyy")
            {
                btnRun.Enabled = true;
            }
            else
            {
                btnRun.Enabled = false;
            }
        }

        private void txtEndDate_TextChanged(object sender, EventArgs e)
        {
            if (txtStartDate.Text != "" && txtStartDate.Text != "dd/mm/yyyy" && txtEndDate.Text != "" && txtEndDate.Text != "dd/mm/yyyy")
            {
                btnRun.Enabled = true;
            }
            else
            {
                btnRun.Enabled = false;
            }
        }

        private void calStart_DateSelected(object sender, DateRangeEventArgs e)
        {
            txtStartDate.Text = calStart.SelectionStart.ToString("dd/MM/yyyy");
            calEnd.MinDate = calStart.SelectionStart;
        }

        private void calEnd_DateSelected(object sender, DateRangeEventArgs e)
        {
            txtEndDate.Text = calEnd.SelectionStart.ToString("dd/MM/yyyy");
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            //if (tabControl1.SelectedIndex == 3)
            //{
            //    currentindex = 3;
            //}
        }

        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            Process.Start("Help.doc");
        }

        private void btnSummary_Click(object sender, EventArgs e)
        {
            string filename = DateTime.Today.ToString("dd/mm/yyyy").Replace("/", "-") + " - Summary Report.csv";

            var database = new StreamWriter(filename);
            database.WriteLine(btnPool.Text + "," + btnInHouse.Text + "," + btnReady.Text + "," + btnNotReady.Text + "," + btnSent.Text + "," + btnRepair.Text + "," + btnOverdue.Text);
            database.WriteLine("");
            database.WriteLine("Serial,Item Number,Status,Customer Number,Customer Name,Loaner SR,Complaint SR,Date Sent,Date Returned,Date Last PMd,Model,Software,Ready,Requires Repair,Notes");
            for (int i = 0; i < MasterList[0].Count; i++)
            {
                string entry = "";


                entry = MasterList[0][i] + "," + MasterList[1][i] + "," + MasterList[2][i] + ","+MasterList[14][i]+"," + MasterList[3][i] + "," + MasterList[4][i] + "," + MasterList[5][i] + "," + MasterList[6][i] + "," + MasterList[7][i] + "," + MasterList[8][i] + "," + MasterList[9][i] + "," + MasterList[10][i] + "," + MasterList[11][i] + "," + MasterList[12][i] + "," + MasterList[13][i];
                database.WriteLine(entry);
            }
            database.Close();

            Process.Start(filename);
        }

        private void radSummary_CheckedChanged(object sender, EventArgs e)
        {
            if (radSummary.Checked == true)
            {
                if (grpMovement.Visible == true)
                {
                    Util.Animate(grpMovement, Util.Effect.Slide, 180, 270);
                    grpMovement.Visible = false;
                }
                if (grpSummary.Visible == false)
                {
                    
                    Util.Animate(grpSummary, Util.Effect.Slide, 180, 90);
                    grpSummary.Visible = true;
                    //EnableLoanerBreakdownChart();
                    //EnableTrendChart();
                }
            }
        }

        private void radMovement_CheckedChanged(object sender, EventArgs e)
        {
            if (radMovement.Checked == true)
            {
                if (grpSummary.Visible == true)
                {
                    Util.Animate(grpSummary, Util.Effect.Slide, 180, 270);
                    grpSummary.Visible = false;
                }
                if (grpMovement.Visible == false)
                {
                    
                    Util.Animate(grpMovement, Util.Effect.Slide, 180, 90);
                    grpMovement.Visible = true;
                }
            }
        }

        private void radNewEntry_CheckedChanged(object sender, EventArgs e)
        {
            if (radNewEntry.Checked == true)
            {
                Util.Animate(lblNewOp, Util.Effect.Slide, 80, 90);
                lblNewOp.Text = "New Entry";
                Util.Animate(lblNewOp, Util.Effect.Slide, 80, 90);

                ClearNewEntryForm();
                grpForms.Enabled = true;
                if (lblEditHelp.Visible == true)
                {
                    Util.Animate(lblEditHelp, Util.Effect.Slide, 100, 180);
                    lblEditHelp.Visible = false;
                }
                btnNewSubmit.Text = "Create Entry";
            }
        }

        private void radEditEntry_CheckedChanged(object sender, EventArgs e)
        {
            if (radEditEntry.Checked == true)
            {
                Util.Animate(lblNewOp, Util.Effect.Slide, 80, 90);
                lblNewOp.Text = "Edit Entry";
                Util.Animate(lblNewOp, Util.Effect.Slide, 80, 90);

                grpForms.Enabled = true;
                if (lblEditHelp.Visible == false)
                {
                    Util.Animate(lblEditHelp, Util.Effect.Slide, 100, 180);
                    lblEditHelp.Visible = true;
                }
                btnNewSubmit.Text = "Confirm Edit";
            }
        }


        private void radDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (radDelete.Checked == true)
            {
                Util.Animate(lblNewOp, Util.Effect.Slide, 80, 90);
                lblNewOp.Text = "Delete Entry";
                Util.Animate(lblNewOp, Util.Effect.Slide, 80, 90);

                grpForms.Enabled = true;

                if (lblEditHelp.Visible == false)
                {
                    Util.Animate(lblEditHelp, Util.Effect.Slide, 100, 180);
                    lblEditHelp.Visible = true;
                }
                btnNewSubmit.Text = "DELETE";
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (lblSuccess.Visible == true)
            {
                Util.Animate(lblSuccess, Util.Effect.Slide, 150, 270);
            }
            lblSuccess.Visible = false;

            if (lblSuccess2.Visible == true)
            {
                Util.Animate(lblSuccess2, Util.Effect.Slide, 150, 270);
            }
            lblSuccess2.Visible = false;

            if (lblSuccess3.Visible == true)
            {
                Util.Animate(lblSuccess3, Util.Effect.Slide, 150, 270);
            }
            lblSuccess3.Visible = false;

            if (lblSuccess4.Visible == true)
            {
                Util.Animate(lblSuccess4, Util.Effect.Slide, 150, 270);
            }
            lblSuccess4.Visible = false;

            if (lblSuccess5.Visible == true)
            {
                Util.Animate(lblSuccess5, Util.Effect.Slide, 150, 270);
                Util.Animate(btnChangePass, Util.Effect.Slide, 150, 90);
            }
            lblSuccess5.Visible = false;
            btnChangePass.Visible = true;

            timer1.Enabled = false;
        }

        private void btnClearOp_Click(object sender, EventArgs e)
        {
            ClearOpFields();
        }

        private void radPM_Click(object sender, EventArgs e)
        {
            
        }

        private void radSend_Click(object sender, EventArgs e)
        {
            
        }

        private void radReceiving_Click(object sender, EventArgs e)
        {
           
        }

        private void radSend_CheckedChanged(object sender, EventArgs e)
        {
            if (radSend.Checked == true)
            {
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);
                lblOpMode.Text = "Sending Loaner";
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);

                dgvAccessories.Rows.Clear();
                dgvAccessories.Visible = true;
                lblOpAccessories.Visible = true;
                lblOpAccessories.Text = "Requested Accessories";

                lblPending.Visible = true;
                lstPending.Visible = true;
                var pending_reader = new StreamReader(@"T:\\Databases\\PendingLoaners.dbs");
                lstPending.Items.Clear();

                while (!pending_reader.EndOfStream)
                {
                    string line = pending_reader.ReadLine();
                    lstPending.Items.Add(line);
                }
                pending_reader.Close();

                grpOperations.Enabled = true;
                txtCopy.Clear();

                if (grpOpCustomer.Visible == false)
                    Util.Animate(grpOpCustomer, Util.Effect.Slide, 100, 90);
                grpOpCustomer.Visible = true;

                if (grpOpSR.Visible == false)
                    Util.Animate(grpOpSR, Util.Effect.Slide, 100, 90);
                grpOpSR.Visible = true;

                chkOpRepair.Visible = false;
                chkOpRepair.Checked = false;
                txtOpSN.Focus();
            }
        }

        private void radPM_CheckedChanged(object sender, EventArgs e)
        {
            if (radPM.Checked == true)
            {
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);
                lblOpMode.Text = "Finish PM";
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);

                dgvAccessories.Rows.Clear();
                dgvAccessories.Visible = false;
                lblOpAccessories.Visible = false;
                
                lblPending.Visible = false;
                lstPending.Visible = false;

                grpOperations.Enabled = true;
                txtCopy.Clear();

                if (grpOpSR.Visible == true)
                    Util.Animate(grpOpSR, Util.Effect.Slide, 100, 90);
                grpOpSR.Visible = false;

                if (grpOpCustomer.Visible == true)
                    Util.Animate(grpOpCustomer, Util.Effect.Slide, 100, 90);
                grpOpCustomer.Visible = false;



                chkOpRepair.Visible = true;
                chkOpRepair.Text = "Failed PM and Requires repair";
                chkOpRepair.Checked = false;
                txtOpSN.Focus();
            }
        }

        private void radReceiving_CheckedChanged(object sender, EventArgs e)
        {
            if (radReceiving.Checked == true)
            {
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);
                lblOpMode.Text = "Receiving Loaner";
                Util.Animate(lblOpMode, Util.Effect.Slide, 80, 90);

                dgvAccessories.Rows.Clear();
                dgvAccessories.Visible = true;
                lblOpAccessories.Visible = true;
                lblOpAccessories.Text = "Accessories sent out with this loaner";

                lblPending.Visible = false;
                lstPending.Visible = false;

                grpOperations.Enabled = true;
                txtCopy.Clear();

                if (grpOpCustomer.Visible == false)
                    Util.Animate(grpOpCustomer, Util.Effect.Slide, 100, 90);
                grpOpCustomer.Visible = true;

                if (grpOpSR.Visible == false)
                    Util.Animate(grpOpSR, Util.Effect.Slide, 100, 90);
                grpOpSR.Visible = true;
                chkOpRepair.Visible = true;
                chkOpRepair.Checked = false;
                chkOpRepair.Text = "Repair required";
                txtOpSN.Focus();
            }
        }

        private void btnRequestEmail_TextChanged(object sender, EventArgs e)
        {
           
        }
        
        private void btnInfoEdit_Click(object sender, EventArgs e)
        {
            
            Util.Animate(tabControl1, Util.Effect.Slide, 150, 0);
            tabControl1.SelectedIndex = 3;
            tabControl2.SelectedIndex = 0;
            radEditEntry.Checked = true;
            txtNewSN.Text = txtInfoSN.Text;
            Util.Animate(tabControl1, Util.Effect.Slide, 150, 180);
        }

        private void btnSearchClear_Click(object sender, EventArgs e)
        {
            txtSearch.Clear();
            SearchData();
        }

        private void LoadUsers()
        {
            Users.Clear();
            PW.Clear();
            Access.Clear();

            listUsers.Items.Clear();

            var reader = new StreamReader(File.OpenRead(@"T:\Databases\pdata.dbs"));
            //var reader = new StreamReader(File.OpenRead("pdata.dbs"));

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(',');

                Users.Add(values[0]);
                PW.Add(values[1]);
                Access.Add(values[2]);
            }
            reader.Close();

            listUsers.Items.AddRange(Users.ToArray());
        }

        private void radNewUser_CheckedChanged(object sender, EventArgs e)
        {
            if(radNewUser.Checked== true)
            {
                Util.Animate(lblUserOp, Util.Effect.Slide, 80, 90);
                lblUserOp.Text = "New User";
                Util.Animate(lblUserOp, Util.Effect.Slide, 80, 90);

                txtUsername.Clear();
                txtUserPassword.Clear();
                cmbUserAccess.SelectedIndex = -1;
                txtUsername.Enabled = true;
                listUsers.Enabled = false;
                grpUserForm.Enabled = true;
            }
        }

        private void listUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listUsers.SelectedIndex >= 0)
            {
                if (radEditUser.Checked == true || radDeleteUser.Checked == true)
                {
                    txtUsername.Text = listUsers.Items[listUsers.SelectedIndex].ToString();

                    for (int i = 0; i < Users.Count; i++)
                    {
                        if (Users[i] == txtUsername.Text)
                        {
                            txtUserPassword.Text = PW[i];
                            cmbUserAccess.SelectedIndex = (Int32.Parse(Access[i]) - 1);
                            break;
                        }
                    }

                }
            }
        }

        private void btnUserConfirm_Click(object sender, EventArgs e)
        {
            if (txtUsername.Text == "" || txtUserPassword.Text == "" || cmbUserAccess.SelectedIndex < 0 ||txtUsername.Text.Contains(',') ||txtUserPassword.Text.Contains(','))
            {
                MessageBox.Show("Invalid input");
                return;
            }
            LoadUsers();

            if (radNewUser.Checked == true)
            {
                var writer = new StreamWriter(@"T:\Databases\pdata.dbs", true);
                writer.WriteLine(txtUsername.Text + "," + txtUserPassword.Text + "," + (cmbUserAccess.SelectedIndex+1).ToString());
                writer.Close();
                LoadUsers();

                txtUsername.Clear();
                txtUserPassword.Clear();
                cmbUserAccess.SelectedIndex = -1;

                lblSuccess4.Text = "New User successful";
                Util.Animate(lblSuccess4, Util.Effect.Slide, 90, 90);
                timer1.Enabled = true;
                timer1.Start();

                return;
            }

            if (radEditUser.Checked == true)
            {
                for(int i=0;i<Users.Count;i++)
                { 
                    if (Users[i] == txtUsername.Text)
                    {
                        Users[i] = txtUsername.Text;
                        PW[i] = txtUserPassword.Text;
                        Access[i] = (cmbUserAccess.SelectedIndex + 1).ToString();
                        break;
                    }
                }
               
                var writer = new StreamWriter(@"T:\Databases\pdata.dbs",false);

                for (int i = 0; i < Users.Count; i++)
                {
                    writer.WriteLine(Users[i] + "," + PW[i] + "," + Access[i]);
                }
                writer.Close();
                LoadUsers();

                txtUsername.Clear();
                txtUserPassword.Clear();
                cmbUserAccess.SelectedIndex = -1;

                lblSuccess4.Text = "Edit User successful";
                Util.Animate(lblSuccess4, Util.Effect.Slide, 90, 90);
                timer1.Enabled = true;
                timer1.Start();

                return;
            }

            if (radDeleteUser.Checked == true)
            {
                var writer = new StreamWriter(@"T:\Databases\pdata.dbs",false);

                for (int i = 0; i < Users.Count; i++)
                {
                    if (Users[i] != txtUsername.Text)
                    {
                        writer.WriteLine(Users[i] + "," + PW[i] + "," + Access[i]);
                    }
                }
                writer.Close();
                LoadUsers();

                txtUsername.Clear();
                txtUserPassword.Clear();
                cmbUserAccess.SelectedIndex = -1;

                lblSuccess4.Text = "Delete User successful";
                Util.Animate(lblSuccess4, Util.Effect.Slide, 90, 90);
                timer1.Enabled = true;
                timer1.Start();

                return;
            }
        }

        private void radEditUser_CheckedChanged(object sender, EventArgs e)
        {
            if(radEditUser.Checked== true)
            {
                listUsers.Enabled = true;
                grpUserForm.Enabled = true;
                txtUsername.Enabled = false;

                Util.Animate(lblUserOp, Util.Effect.Slide, 80, 90);
                lblUserOp.Text = "Edit User";
                Util.Animate(lblUserOp, Util.Effect.Slide, 80, 90);
            }
        }

        private void radDeleteUser_CheckedChanged(object sender, EventArgs e)
        {
            if (radDeleteUser.Checked == true)
            {
                listUsers.Enabled = true;
                grpUserForm.Enabled = true;
                txtUsername.Enabled = false;

                Util.Animate(lblUserOp, Util.Effect.Slide, 80, 90);
                lblUserOp.Text = "Delete User";
                Util.Animate(lblUserOp, Util.Effect.Slide, 80, 90);
            }
        }

        private void cmbUserAccess_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listUsers_Click(object sender, EventArgs e)
        {
           
        }

        private void txtUsername_Leave(object sender, EventArgs e)
        {
            if (radNewUser.Checked == true)
            {
                for (int i = 0; i < Users.Count; i++)
                {
                    if (Users[i] == txtUsername.Text)
                    {
                        MessageBox.Show("Username already exists! Switching to Edit mode");

                        radEditUser.Checked = true;
                        listUsers.Enabled = true;
                        grpUserForm.Enabled = true;
                        txtUsername.Enabled = false;
                        listUsers.SelectedIndex = i;

                        Util.Animate(lblUserOp, Util.Effect.Slide, 80, 90);
                        lblUserOp.Text = "Edit User";
                        Util.Animate(lblUserOp, Util.Effect.Slide, 80, 90);
                        return;
                    }
                }
            }
        }

        private void txtUsername_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
        }

        private void txtUserPassword_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtUserPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
        }

        private void btnChangePass_Click(object sender, EventArgs e)
        {
            btnChangePass.Visible = false;
            Util.Animate(grpChangePass, Util.Effect.Slide, 150, 180);
            grpChangePass.Visible = true;

            txtChangePass.Focus();           
        }

        private void btnChangeX_Click(object sender, EventArgs e)
        {
            btnChangePass.Visible = true;
            Util.Animate(grpChangePass, Util.Effect.Slide, 150, 180);
            grpChangePass.Visible = false;
            txtChangePass.Text = "";
        }

        private void ChangePassword()
        {
            if (txtChangePass.Text.Contains(','))
            {
                MessageBox.Show("Invalid Password!");
                return;
            }

            if (txtChangePass.Text != "")
            {
                for (int i = 0; i < Users.Count; i++)
                {
                    if (Users[i] == currentuser)
                    {
                        PW[i] = txtChangePass.Text;

                        var writer = new StreamWriter(@"T:\Databases\pdata.dbs", false);

                        for (int j = 0; j < Users.Count; j++)
                        {
                            writer.WriteLine(Users[j] + "," + PW[j] + "," + Access[j]);
                        }
                        writer.Close();

                        var autowriter = new StreamWriter(@"C:\Users\Public\loginfile");
                        autowriter.WriteLine(currentuser);
                        autowriter.WriteLine(txtChangePass.Text);
                        autowriter.Close();

                        LoadUsers();

                        txtChangePass.Clear();
                        //btnChangePass.Visible = true;
                        Util.Animate(grpChangePass, Util.Effect.Slide, 150, 180);
                        grpChangePass.Visible = false;

                        Util.Animate(lblSuccess5, Util.Effect.Slide, 150, 90);
                        lblSuccess5.Visible = true;
                        timer1.Enabled = true;
                        timer1.Start();

                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("Password can not be blank!");
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            ChangePassword();
        }

        private void txtChangePass_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ChangePassword();
            }
        }

        private void ClearChart()
        {
            foreach (var series in chartData.Series)
            {
                series.Points.Clear();
            }
            chartData.Series.Clear();
            chartData.Titles.Clear();
            chartData.Legends.Clear();
        }

        private void EnableStatusChart()
        {
            #region
            int aedplus_notready = 0;
            int aedplus_ready = 0;
            int aedplus_repair = 0;
            int aedplus_out30 = 0;
            int aedplus_overdue = 0;
            int aedplus_perma = 0;

            int aedpro_notready = 0;
            int aedpro_ready = 0;
            int aedpro_repair = 0;
            int aedpro_out30 = 0;
            int aedpro_overdue = 0;
            int aedpro_perma = 0;

            int autopulse_notready = 0;
            int autopulse_ready = 0;
            int autopulse_repair = 0;
            int autopulse_out30 = 0;
            int autopulse_overdue = 0;
            int autopulse_perma = 0;

            int cct_notready = 0;
            int cct_ready = 0;
            int cct_repair = 0;
            int cct_out30 = 0;
            int cct_overdue = 0;
            int cct_perma = 0;

            int eseries_notready = 0;
            int eseries_ready = 0;
            int eseries_repair = 0;
            int eseries_out30 = 0;
            int eseries_overdue = 0;
            int eseries_perma = 0;

            int mseries_notready = 0;
            int mseries_ready = 0;
            int mseries_repair = 0;
            int mseries_out30 = 0;
            int mseries_overdue = 0;
            int mseries_perma = 0;

            int propaq_notready = 0;
            int propaq_ready = 0;
            int propaq_repair = 0;
            int propaq_out30 = 0;
            int propaq_overdue = 0;
            int propaq_perma = 0;

            int rseries_notready = 0;
            int rseries_ready = 0;
            int rseries_repair = 0;
            int rseries_out30 = 0;
            int rseries_overdue = 0;
            int rseries_perma = 0;

            int xseries_notready = 0;
            int xseries_ready = 0;
            int xseries_repair = 0;
            int xseries_out30 = 0;
            int xseries_overdue = 0;
            int xseries_perma = 0;
            #endregion

            #region Collect data from MasterList
            for (int i = 0; i < MasterList[0].Count; i++)
            {
                if (MasterList[2][i] == "IN")
                {
                    if (MasterList[11][i] == "N")
                    {
                        if (MasterList[12][i] == "Y")
                        {
                                #region
                                if (MasterList[9][i] == "AED-PLUS")
                                {
                                    aedplus_repair++;
                                }
                                if (MasterList[9][i] == "AED-PRO")
                                {
                                    aedpro_repair++;
                                }
                                if (MasterList[9][i] == "AUTOPULSE")
                                {
                                    autopulse_repair++;
                                }
                                if (MasterList[9][i] == "CCT")
                                {
                                    cct_repair++;
                                }
                                if (MasterList[9][i] == "E-SERIES")
                                {
                                    eseries_repair++;
                                }
                                if (MasterList[9][i] == "M-SERIES")
                                {
                                    mseries_repair++;
                                }
                                if (MasterList[9][i] == "PROPAQ")
                                {
                                    propaq_repair++;
                                }
                                if (MasterList[9][i] == "R-SERIES")
                                {
                                    rseries_repair++;
                                }
                                if (MasterList[9][i] == "X-SERIES")
                                {
                                    xseries_repair++;
                                }
                                #endregion
                        }
                        else
                        {
                            #region
                            if (MasterList[9][i] == "AED-PLUS")
                            {
                                aedplus_notready++;
                            }
                            if (MasterList[9][i] == "AED-PRO")
                            {
                                aedpro_notready++;
                            }
                            if (MasterList[9][i] == "AUTOPULSE")
                            {
                                autopulse_notready++;
                            }
                            if (MasterList[9][i] == "CCT")
                            {
                                cct_notready++;
                            }
                            if (MasterList[9][i] == "E-SERIES")
                            {
                                eseries_notready++;
                            }
                            if (MasterList[9][i] == "M-SERIES")
                            {
                                mseries_notready++;
                            }
                            if (MasterList[9][i] == "PROPAQ")
                            {
                                propaq_notready++;
                            }
                            if (MasterList[9][i] == "R-SERIES")
                            {
                                rseries_notready++;
                            }
                            if (MasterList[9][i] == "X-SERIES")
                            {
                                xseries_notready++;
                            }
                            #endregion
                        }
                    }
                    if (MasterList[11][i] == "Y")
                    {
                        #region
                        if (MasterList[9][i] == "AED-PLUS")
                        {
                            aedplus_ready++;
                        }
                        if (MasterList[9][i] == "AED-PRO")
                        {
                            aedpro_ready++;
                        }
                        if (MasterList[9][i] == "AUTOPULSE")
                        {
                            autopulse_ready++;
                        }
                        if (MasterList[9][i] == "CCT")
                        {
                            cct_ready++;
                        }
                        if (MasterList[9][i] == "E-SERIES")
                        {
                            eseries_ready++;
                        }
                        if (MasterList[9][i] == "M-SERIES")
                        {
                            mseries_ready++;
                        }
                        if (MasterList[9][i] == "PROPAQ")
                        {
                            propaq_ready++;
                        }
                        if (MasterList[9][i] == "R-SERIES")
                        {
                            rseries_ready++;
                        }
                        if (MasterList[9][i] == "X-SERIES")
                        {
                            xseries_ready++;
                        }
                        #endregion
                    }
                }

                if (MasterList[2][i] == "OUT")
                {
                    DateTime parsedDate;
                    string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };
                    DateTime.TryParseExact(MasterList[6][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                    if (DateTime.Today.AddMonths(-1).CompareTo(parsedDate) > 0)
                    {
                        #region
                        if (MasterList[9][i] == "AED-PLUS")
                        {
                            aedplus_overdue++;
                        }
                        if (MasterList[9][i] == "AED-PRO")
                        {
                            aedpro_overdue++;
                        }
                        if (MasterList[9][i] == "AUTOPULSE")
                        {
                            autopulse_overdue++;
                        }
                        if (MasterList[9][i] == "CCT")
                        {
                            cct_overdue++;
                        }
                        if (MasterList[9][i] == "E-SERIES")
                        {
                            eseries_overdue++;
                        }
                        if (MasterList[9][i] == "M-SERIES")
                        {
                            mseries_overdue++;
                        }
                        if (MasterList[9][i] == "PROPAQ")
                        {
                            propaq_overdue++;
                        }
                        if (MasterList[9][i] == "R-SERIES")
                        {
                            rseries_overdue++;
                        }
                        if (MasterList[9][i] == "X-SERIES")
                        {
                            xseries_overdue++;
                        }
                        #endregion
                    }
                    else
                    {
                        #region
                        if (MasterList[9][i] == "AED-PLUS")
                        {
                            aedplus_out30++;
                        }
                        if (MasterList[9][i] == "AED-PRO")
                        {
                            aedpro_out30++;
                        }
                        if (MasterList[9][i] == "AUTOPULSE")
                        {
                            autopulse_out30++;
                        }
                        if (MasterList[9][i] == "CCT")
                        {
                            cct_out30++;
                        }
                        if (MasterList[9][i] == "E-SERIES")
                        {
                            eseries_out30++;
                        }
                        if (MasterList[9][i] == "M-SERIES")
                        {
                            mseries_out30++;
                        }
                        if (MasterList[9][i] == "PROPAQ")
                        {
                            propaq_out30++;
                        }
                        if (MasterList[9][i] == "R-SERIES")
                        {
                            rseries_out30++;
                        }
                        if (MasterList[9][i] == "X-SERIES")
                        {
                            xseries_out30++;
                        }
                        #endregion
                    }

                }

                if (MasterList[2][i] == "PERMA")
                {
                    #region
                    if (MasterList[9][i] == "AED-PLUS")
                    {
                        aedplus_perma++;
                    }
                    if (MasterList[9][i] == "AED-PRO")
                    {
                        aedpro_perma++;
                    }
                    if (MasterList[9][i] == "AUTOPULSE")
                    {
                        autopulse_perma++;
                    }
                    if (MasterList[9][i] == "CCT")
                    {
                        cct_perma++;
                    }
                    if (MasterList[9][i] == "E-SERIES")
                    {
                        eseries_perma++;
                    }
                    if (MasterList[9][i] == "M-SERIES")
                    {
                        mseries_perma++;
                    }
                    if (MasterList[9][i] == "PROPAQ")
                    {
                        propaq_perma++;
                    }
                    if (MasterList[9][i] == "R-SERIES")
                    {
                        rseries_perma++;
                    }
                    if (MasterList[9][i] == "X-SERIES")
                    {
                        xseries_perma++;
                    }
                    #endregion
                }
            }
            #endregion

            ClearChart();
            chartData.Titles.Add("Loaner Status Breakdown by Model");
            chartData.ChartAreas[0].Area3DStyle.Enable3D = false;
            chartData.Legends.Add("A");

            if (cmbBreakdownOptions.SelectedIndex == 0 || cmbBreakdownOptions.SelectedIndex == 2)
            {
                chartData.Series.Add("Not Ready");
                chartData.Series.Add("Ready to Ship");
                chartData.Series.Add("Repair Required");

                if (chkStack100.Checked == false)
                {
                    chartData.Series["Not Ready"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar;
                    chartData.Series["Ready to Ship"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar;
                    chartData.Series["Repair Required"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar;
                }
                else
                {
                    chartData.Series["Not Ready"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar100;
                    chartData.Series["Ready to Ship"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar100;
                    chartData.Series["Repair Required"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar100;
                }

                chartData.Series["Not Ready"].IsVisibleInLegend = true;
                chartData.Series["Ready to Ship"].IsVisibleInLegend = true;
                chartData.Series["Repair Required"].IsVisibleInLegend = true;

                chartData.Series["Not Ready"]["BarLabelStyle"] = "Center";
                chartData.Series["Ready to Ship"]["BarLabelStyle"] = "Center";
                chartData.Series["Repair Required"]["BarLabelStyle"] = "Center";
            }

            if (cmbBreakdownOptions.SelectedIndex == 0 || cmbBreakdownOptions.SelectedIndex == 1)
            {
                chartData.Series.Add("Out within 30 Days");
                chartData.Series.Add("Overdue");
                chartData.Series.Add("Permanent");

                if (chkStack100.Checked == false)
                {
                    chartData.Series["Out within 30 Days"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar;
                    chartData.Series["Overdue"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar;
                    chartData.Series["Permanent"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar;
                }
                else
                {
                    chartData.Series["Out within 30 Days"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar100;
                    chartData.Series["Overdue"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar100;
                    chartData.Series["Permanent"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar100;
                }

                chartData.Series["Out within 30 Days"].IsVisibleInLegend = true;
                chartData.Series["Overdue"].IsVisibleInLegend = true;
                chartData.Series["Permanent"].IsVisibleInLegend = true;

                chartData.Series["Out within 30 Days"]["BarLabelStyle"] = "Center";
                chartData.Series["Overdue"]["BarLabelStyle"] = "Center";
                chartData.Series["Permanent"]["BarLabelStyle"] = "Center";
            }

            if (cmbBreakdownOptions.SelectedIndex == 3)
            {
                chartData.Series.Add("In");
                chartData.Series.Add("Out");

                if (chkStack100.Checked == false)
                {
                    chartData.Series["In"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar;
                    chartData.Series["Out"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar;
                }
                else
                {
                    chartData.Series["In"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar100;
                    chartData.Series["Out"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar100;
                }

                chartData.Series["In"].IsVisibleInLegend = true;
                chartData.Series["Out"].IsVisibleInLegend = true;

                chartData.Series["In"]["BarLabelStyle"] = "Center";
                chartData.Series["Out"]["BarLabelStyle"] = "Center";
            }
            #region #Series adding region
            if (cmbBreakdownOptions.SelectedIndex == 0 || cmbBreakdownOptions.SelectedIndex == 2)
            {
                chartData.Series["Not Ready"].Points.AddXY("X-SERIES", xseries_notready);
                chartData.Series["Not Ready"].Points.AddXY("R-SERIES", rseries_notready);
                chartData.Series["Not Ready"].Points.AddXY("PROPAQ", propaq_notready);
                chartData.Series["Not Ready"].Points.AddXY("M-SERIES", mseries_notready);
                chartData.Series["Not Ready"].Points.AddXY("E-SERIES", eseries_notready);
                chartData.Series["Not Ready"].Points.AddXY("CCT", cct_notready);
                chartData.Series["Not Ready"].Points.AddXY("AUTOPULSE", autopulse_notready);
                chartData.Series["Not Ready"].Points.AddXY("AED-PRO", aedpro_notready);
                chartData.Series["Not Ready"].Points.AddXY("AED-PLUS", aedplus_notready);

                chartData.Series["Ready to Ship"].Points.AddXY("X-SERIES", xseries_ready);
                chartData.Series["Ready to Ship"].Points.AddXY("R-SERIES", rseries_ready);
                chartData.Series["Ready to Ship"].Points.AddXY("PROPAQ", propaq_ready);
                chartData.Series["Ready to Ship"].Points.AddXY("M-SERIES", mseries_ready);
                chartData.Series["Ready to Ship"].Points.AddXY("E-SERIES", eseries_ready);
                chartData.Series["Ready to Ship"].Points.AddXY("CCT", cct_ready);
                chartData.Series["Ready to Ship"].Points.AddXY("AUTOPULSE", autopulse_ready);
                chartData.Series["Ready to Ship"].Points.AddXY("AED-PRO", aedpro_ready);
                chartData.Series["Ready to Ship"].Points.AddXY("AED-PLUS", aedplus_ready);

                chartData.Series["Repair Required"].Points.AddXY("X-SERIES", xseries_repair);
                chartData.Series["Repair Required"].Points.AddXY("R-SERIES", rseries_repair);
                chartData.Series["Repair Required"].Points.AddXY("PROPAQ", propaq_repair);
                chartData.Series["Repair Required"].Points.AddXY("M-SERIES", mseries_repair);
                chartData.Series["Repair Required"].Points.AddXY("E-SERIES", eseries_repair);
                chartData.Series["Repair Required"].Points.AddXY("CCT", cct_repair);
                chartData.Series["Repair Required"].Points.AddXY("AUTOPULSE", autopulse_repair);
                chartData.Series["Repair Required"].Points.AddXY("AED-PRO", aedpro_repair);
                chartData.Series["Repair Required"].Points.AddXY("AED-PLUS", aedplus_repair);
            }


            if (cmbBreakdownOptions.SelectedIndex == 0 || cmbBreakdownOptions.SelectedIndex == 1)
            {
                chartData.Series["Out within 30 Days"].Points.AddXY("X-SERIES", xseries_out30);
                chartData.Series["Out within 30 Days"].Points.AddXY("R-SERIES", rseries_out30);
                chartData.Series["Out within 30 Days"].Points.AddXY("PROPAQ", propaq_out30);
                chartData.Series["Out within 30 Days"].Points.AddXY("M-SERIES", mseries_out30);
                chartData.Series["Out within 30 Days"].Points.AddXY("E-SERIES", eseries_out30);
                chartData.Series["Out within 30 Days"].Points.AddXY("CCT", cct_out30);
                chartData.Series["Out within 30 Days"].Points.AddXY("AUTOPULSE", autopulse_out30);
                chartData.Series["Out within 30 Days"].Points.AddXY("AED-PRO", aedpro_out30);
                chartData.Series["Out within 30 Days"].Points.AddXY("AED-PLUS", aedplus_out30);

                chartData.Series["Overdue"].Points.AddXY("X-SERIES", xseries_overdue);
                chartData.Series["Overdue"].Points.AddXY("R-SERIES", rseries_overdue);
                chartData.Series["Overdue"].Points.AddXY("PROPAQ", propaq_overdue);
                chartData.Series["Overdue"].Points.AddXY("M-SERIES", mseries_overdue);
                chartData.Series["Overdue"].Points.AddXY("E-SERIES", eseries_overdue);
                chartData.Series["Overdue"].Points.AddXY("CCT", cct_overdue);
                chartData.Series["Overdue"].Points.AddXY("AUTOPULSE", autopulse_overdue);
                chartData.Series["Overdue"].Points.AddXY("AED-PRO", aedpro_overdue);
                chartData.Series["Overdue"].Points.AddXY("AED-PLUS", aedplus_overdue);

                chartData.Series["Permanent"].Points.AddXY("X-SERIES", xseries_perma);
                chartData.Series["Permanent"].Points.AddXY("R-SERIES", rseries_perma);
                chartData.Series["Permanent"].Points.AddXY("PROPAQ", propaq_perma);
                chartData.Series["Permanent"].Points.AddXY("M-SERIES", mseries_perma);
                chartData.Series["Permanent"].Points.AddXY("E-SERIES", eseries_perma);
                chartData.Series["Permanent"].Points.AddXY("CCT", cct_perma);
                chartData.Series["Permanent"].Points.AddXY("AUTOPULSE", autopulse_perma);
                chartData.Series["Permanent"].Points.AddXY("AED-PRO", aedpro_perma);
                chartData.Series["Permanent"].Points.AddXY("AED-PLUS", aedplus_perma);
            }

            if (cmbBreakdownOptions.SelectedIndex == 3)
            {
                chartData.Series["In"].Points.AddXY("X-SERIES", xseries_notready + xseries_ready + xseries_repair);
                chartData.Series["In"].Points.AddXY("R-SERIES", rseries_notready + rseries_ready + rseries_repair);
                chartData.Series["In"].Points.AddXY("PROPAQ", propaq_notready + propaq_ready + propaq_repair);
                chartData.Series["In"].Points.AddXY("M-SERIES", mseries_notready + mseries_ready + mseries_repair);
                chartData.Series["In"].Points.AddXY("E-SERIES", eseries_notready + eseries_ready + eseries_repair);
                chartData.Series["In"].Points.AddXY("CCT", cct_notready + cct_ready + cct_repair);
                chartData.Series["In"].Points.AddXY("AUTOPULSE", autopulse_notready + autopulse_ready + autopulse_repair);
                chartData.Series["In"].Points.AddXY("AED-PRO", aedpro_notready + aedpro_ready + aedpro_repair);
                chartData.Series["In"].Points.AddXY("AED-PLUS", aedplus_notready + aedplus_ready + aedplus_repair);

                chartData.Series["Out"].Points.AddXY("X-SERIES", xseries_out30 + xseries_overdue + xseries_perma);
                chartData.Series["Out"].Points.AddXY("R-SERIES", rseries_out30 + rseries_overdue + rseries_perma);
                chartData.Series["Out"].Points.AddXY("PROPAQ", propaq_out30 + propaq_overdue + propaq_perma);
                chartData.Series["Out"].Points.AddXY("M-SERIES", mseries_out30 + mseries_overdue + mseries_perma);
                chartData.Series["Out"].Points.AddXY("E-SERIES", eseries_out30 + eseries_overdue + eseries_perma);
                chartData.Series["Out"].Points.AddXY("CCT", cct_out30 + cct_overdue+ cct_perma);
                chartData.Series["Out"].Points.AddXY("AUTOPULSE", autopulse_out30 + autopulse_overdue + autopulse_perma);
                chartData.Series["Out"].Points.AddXY("AED-PRO", aedpro_out30 + aedpro_overdue + aedpro_perma);
                chartData.Series["Out"].Points.AddXY("AED-PLUS", aedplus_out30 + aedplus_overdue + aedplus_perma);
            }
            #endregion

            
            

            foreach (System.Windows.Forms.DataVisualization.Charting.Series s in chartData.Series)
            {
                foreach (System.Windows.Forms.DataVisualization.Charting.DataPoint dp in s.Points)
                {
                    if (dp.YValues[0].ToString() == "0")
                    {
                        dp.IsValueShownAsLabel = false;
                    }
                    else
                    {
                        dp.IsValueShownAsLabel = true;
                        dp.Label = "#VALY";
                    }
                }
            }

            /*
            chartData.Series["Not Ready"].Label = "#VALY";
            chartData.Series["Ready to Ship"].Label = "#VALY";
            chartData.Series["Repair Required"].Label = "#VALY";
            chartData.Series["Out within 30 Days"].Label = "#VALY";
            chartData.Series["Overdue"].Label = "#VALY";
            chartData.Series["Permanent"].Label = "#VALY";
            
            For Each s As Series In Chart1.Series
        For Each dp As DataPoint In s.Points
            If dp.YValues(0) = 0 Then

                dp.IsValueShownAsLabel = False
                */
            

            /*
            chartData.Series["Not Ready"].Label = "#SERIESNAME\n#VALY units";
            chartData.Series["Ready to Ship"].Label = "#SERIESNAME\n#VALY units";
            chartData.Series["Repair Required"].Label = "#SERIESNAME\n#VALY units";
            chartData.Series["Out within 30 Days"].Label = "#SERIESNAME\n#VALY units";
            chartData.Series["Overdue"].Label = "#SERIESNAME\n#VALY units";
            chartData.Series["Permanent"].Label = "#SERIESNAME\n#VALY units";
            */
            
            
            
            
        }

        private void EnableLoanerBreakdownChart()
        {
            ClearChart();

            chartData.Series.Add("Loaner Status");

            if (cmbBreakdownOptions.SelectedIndex == 0)
            {
                chartData.Titles.Add("Loaner Status Breakdown - All Loaners");
            }

            if (cmbBreakdownOptions.SelectedIndex == 1)
            {
                chartData.Titles.Add("Loaner Status Breakdown - Out Loaners");
            }

            if (cmbBreakdownOptions.SelectedIndex == 2)
            {
                chartData.Titles.Add("Loaner Status Breakdown - In-House Loaners");
            }

            if (cmbBreakdownOptions.SelectedIndex == 3)
            {
                chartData.Titles.Add("Loaner Status Breakdown - In vs Out");
            }

            //chartData.Series["Loaner Status"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            chartData.Series["Loaner Status"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;

            if (cmbBreakdownOptions.SelectedIndex == 0 || cmbBreakdownOptions.SelectedIndex == 2)
            {
                chartData.Series["Loaner Status"].Points.AddXY("Not Prepared", Int32.Parse(btnNotReady.Text.Split(' ')[0]));
            }

            if (cmbBreakdownOptions.SelectedIndex == 0 || cmbBreakdownOptions.SelectedIndex == 2)
            {
                chartData.Series["Loaner Status"].Points.AddXY("Ready to Ship", Int32.Parse(btnReady.Text.Split(' ')[0]));
            }
            
            if (cmbBreakdownOptions.SelectedIndex == 0 || cmbBreakdownOptions.SelectedIndex == 2)
            {
                chartData.Series["Loaner Status"].Points.AddXY("Require Repairs", Int32.Parse(btnRepair.Text.Split(' ')[0]));
            }

            if (cmbBreakdownOptions.SelectedIndex == 0 || cmbBreakdownOptions.SelectedIndex == 1)
            {
                chartData.Series["Loaner Status"].Points.AddXY("Out within 30 Days", Int32.Parse(btnSent.Text.Split(' ')[0]));
            }

            if (cmbBreakdownOptions.SelectedIndex == 0 || cmbBreakdownOptions.SelectedIndex == 1)
            {
                chartData.Series["Loaner Status"].Points.AddXY("Overdue", Int32.Parse(btnOverdue.Text.Split(' ')[0]));
            }

            if (cmbBreakdownOptions.SelectedIndex == 0 || cmbBreakdownOptions.SelectedIndex == 1)
            {
                chartData.Series["Loaner Status"].Points.AddXY("Permanent/Long Term", permanent);
            }

            if (cmbBreakdownOptions.SelectedIndex == 3)
            {
                chartData.Series["Loaner Status"].Points.AddXY("In", Int32.Parse(btnInHouse.Text.Split(' ')[0]));
                chartData.Series["Loaner Status"].Points.AddXY("Out", Int32.Parse(btnOut.Text.Split(' ')[0]));
            }

            chartData.Series["Loaner Status"].IsVisibleInLegend = false;
            chartData.Series["Loaner Status"].Label = "#VALX\n#VALY units\n#PERCENT";
            chartData.Series["Loaner Status"]["PieLabelStyle"] = "Outside";


            chartData.ChartAreas[0].Area3DStyle.Enable3D = true;
        }

        private DateTime GetDateofSundayoftheWeek(DateTime day)
        {
            while (true)
            {
                if (day.DayOfWeek == DayOfWeek.Sunday)
                {
                    return day;
                }
                else
                {
                    day = day.AddDays(-1);
                }
            }
        }

        private void EnableTrendChart()
        {
            try
            {
                var reader = new StreamReader("record.dbs");
                string dummy;
                bool first = true;
                DateTime parsedDate;
                DateTime thissunday;
                string[] formatstring = { "d/M/yyyy","dd/M/yyyy","dd/MM/yyyy", "d/MM/yyyy" };

                int[] sent = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 ,0};
                int[] received = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0 };


                while (!reader.EndOfStream)
                {
                    if (first == true)
                    {
                        dummy = reader.ReadLine();
                        first = false;
                    }
                    else
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');
                        DateTime.TryParseExact(values[0], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);
                        thissunday = GetDateofSundayoftheWeek(DateTime.Today);

                        for (int i = 0; i < sent.Length-1; i++)
                        {
                            if (thissunday.AddDays(i * (-7)).CompareTo(parsedDate) > 0 && thissunday.AddDays((i + 1) * (-7)).CompareTo(parsedDate) < 0)
                            {
                                if (values[10] == "E-SERIES" && chkChartESeries.Checked == true || (values[10] == "M-SERIES" || values[10] == "CCT") && chkChartMSeries.Checked == true || values[10] == "R-SERIES" && chkChartRSeries.Checked == true || values[10] == "X-SERIES" && chkChartXSeries.Checked == true || values[10] == "PROPAQ" && chkChartPropaq.Checked == true || values[10] == "AED-PLUS" && chkChartAEDPlus.Checked == true || values[10] == "AED-PRO" && chkChartAEDPro.Checked == true || values[10] == "AUTOPULSE" && chkChartAutopulse.Checked == true)
                                {
                                    if (values[1] == "SEND")
                                    {
                                        sent[i] = sent[i] + 1;
                                    }

                                    if (values[1] == "RECEIVE")
                                    {
                                        received[i] = received[i] + 1;
                                    }
                                }
                            }
                        }

                        if (thissunday.CompareTo(parsedDate) < 0)
                        {
                            if (values[10] == "E-SERIES" && chkChartESeries.Checked == true || (values[10] == "M-SERIES" || values[10] == "CCT") && chkChartMSeries.Checked == true || values[10] == "R-SERIES" && chkChartRSeries.Checked == true || values[10] == "X-SERIES" && chkChartXSeries.Checked == true || values[10] == "PROPAQ" && chkChartPropaq.Checked == true || values[10] == "AED-PLUS" && chkChartAEDPlus.Checked == true || values[10] == "AED-PRO" && chkChartAEDPro.Checked == true || values[10] == "AUTOPULSE" && chkChartAutopulse.Checked == true)
                            {
                                if (values[1] == "SEND")
                                {
                                    sent[10] = sent[10] + 1;
                                }

                                if (values[1] == "RECEIVE")
                                {
                                    received[10] = received[10] + 1;
                                }
                            }
                        }


                    }
                }
                reader.Close();

                //MessageBox.Show(sent[10].ToString());

                ClearChart();
                
                chartData.Titles.Add("Loaner Trends in the Past 10 Weeks");
                chartData.ChartAreas[0].Area3DStyle.Enable3D = false;
                chartData.Series.Add("Loaners Sent per Week");
                chartData.Series.Add("Loaners Received per Week");
                chartData.Legends.Add("A");
                chartData.Series["Loaners Sent per Week"].IsVisibleInLegend = false;
                chartData.Series["Loaners Received per Week"].IsVisibleInLegend = false;
                chartData.ChartAreas[0].AxisX.LabelStyle.Interval = 1;
                chartData.ChartAreas[0].AxisX.MajorGrid.Interval = 1;
                chartData.ChartAreas[0].AxisX.MajorTickMark.Enabled = true;


                thissunday = GetDateofSundayoftheWeek(DateTime.Today);

                if (chkChartSent.Checked == true)
                {

                    chartData.Series["Loaners Sent per Week"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                    chartData.Series["Loaners Sent per Week"].BorderWidth = 2;
                    chartData.Series["Loaners Sent per Week"].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Circle;
                    chartData.Series["Loaners Sent per Week"].IsVisibleInLegend = true;

                    if (chkReportDates.Checked == true)
                    {
                        for (int i = 8; i >= 0; i--)
                        {

                            chartData.Series["Loaners Sent per Week"].Points.AddXY((i + 1).ToString()+ " weeks ago", sent[i]);
                        }
                        chartData.Series["Loaners Sent per Week"].Points.AddXY("This week", sent[10]);
                    }
                    else
                    {
                        for (int i = 8; i >= 0; i--)
                        {

                            chartData.Series["Loaners Sent per Week"].Points.AddXY(thissunday.AddDays((-i - 1) * 7).ToString("MMM/dd"), sent[i]);
                        }
                        chartData.Series["Loaners Sent per Week"].Points.AddXY(thissunday.ToString("MMM/dd"), sent[10]);
                    }
                }

                if (chkChartReceived.Checked == true)
                {

                    chartData.Series["Loaners Received per Week"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                    chartData.Series["Loaners Received per Week"].BorderWidth = 2;
                    chartData.Series["Loaners Received per Week"].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Circle;
                    chartData.Series["Loaners Received per Week"].IsVisibleInLegend = true;

                    if (chkReportDates.Checked == true)
                    {
                        for (int i = 8; i >= 0; i--)
                        {
                            chartData.Series["Loaners Received per Week"].Points.AddXY((i + 1).ToString() + " weeks ago", received[i]);
                        }
                        chartData.Series["Loaners Received per Week"].Points.AddXY("This week", received[10]);
                    }
                    else{
                        for (int i = 8; i >= 0; i--)
                        {
                            chartData.Series["Loaners Received per Week"].Points.AddXY(thissunday.AddDays((-i - 1) * 7).ToString("MMM/dd"), received[i]);
                        }
                        chartData.Series["Loaners Received per Week"].Points.AddXY(thissunday.ToString("MMM/dd"), received[10]);
                    }
                }

            }
            catch
            {
                MessageBox.Show("Archives currently in use. Please try again.");
                return;
            }
                       
        }

        private void btnSettingQuantity_Click(object sender, EventArgs e)
        {
           
        }

        private void chkChartSent_CheckedChanged(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void chkChartReceived_CheckedChanged(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void chkChartAll_CheckedChanged(object sender, EventArgs e)
        {
            if (chkChartAll.Checked == false)
            {
                grpChartFilter.Enabled = true;
                chkChartESeries.Checked = false;
                chkChartMSeries.Checked = false;
                chkChartRSeries.Checked = false;
                chkChartXSeries.Checked = false;
                chkChartPropaq.Checked = false;
                chkChartAEDPlus.Checked = false;
                chkChartAEDPro.Checked = false;
                chkChartAutopulse.Checked = false;
            }
            else
            {
                grpChartFilter.Enabled = false;
                chkChartESeries.Checked = true;
                chkChartMSeries.Checked = true;
                chkChartRSeries.Checked = true;
                chkChartXSeries.Checked = true;
                chkChartPropaq.Checked = true;
                chkChartAEDPlus.Checked = true;
                chkChartAEDPro.Checked = true;
                chkChartAutopulse.Checked = true;

            }
            EnableTrendChart();
        }

        private void chkChartESeries_Click(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void chkChartMSeries_Click(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void chkChartRSeries_Click(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void chkChartXSeries_Click(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void chkChartPropaq_Click(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void chkChartAEDPlus_Click(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void chkChartAEDPro_Click(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void chkChartAutopulse_Click(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void cmbChartView_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbChartView.SelectedIndex == 0)
            {
                //EnableLoanerBreakdownChart();
                cmbBreakdownOptions.SelectedIndex = -1;
                cmbBreakdownOptions.SelectedIndex = 0;
                grpChartBreakdown.Visible = true;
                chkStack100.Visible = false;
                grpChartTrend.Visible = false;
            }

            if (cmbChartView.SelectedIndex == 1)
            {
                EnableTrendChart();
                grpChartBreakdown.Visible = false;
                grpChartTrend.Visible = true;
            }

            if (cmbChartView.SelectedIndex == 2)
            {
                EnableStatusChart();
                cmbBreakdownOptions.SelectedIndex = -1;
                cmbBreakdownOptions.SelectedIndex = 0;
                grpChartBreakdown.Visible = true;
                chkStack100.Visible = true;
                grpChartTrend.Visible = false;
            }
        }

        private void cmbBreakdownOptions_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbChartView.SelectedIndex == 0)
            {
                EnableLoanerBreakdownChart();
            }
            if (cmbChartView.SelectedIndex == 2)
            {
                EnableStatusChart();
            }
            //grpChartTrend.Visible = false;
        }

        private void ClearConfig()
        {
            lblOpt1.Text = "...";
            lblOpt2.Text = "...";
            lblOpt3.Text = "...";
            lblOpt4.Text = "...";
            lblOpt5.Text = "...";
            lblOpt6.Text = "...";
            lblOpt7.Text = "...";
            lblOpt8.Text = "...";
            lblOpt9.Text = "...";
            lblOpt10.Text = "...";
            lblOpt11.Text = "...";
            lblOpt12.Text = "...";
            lblOpt13.Text = "...";
            lblOpt14.Text = "...";
            lblOpt15.Text = "...";
        }

        private void FindConfig(string config)
        {
            ClearConfig();

            //X-SERIES and PROPAQ Find
            if (config.Contains('-') == true)
            {
                //Propaq Find
                if (config[0] == '2'|| config[0] == '3')
                {
                    var reader = new StreamReader("propaq.dbs");

                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');

                        if (config[0].ToString() == values[0])
                        {
                            lblOpt1.Text = values[1];
                        }

                        if (config[1].ToString() + config[2].ToString() == values[2])
                        {
                            lblOpt2.Text = values[3];
                        }

                        if (config[3].ToString() == "-")
                        {
                            //
                        }
                        

                        if (config[4].ToString() == values[4])
                        {
                            lblOpt3.Text = values[5];
                        }

                        if (config[5].ToString() == values[6])
                        {
                            lblOpt4.Text = values[7];
                        }

                        if (config[6].ToString() == values[8])
                        {
                            lblOpt5.Text = values[9];
                        }

                        if (config[7].ToString() == values[10])
                        {
                            lblOpt6.Text = values[11];
                        }

                        if (config[8].ToString() == values[12])
                        {
                            lblOpt7.Text = values[13];
                        }

                        if (config[9].ToString() == values[14])
                        {
                            lblOpt8.Text = values[15];
                        }

                        if (config[10].ToString() == values[16])
                        {
                            lblOpt9.Text = values[17];
                        }

                        if (config[11].ToString() == "-")
                        {
                            //Hyphen
                        }

                        if (config[12].ToString() + config[13].ToString() == values[18])
                        {
                            lblOpt10.Text = values[19];
                        }
                    }
                    reader.Close();
                    return;
                }

                //X-Series
                if (config[0] == '6')
                {
                    var reader = new StreamReader("xseries.dbs");

                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');

                        if (config[0].ToString() == values[0])
                        {
                            lblOpt1.Text = values[1];
                        }

                        if (config[1].ToString() + config[2].ToString() == values[2])
                        {
                            lblOpt2.Text = values[3];
                        }

                        if (config[3].ToString() == "-")
                        {
                            //
                        }

                        //if (config[4].ToString() == values[6])

                        //lblOpt4.Text = "...";

                        if (config[4].ToString() == values[4])
                        {
                            lblOpt3.Text = values[5];
                        }

                        if (config[5].ToString() == values[6])
                        {
                            lblOpt4.Text = values[7];
                        }

                        if (config[6].ToString() == values[8])
                        {
                            lblOpt5.Text = values[9];
                        }

                        if (config[7].ToString() == values[10])
                        {
                            lblOpt6.Text = values[11];
                        }

                        if (config[8].ToString() == values[12])
                        {
                            lblOpt7.Text = values[13];
                        }

                        if (config[9].ToString() == values[14])
                        {
                            lblOpt8.Text = values[15];
                        }

                        if (config[10].ToString() == values[16])
                        {
                            lblOpt9.Text = values[17];
                        }

                        if (config[11].ToString() == "-")
                        {
                            //Hyphen
                        }

                        if (config[12].ToString()+config[13].ToString() == values[18])
                        {
                            lblOpt10.Text = values[19];
                        }
                        /*
                        if (config[14].ToString() + config[15].ToString() == values[20])
                        {
                            lblOpt11.Text = values[21];
                        }

                        if (config[16].ToString() + config[17].ToString() == values[22])
                        {
                            lblOpt12.Text = values[23];
                        }
                        */
                    }
                    reader.Close();
                    return;
                }
            }


            //R-SERIES Find
            if (config[0] == '1' || config[0] == '3')
            {
                var reader = new StreamReader("rseries.dbs");

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    if (config[0].ToString() == values[0])
                    {
                        lblOpt1.Text = values[1];
                    }

                    if (config[1].ToString() + config[2].ToString() == values[2])
                    {
                        lblOpt2.Text = values[3];
                    }

                    if (config[3].ToString() == values[4])
                    {
                        lblOpt3.Text = values[5];
                    }

                    //if (config[4].ToString() == values[6])
                    
                    lblOpt4.Text = "...";

                    if (config[5].ToString() == values[8])
                    {
                        lblOpt5.Text = values[9];
                    }

                    if (config[6].ToString() == values[10])
                    {
                        lblOpt6.Text = values[11];
                    }

                    if (config[7].ToString() == values[12])
                    {
                        lblOpt7.Text = values[13];
                    }

                    if (config[8].ToString() == values[14])
                    {
                        lblOpt8.Text = values[15];
                    }

                    if (config[9].ToString() == values[16])
                    {
                        lblOpt9.Text = values[17];
                    }

                    if (config[10].ToString() == values[18])
                    {
                        lblOpt10.Text = values[19];
                    }

                    if (config[11].ToString() == values[20])
                    {
                        lblOpt11.Text = values[21];
                    }

                    if (config[12].ToString() == values[22])
                    {
                        lblOpt12.Text = values[23];
                    }

                    if (config[13].ToString() == values[24])
                    {
                        lblOpt13.Text = values[25];
                    }

                    if (config[14].ToString() + config[15].ToString() == values[26])
                    {
                        lblOpt14.Text = values[27];
                    }

                    if (config[16].ToString() == values[28])
                    {
                        lblOpt15.Text = values[29];
                    }

                }
                reader.Close();
                return;
            }

            //M-SERIES Find
            if (config[0] == '4' || config[0] == '6')
            {
                var reader = new StreamReader("mseries.dbs");

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    if (config[0].ToString() == values[0])
                    {
                        lblOpt1.Text = values[1];
                    }

                    if (config[1].ToString()+config[2].ToString() == values[2])
                    {
                        lblOpt2.Text = values[3];
                    }

                    if (config[3].ToString() == values[4])
                    {
                        lblOpt3.Text = values[5];
                    }

                    if (config[4].ToString() == values[6])
                    {
                        lblOpt4.Text = values[7];
                    }

                    if (config[5].ToString() == values[8])
                    {
                        lblOpt5.Text = values[9];
                    }

                    if (config[6].ToString() == values[10])
                    {
                        lblOpt6.Text = values[11];
                    }

                    if (config[7].ToString() == values[12])
                    {
                        lblOpt7.Text = values[13];
                    }

                    if (config[8].ToString() == values[14])
                    {
                        lblOpt8.Text = values[15];
                    }

                    if (config[9].ToString() == values[16])
                    {
                        lblOpt9.Text = values[17];
                    }

                    if (config[10].ToString() == values[18])
                    {
                        lblOpt10.Text = values[19];
                    }

                    if (config[11].ToString() == values[20])
                    {
                        lblOpt11.Text = values[21];
                    }

                    if (config[12].ToString() == values[22])
                    {
                        lblOpt12.Text = values[23];
                    }

                    if (config[13].ToString() == values[24])
                    {
                        lblOpt13.Text = values[25];
                    }

                    if (config[14].ToString()+config[15].ToString() == values[26])
                    {
                        lblOpt14.Text = values[27];
                    }

                    if (config[16].ToString() == values[28])
                    {
                        lblOpt15.Text = values[29];
                    }

                }
                reader.Close();
                return;
            }

            //AED-PLUS/PRO Find
            if (config[0] == '2' || config[0] == '7' || config[0] == '9')
            {
                var reader = new StreamReader("aed.dbs");

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    if (config[0].ToString() == values[0])
                    {
                        lblOpt1.Text = values[1];
                    }

                    if (config[1].ToString() + config[2].ToString() == values[2])
                    {
                        lblOpt2.Text = values[3];
                    }

                    if (config[3].ToString() == values[4])
                    {
                        lblOpt3.Text = values[5];
                    }

                    if (config[4].ToString()+config[5].ToString() == values[6])
                    {
                        lblOpt4.Text = values[7];
                    }

                    if (config[6].ToString() == values[8])
                    {
                        lblOpt5.Text = values[9];
                    }

                    if (config[7].ToString() == values[10])
                    {
                        lblOpt6.Text = values[11];
                    }

                    if (config[8].ToString() == values[12])
                    {
                        lblOpt7.Text = values[13];
                    }

                    if (config[9].ToString()+config[10].ToString() == values[14])
                    {
                        lblOpt8.Text = values[15];
                    }

                    if (config[11].ToString()+config[12].ToString() == values[16])
                    {
                        lblOpt9.Text = values[17];
                    }

                    if (config[13].ToString() == values[18])
                    {
                        lblOpt10.Text = values[19];
                    }

                    if (config[14].ToString()+config[15].ToString() == values[20])
                    {
                        lblOpt11.Text = values[21];
                    }

                    if (config[16].ToString() == values[22])
                    {
                        lblOpt12.Text = values[23];
                    }
                }
                reader.Close();
                return;
            }

            //E-SERIES Find
            if (config[0] == '5' || config[0] == '8')
            {
                var reader = new StreamReader("eseries.dbs");

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    if (config[0].ToString() == values[0])
                    {
                        lblOpt1.Text = values[1];
                    }

                    if (config[1].ToString() + config[2].ToString() == values[2])
                    {
                        lblOpt2.Text = values[3];
                    }
                    /*
                    if (config[3].ToString() == values[4])
                    {
                        
                    }
                     */
                    lblOpt3.Text = "...";

                    if (config[4].ToString() == values[4])
                    {
                        lblOpt4.Text = values[5];
                    }

                    if (config[5].ToString() == values[6])
                    {
                        lblOpt5.Text = values[7];
                    }

                    if (config[6].ToString() == values[8])
                    {
                        lblOpt6.Text = values[9];
                    }

                    if (config[7].ToString() == values[12])
                    {
                        
                    }
                    lblOpt7.Text = "...";

                    if (config[8].ToString() == values[14])
                    {
                        
                    }
                    lblOpt8.Text = "...";

                    if (config[9].ToString() == values[10])
                    {
                        lblOpt9.Text = values[11];
                    }

                    //if (config[10].ToString()+config[11].ToString() == values[18])
                    
                    lblOpt10.Text = "...";

                    if (config[11].ToString() == values[12])
                    {
                        lblOpt11.Text = values[13];
                    }

                    if (config[12].ToString() == values[14])
                    {
                        lblOpt12.Text = values[15];
                    }

                    if (config[13].ToString() == values[16])
                    {
                        lblOpt13.Text = values[17];
                    }

                    if (config[14].ToString() + config[15].ToString() == values[18])
                    {
                        lblOpt14.Text = values[19];
                    }

                    if (config[16].ToString() == values[20])
                    {
                        lblOpt15.Text = values[21];
                    }

                }
                reader.Close();
                return;
            }
        }

        private void btnConfigFind_Click(object sender, EventArgs e)
        {
            if (radConfigExact.Checked == true)
            {
                FindConfig(txtConfig.Text);
                FindExactItemMatches(txtConfig.Text);
            }

            if (radConfigSimilar.Checked == true)
            {
                FindConfig(txtConfig.Text);
                FindSimilar(txtConfig.Text);
            }
        }

        private void FindExactItemMatches(string config)
        {
            listMatchConfig.Items.Clear();

            for (int i = 0; i < MasterList[0].Count; i++)
            {
                if (MasterList[1][i] == config)
                {
                    if (MasterList[2][i] == "OUT")
                    {
                        listMatchConfig.Items.Add(MasterList[0][i] + " - " + MasterList[2][i]);
                    }

                    if (MasterList[2][i] == "IN")
                    {
                        if (MasterList[11][i] == "Y")
                        {
                            listMatchConfig.Items.Add(MasterList[0][i] + " - READY TO SHIP");
                        }
                        else
                        {
                            listMatchConfig.Items.Add(MasterList[0][i] + " - not prepared");
                        }
                    }

                }
            }

            lblItemMatches.Text = listMatchConfig.Items.Count + " Loaner(s) match this Item Number";
        }

        private void lblItemMatches_Click(object sender, EventArgs e)
        {

        }

        private void listMatchConfig_DoubleClick(object sender, EventArgs e)
        {
            if (listMatchConfig.SelectedIndex >= 0)
            {
                var values = listMatchConfig.Items[listMatchConfig.SelectedIndex].ToString().Split(' ');

                ClearInfoFields();
                txtInfoSN.Clear();
                txtInfoSN.Text = values[0];

                Util.Animate(tabControl1, Util.Effect.Slide, 150, 180);
                tabControl1.SelectedIndex = 1;
                Util.Animate(tabControl1, Util.Effect.Slide, 150, 0);
            }
        }

        private string ConvertToItemString(string config)
        {
            string Item = "";

            //X-SERIES and PROPAQ Find
            if (config.Contains('-') == true)
            {
                //Propaq Find
                if (config[0] == '2' || config[0] == '3')
                {
                    var reader = new StreamReader("propaq.dbs");

                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');

                        if (config[0].ToString() == values[0])
                        {
                            Item=Item+" "+ values[1];
                        }

                        if (config[1].ToString() + config[2].ToString() == values[2])
                        {
                            Item = Item + " " + values[3];
                        }

                        if (config[3].ToString() == "-")
                        {
                            //
                        }


                        if (config[4].ToString() == values[4])
                        {
                            Item = Item + " " + values[5];
                        }

                        if (config[5].ToString() == values[6])
                        {
                            Item = Item + " " + values[7];
                        }

                        if (config[6].ToString() == values[8])
                        {
                            Item = Item + " " + values[9];
                        }

                        if (config[7].ToString() == values[10])
                        {
                            Item = Item + " " + values[11];
                        }

                        if (config[8].ToString() == values[12])
                        {
                            Item = Item + " " + values[13];
                        }

                        if (config[9].ToString() == values[14])
                        {
                            Item = Item + " " + values[15];
                        }

                        if (config[10].ToString() == values[16])
                        {
                            Item = Item + " " + values[17];
                        }

                        if (config[11].ToString() == "-")
                        {
                            //Hyphen
                        }

                        if (config[12].ToString() + config[13].ToString() == values[18])
                        {
                            Item = Item + " " + values[19];
                        }
                    }
                    reader.Close();
                    return Item;
                }

                //X-Series
                if (config[0] == '6')
                {
                    var reader = new StreamReader("xseries.dbs");

                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');

                        if (config[0].ToString() == values[0])
                        {
                            Item = Item + " " + values[1];
                        }

                        if (config[1].ToString() + config[2].ToString() == values[2])
                        {
                            Item = Item + " " + values[3];
                        }

                        if (config[3].ToString() == "-")
                        {
                            //
                        }

                        //if (config[4].ToString() == values[6])

                        //lblOpt4.Text = "...";

                        if (config[4].ToString() == values[4])
                        {
                            Item = Item + " " + values[5];
                        }

                        if (config[5].ToString() == values[6])
                        {
                            Item = Item + " " + values[7];
                        }

                        if (config[6].ToString() == values[8])
                        {
                            Item = Item + " " + values[9];
                        }

                        if (config[7].ToString() == values[10])
                        {
                            Item = Item + " " + values[11];
                        }

                        if (config[8].ToString() == values[12])
                        {
                            Item = Item + " " + values[13];
                        }

                        if (config[9].ToString() == values[14])
                        {
                            Item = Item + " " + values[15];
                        }

                        if (config[10].ToString() == values[16])
                        {
                            Item = Item + " " + values[17];
                        }

                        if (config[11].ToString() == "-")
                        {
                            //Hyphen
                        }

                        if (config[12].ToString() + config[13].ToString() == values[18])
                        {
                            Item = Item + " " + values[19];
                        }
                        /*
                        if (config[14].ToString() + config[15].ToString() == values[20])
                        {
                            lblOpt11.Text = values[21];
                        }

                        if (config[16].ToString() + config[17].ToString() == values[22])
                        {
                            lblOpt12.Text = values[23];
                        }
                        */
                    }
                    reader.Close();
                    return Item;
                }
            }


            //R-SERIES Find
            if (config[0] == '1' || config[0] == '3')
            {
                var reader = new StreamReader("rseries.dbs");

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    if (config[0].ToString() == values[0])
                    {
                        Item = Item + " " + values[1];
                    }

                    if (config[1].ToString() + config[2].ToString() == values[2])
                    {
                        Item = Item + " " + values[3];
                    }

                    if (config[3].ToString() == values[4])
                    {
                        Item = Item + " " + values[5];
                    }

                    //if (config[4].ToString() == values[6])

                    lblOpt4.Text = "...";

                    if (config[5].ToString() == values[8])
                    {
                        Item = Item + " " + values[9];
                    }

                    if (config[6].ToString() == values[10])
                    {
                        Item = Item + " " + values[11];
                    }

                    if (config[7].ToString() == values[12])
                    {
                        Item = Item + " " + values[13];
                    }

                    if (config[8].ToString() == values[14])
                    {
                        Item = Item + " " + values[15];
                    }

                    if (config[9].ToString() == values[16])
                    {
                        Item = Item + " " + values[17];
                    }

                    if (config[10].ToString() == values[18])
                    {
                        Item = Item + " " + values[19];
                    }

                    if (config[11].ToString() == values[20])
                    {
                        Item = Item + " " + values[21];
                    }

                    if (config[12].ToString() == values[22])
                    {
                        Item = Item + " " + values[23];
                    }

                    if (config[13].ToString() == values[24])
                    {
                        Item = Item + " " + values[25];
                    }

                    if (config[14].ToString() + config[15].ToString() == values[26])
                    {
                        Item = Item + " " + values[27];
                    }

                    if (config[16].ToString() == values[28])
                    {
                        Item = Item + " " + values[29];
                    }

                }
                reader.Close();
                return Item;
            }

            //M-SERIES Find
            if (config[0] == '4' || config[0] == '6')
            {
                var reader = new StreamReader("mseries.dbs");

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    if (config[0].ToString() == values[0])
                    {
                        Item = Item + " " + values[1];
                    }

                    if (config[1].ToString() + config[2].ToString() == values[2])
                    {
                        Item = Item + " " + values[3];
                    }

                    if (config[3].ToString() == values[4])
                    {
                        Item = Item + " " + values[5];
                    }

                    if (config[4].ToString() == values[6])
                    {
                        Item = Item + " " + values[7];
                    }

                    if (config[5].ToString() == values[8])
                    {
                        Item = Item + " " + values[9];
                    }

                    if (config[6].ToString() == values[10])
                    {
                        Item = Item + " " + values[11];
                    }

                    if (config[7].ToString() == values[12])
                    {
                        Item = Item + " " + values[13];
                    }

                    if (config[8].ToString() == values[14])
                    {
                        Item = Item + " " + values[15];
                    }

                    if (config[9].ToString() == values[16])
                    {
                        Item = Item + " " + values[17];
                    }

                    if (config[10].ToString() == values[18])
                    {
                        Item = Item + " " + values[19];
                    }

                    if (config[11].ToString() == values[20])
                    {
                        Item = Item + " " + values[21];
                    }

                    if (config[12].ToString() == values[22])
                    {
                        Item = Item + " " + values[23];
                    }

                    if (config[13].ToString() == values[24])
                    {
                        Item = Item + " " + values[25];
                    }

                    if (config[14].ToString() + config[15].ToString() == values[26])
                    {
                        Item = Item + " " + values[27];
                    }

                    if (config[16].ToString() == values[28])
                    {
                        Item = Item + " " + values[29];
                    }

                }
                reader.Close();
                return Item;
            }

            //AED-PLUS/PRO Find
            if (config[0] == '2' || config[0] == '7' || config[0] == '9')
            {
                var reader = new StreamReader("aed.dbs");

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    if (config[0].ToString() == values[0])
                    {
                        Item = Item + " " + values[1];
                    }

                    if (config[1].ToString() + config[2].ToString() == values[2])
                    {
                        Item = Item + " " + values[3];
                    }

                    if (config[3].ToString() == values[4])
                    {
                        Item = Item + " " + values[5];
                    }

                    if (config[4].ToString() + config[5].ToString() == values[6])
                    {
                        Item = Item + " " + values[7];
                    }

                    if (config[6].ToString() == values[8])
                    {
                        Item = Item + " " + values[9];
                    }

                    if (config[7].ToString() == values[10])
                    {
                        Item = Item + " " + values[11];
                    }

                    if (config[8].ToString() == values[12])
                    {
                        Item = Item + " " + values[13];
                    }

                    if (config[9].ToString() + config[10].ToString() == values[14])
                    {
                        Item = Item + " " + values[15];
                    }

                    if (config[11].ToString() + config[12].ToString() == values[16])
                    {
                        Item = Item + " " + values[17];
                    }

                    if (config[13].ToString() == values[18])
                    {
                        Item = Item + " " + values[19];
                    }

                    if (config[14].ToString() + config[15].ToString() == values[20])
                    {
                        Item = Item + " " + values[21];
                    }

                    if (config[16].ToString() == values[22])
                    {
                        Item = Item + " " + values[23];
                    }
                }
                reader.Close();
                return Item;
            }

            //E-SERIES Find
            if (config[0] == '5' || config[0] == '8')
            {
                var reader = new StreamReader("eseries.dbs");

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    if (config[0].ToString() == values[0])
                    {
                        Item = Item + " " + values[1];
                    }

                    if (config[1].ToString() + config[2].ToString() == values[2])
                    {
                        Item = Item + " " + values[3];
                    }
                    /*
                    if (config[3].ToString() == values[4])
                    {
                        
                    }
                     */
                    lblOpt3.Text = "...";

                    if (config[4].ToString() == values[4])
                    {
                        Item = Item + " " + values[5];
                    }

                    if (config[5].ToString() == values[6])
                    {
                        Item = Item + " " + values[7];
                    }

                    if (config[6].ToString() == values[8])
                    {
                        Item = Item + " " + values[9];
                    }

                    if (config[7].ToString() == values[12])
                    {

                    }
                    lblOpt7.Text = "...";

                    if (config[8].ToString() == values[14])
                    {

                    }
                    lblOpt8.Text = "...";

                    if (config[9].ToString() == values[10])
                    {
                        Item = Item + " " + values[11];
                    }

                    //if (config[10].ToString()+config[11].ToString() == values[18])

                    lblOpt10.Text = "...";

                    if (config[11].ToString() == values[12])
                    {
                        Item = Item + " " + values[13];
                    }

                    if (config[12].ToString() == values[14])
                    {
                        Item = Item + " " + values[15];
                    }

                    if (config[13].ToString() == values[16])
                    {
                        Item = Item + " " + values[17];
                    }

                    if (config[14].ToString() + config[15].ToString() == values[18])
                    {
                        Item = Item + " " + values[19];
                    }

                    if (config[16].ToString() == values[20])
                    {
                        Item = Item + " " + values[21];
                    }

                }
                reader.Close();
                return Item;
            }
            return Item;
        }

        private void FindSimilar(string config)
        {
            listMatchConfig.Items.Clear();

            //AED-PLUS
            if (config[0].ToString() == "2" || config[0].ToString() == "7")
            {
                lblConfigWarning.Text = "WARNING: Similar matches are only matched by Model, Voice/No Voice and Language.";
                for (int i = 0; i < MasterList[0].Count; i++)
                {
                    //Match Model
                    if ((MasterList[9][i] == "AED-PLUS") && (config[9].ToString() + config[10].ToString() != "99") && MasterList[1][i].Length>10)
                    {
                        if ((MasterList[1][i][14].ToString() + MasterList[1][i][15].ToString() == config[14].ToString() + config[15].ToString()) && (config[7].ToString() == MasterList[1][i][7].ToString()))
                        {
                            if (MasterList[2][i] == "OUT")
                            {
                                listMatchConfig.Items.Add(MasterList[0][i] + " - " + MasterList[2][i]);
                            }

                            if (MasterList[2][i] == "IN")
                            {
                                if (MasterList[11][i] == "Y")
                                {
                                    listMatchConfig.Items.Add(MasterList[0][i] + " - READY TO SHIP");
                                }
                                else
                                {
                                    listMatchConfig.Items.Add(MasterList[0][i] + " - not prepared");
                                }
                            }
                        }
                    }
                }
            }

            //AED-Pro
            if (config[0].ToString() == "9" || config[0].ToString() == "7")
            {
                lblConfigWarning.Text = "WARNING: Similar matches are only matched by Model, Voice/No Voice, and Language.";
                for (int i = 0; i < MasterList[0].Count; i++)
                {
                    //Match Type
                    if (MasterList[9][i] == "AED-PRO" && (config[9].ToString() + config[10].ToString() == "99") && MasterList[1][i].Length > 10)
                    {
                        if ((MasterList[1][i][14].ToString() + MasterList[1][i][15].ToString() == config[14].ToString() + config[15].ToString())  && (config[7].ToString() == MasterList[1][i][7].ToString()))
                        {
                            if (MasterList[2][i] == "OUT")
                            {
                                listMatchConfig.Items.Add(MasterList[0][i] + " - OUT");
                            }

                            if (MasterList[2][i] == "IN")
                            {
                                if (MasterList[11][i] == "Y")
                                {
                                    listMatchConfig.Items.Add(MasterList[0][i] + " - READY TO SHIP");
                                }
                                else
                                {
                                    listMatchConfig.Items.Add(MasterList[0][i] + " - not prepared");
                                }
                            }
                        }
                    }
                }
            }

            //Similar E-SERIES
            if (config[0].ToString() == "5" || config[0].ToString() == "8")
            {
                lblConfigWarning.Text = "WARNING: Similar matches are only matched by Model, Language, ECG type, Advisory/Manual, and Pace/No Pace.";

                for (int i = 0; i < MasterList[0].Count; i++)
                {
                    //Match Type
                    if (MasterList[9][i] == "E-SERIES" && MasterList[1][i].Length > 10)
                    {
                        //Matches on ECG (3-5/12), Language, Advisory/Manual, Pace/No Pace,
                        if (config[4].ToString() == MasterList[1][i][4].ToString() && (MasterList[1][i][14].ToString() + MasterList[1][i][15].ToString() == config[14].ToString() + config[15].ToString()) && (config[12].ToString()==MasterList[1][i][12].ToString()) && (config[11].ToString()==MasterList[1][i][11].ToString()))
                        {
                            if (MasterList[2][i] == "OUT")
                            {
                                listMatchConfig.Items.Add(MasterList[0][i] + " - OUT");
                            }

                            if (MasterList[2][i] == "IN")
                            {
                                if (MasterList[11][i] == "Y")
                                {
                                    listMatchConfig.Items.Add(MasterList[0][i] + " - READY TO SHIP");
                                }
                                else
                                {
                                    listMatchConfig.Items.Add(MasterList[0][i] + " - not prepared");
                                }
                            }

                        }
                    }
                }
            }

            lblItemMatches.Text = listMatchConfig.Items.Count + " Loaner(s) match this Item Number";
        }

        private void radConfigSimilar_CheckedChanged(object sender, EventArgs e)
        {
            if (radConfigSimilar.Checked == true)
            {
                FindConfig(txtConfig.Text);
                FindSimilar(txtConfig.Text);
            }
        }

        private void radConfigExact_CheckedChanged(object sender, EventArgs e)
        {
            if (radConfigExact.Checked == true)
            {
                FindConfig(txtConfig.Text);
                FindExactItemMatches(txtConfig.Text);
            }
        }

        private void PopulateCustomerInfo()
        {
            if (txtInfoAccount.Text != "")
            {
                for (int j = 0; j < ContactList[0].Count; j++)
                {
                    if (ContactList[0][j] == txtInfoAccount.Text)
                    {
                        lblInfoCustName.Text = ContactList[2][j];
                        lblInfoAccount.Text = ContactList[1][j];
                        lblInfoCustNum.Text = ContactList[0][j];
                        lblInfoContact.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ContactList[3][j].ToLower());
                        lblInfoPhone.Text = ContactList[4][j];
                        lblInfoEmail.Text = ContactList[5][j].ToLower();
                        if (login == 2 || login == 3)
                        {
                            if (lblInfoEmail.Text == "UNK" || lblInfoEmail.Text == "" || lblInfoEmail.Text == "...")
                            {
                                btnEmail.Visible = false;
                            }
                            else
                            {
                                btnEmail.Visible = true;
                            }
                        }
                        for (int k = 0; k < MasterList[14].Count; k++)
                        {
                            if (MasterList[14][k] == lblInfoCustNum.Text && MasterList[2][k] == "OUT")
                            {
                                if (MasterList[0][k] == txtInfoSN.Text)
                                {
                                    listInfoPossess.Items.Add(MasterList[0][k] + " - Currently Viewing");
                                }
                                else
                                {
                                    listInfoPossess.Items.Add(MasterList[0][k]);
                                }
                            }
                        }
                        lblInfoPossesion.Text = listInfoPossess.Items.Count.ToString() + " Loaner(s) in possession";
                        return;
                    }
                }
                ClearInfoFields();
            }
        }

        private void txtInfoAccount_TextChanged(object sender, EventArgs e)
        {
            PopulateCustomerInfo();
        }

        private void btnInfoEditCust_Click(object sender, EventArgs e)
        {
            Util.Animate(tabControl1, Util.Effect.Slide, 150, 0);
            tabControl1.SelectedIndex = 3;
            tabControl2.SelectedIndex = 1;
            radCustEdit.Checked = true;
            txtCustNum.Text = txtInfoAccount.Text;
            Util.Animate(tabControl1, Util.Effect.Slide, 150, 180);
        }

        private void txtConfig_TextChanged(object sender, EventArgs e)
        {

        }

        private void PrintTextBoxContent()
        {


            #region Printer Selection
            PrintDialog printDlg = new PrintDialog();
            #endregion

            #region Create Document
            PrintDocument printDoc = new PrintDocument();
            printDoc.DocumentName = "Print Document";
            printDoc.PrintPage += printDoc_PrintPage;
            printDlg.Document = printDoc;
            #endregion

            if (printDlg.ShowDialog() == DialogResult.OK)
                printDoc.Print();
        }

        void printDoc_PrintPage(object sender, PrintPageEventArgs e)
        {
            FontFamily[] fontFamilies;

            PrivateFontCollection pfc = new PrivateFontCollection();
            pfc.AddFontFile("Code39.ttf");
            fontFamilies = pfc.Families;
            Font code39 = new Font(fontFamilies[0], 20, FontStyle.Regular);
            Font regular = new Font(txtInfoSN.Font.FontFamily, 16, FontStyle.Regular);


            e.Graphics.DrawString("*" + txtInfoSN.Text + "*", code39, Brushes.Black, 100, 15);
            e.Graphics.DrawString(txtInfoSN.Text, regular, Brushes.Black, 120, 48);

        }

        private void btnPrintBarcode_Click(object sender, EventArgs e)
        {
            PrintTextBoxContent();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("clee@zoll.com");
                oRecip.Resolve();
                
                oMsg.Subject = "Bug Report/Feature Request";
                oMsg.Body = "Hi Calvin!\n\nHere is my bug report/feature request: \n\nThanks,\n";
                //oMsg.Attachments.Add("c:/temp/test.txt", Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                oMsg.Display(true);
            }
            catch
            {
                MessageBox.Show("Email function failed. Make sure Outlook is open");
            }
        }

        private void listFields_DoubleClick(object sender, EventArgs e)
        {
            DisplayDetails();
        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SearchData();
            }
        }

        private void listFields_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.listFields.Sort();
        }

        private void ShowOut()
        {

            ClearFields();
            filter = 7;
            //DateTime parsedDate;
            //string[] formatstring = { "dd/MM/yyyy", "d/MM/yyyy", "d/M/yyyy", "dd/M/yyyy" };

            for (int i = 0; i < listG.Count(); i++)
            {
                //DateTime.TryParseExact(listG[i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                if (listC[i] == "OUT")
                {

                    ListViewItem lvi = new ListViewItem(MasterList[0][i]);

                    for (int j = 1; j < MasterList.Count; j++)
                    {
                        if (j < 6 || j > 8)
                        {
                            lvi.SubItems.Add(MasterList[j][i]);
                        }
                        else
                        {
                            if (MasterList[j][i] != "")
                            {
                                DateTime parsedDate;
                                string[] formatstring = { "d/m/yyyy", "dd/m/yyyy", "d/mm/yyyy", "dd/mm/yyyy" };

                                DateTime.TryParseExact(MasterList[j][i], formatstring, null, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal, out parsedDate);

                                lvi.SubItems.Add(parsedDate.ToString("yyyy/mm/dd"));
                            }
                            else
                            {
                                lvi.SubItems.Add("");
                            }
                        }
                        //lvi.SubItems.Add(MasterList[j][i]);
                    }
                    listFields.Items.Add(lvi);
                }
            }
        }

        private void btnOut_Click(object sender, EventArgs e)
        {
            ShowOut();
            SearchData();
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            tabControl1.Width = this.Width - 20;
            listFields.Width = tabControl1.Width - 180;
            grpSummary.Width = tabControl1.Width - 230;
            chartData.Width = grpSummary.Width - 230;
        }

        private void btnRequestGen_Click(object sender, EventArgs e)
        {
            if (txtGenLoanerSR.Text == "" || txtGenPN.Text == "" || txtGenSW.Text == "" || cmbGenUrgency.Text == "" || txtGenAccountNum.Text == "" || txtGenCustName.Text == "" || txtGenShipping.Text == "")
            {
                MessageBox.Show("Please fill in all mandatory fields (Marked with a *).");
                return;
            }
            
            if(Directory.Exists(@"T:\\! LOANER SR FOLDERS\\" + txtGenLoanerSR.Text + "\\") == false)
            {
                Directory.CreateDirectory(@"T:\\! LOANER SR FOLDERS\\" + txtGenLoanerSR.Text + "\\");
            }
            string output = @"T:\\! LOANER SR FOLDERS\\"+txtGenLoanerSR.Text + "\\"+txtGenLoanerSR.Text+"_request.docx";
            
            
            string acc = "• ";
            string data = "";

            List<CheckBox> accessories = new List<CheckBox>();
            accessories.Add(checkBox1);
            accessories.Add(checkBox2);
            accessories.Add(checkBox3);
            accessories.Add(checkBox4);
            accessories.Add(checkBox5);
            accessories.Add(checkBox6);
            accessories.Add(checkBox7);
            accessories.Add(checkBox8);
            accessories.Add(checkBox9);
            accessories.Add(checkBox10);
            accessories.Add(checkBox11);
            accessories.Add(checkBox12);
            accessories.Add(checkBox13);
            accessories.Add(checkBox14);
            accessories.Add(checkBox15);
            accessories.Add(checkBox16);
            accessories.Add(checkBox17);
            accessories.Add(checkBox18);
            accessories.Add(checkBox19);
            accessories.Add(checkBox20);
            accessories.Add(checkBox21);
            accessories.Add(checkBox22);
            accessories.Add(checkBox23);
            accessories.Add(checkBox24);
            accessories.Add(checkBox25);
            accessories.Add(checkBox26);
            accessories.Add(checkBox27);
            accessories.Add(checkBox28);
            accessories.Add(checkBox29);
            accessories.Add(checkBox30);
            accessories.Add(checkBox31);
            accessories.Add(checkBox32);
            accessories.Add(checkBox33);
            accessories.Add(checkBox34);
            accessories.Add(checkBox35);
            accessories.Add(checkBox36);
            accessories.Add(checkBox37);
            accessories.Add(checkBox38);

            foreach (CheckBox i in accessories)
            {
                if (i.Checked == true)
                {
                    switch (i.Text)
                    {
                        case "SpO2 Cable":
                            acc = acc + i.Text + " - #spo2cable#\n• ";
                            break;

                        case "SpO2 Sensor":
                            acc = acc + i.Text + " - #spo2sensor#\n• ";
                            break;

                        case "EtCO2 Sensor":
                            acc = acc + i.Text + " - #etco2sensor#\n• ";
                            break;

                        case "Battery 1":
                            acc = acc + i.Text + " - #battery1#\n• ";
                            break;

                        case "Battery 2":
                            acc = acc + i.Text + " - #battery2#\n• ";
                            break;

                        case "ROC Adaptor":
                            acc = acc + i.Text + " - #roc#\n• ";
                            break;

                        case "Pads 1":
                            acc = acc + i.Text + " - #pads1#\n• ";
                            break;

                        case "Pads 2":
                            acc = acc + i.Text + " - #pads2#\n• ";
                            break;

                        case "Paddles":
                            acc = acc + i.Text + " - #paddles#\n• ";
                            break;

                        case "Data Card":
                            acc = acc + i.Text + " - #datacard#\n• ";
                            break;

                        default:
                            acc = acc + i.Text + "\n• ";
                            break;
                    }
                    
                    data = data + "1";
                }
                else
                {
                    data = data + "0";
                }
                
            }
            acc = acc.TrimEnd(' ');
            acc = acc.TrimEnd('•');

            var writer = new StreamWriter(@"T:\\! LOANER SR FOLDERS\\" + txtGenLoanerSR.Text + "\\" + txtGenLoanerSR.Text + "_OutgoingInventory.txt");
            writer.WriteLine("<"+txtGenCompSR.Text+"<" + txtGenAccountNum.Text + "<" + txtGenCustName.Text);
            writer.WriteLine(data + "<<<<<<<<<<<<<<"+ txtGenComments.Text);
            writer.Close();

            var pending_reader = new StreamReader(@"T:\\Databases\\PendingLoaners.dbs");
            List<string> pendinglist = new List<string>();

            while (!pending_reader.EndOfStream)
            {
                string line = pending_reader.ReadLine();
                pendinglist.Add(line);
            }
            pending_reader.Close();

            if (!pendinglist.Contains(txtGenLoanerSR.Text))
            {
                var pending_writer = new StreamWriter(@"T:\\Databases\\PendingLoaners.dbs",true);
                pending_writer.WriteLine(txtGenLoanerSR.Text);
                pending_writer.Close();
            }

            output = @"T:\\! LOANER SR FOLDERS\\" + txtGenLoanerSR.Text + "\\" + txtGenLoanerSR.Text + "_request.docx";
            DocX request = DocX.Load("requesttemplate.docx");
            

            request.ReplaceText("#compsr#", txtGenCompSR.Text);
            request.ReplaceText("#compsn#", txtGenCompSerial.Text);
            request.ReplaceText("#loanersr#", txtGenLoanerSR.Text);
            request.ReplaceText("#date#", DateTime.Today.ToString("MMMM dd yyyy"));
            request.ReplaceText("#partnum#", txtGenPN.Text);
            request.ReplaceText("#swconfig#", txtGenSW.Text);
            request.ReplaceText("#urgency#", cmbGenUrgency.Text);
            request.ReplaceText("#techinit#", lblLogin.Text.Split('\"')[1]);
            request.ReplaceText("#custname#", txtGenCustName.Text);
            request.ReplaceText("#accountnum#", txtGenAccountNum.Text);
            request.ReplaceText("#contact#", txtGenContact.Text);
            request.ReplaceText("#phone#", txtGenPhone.Text);
            request.ReplaceText("#shipping#", txtGenShipping.Text);
            request.ReplaceText("#po#", txtGenPO.Text);
            

            if (txtGenComments.Text != "")
            {
                acc = acc + "Additional Comments:\n" + txtGenComments.Text;
            }

            request.ReplaceText("#acc#", acc);

            request.SaveAs(output);
            request.Dispose();
            Process.Start("WINWORD.EXE", "\"" + output + "\"");
        }

        private void chkAccessories_CheckedChanged(object sender, EventArgs e)
        {
            /*
            if (chkAccessories.Checked == true)
            {
                grpAcc.Visible = true;
            }
            else
            {
                grpAcc.Visible = false;
            }*/
        }

        private void txtGenAccountNum_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < ContactList[0].Count; i++)
            {
                if (txtGenAccountNum.Text == ContactList[0][i])
                {
                    txtGenCustName.Text = ContactList[2][i];
                    txtGenContact.Text = ContactList[3][i];
                    txtGenPhone.Text = ContactList[4][i];
                    return;
                }
            }
            txtGenCustName.Text = "";
            txtGenContact.Text = "";
            txtGenPhone.Text = "";
        }

        private void btnGenClear_Click(object sender, EventArgs e)
        {
            foreach(TextBox i in grpRequest.Controls.OfType<TextBox>())
            {
                i.Clear();
            }

            foreach (ComboBox i in grpRequest.Controls.OfType<ComboBox>())
            {
                i.SelectedIndex=-1;
            }

            foreach (CheckBox i in grpAcc.Controls.OfType<CheckBox>())
            {
                i.Checked = false;
            }
        }

        private void txtOpLoaner_TextChanged(object sender, EventArgs e)
        {
            if (radSend.Checked == true)
            {
                try
                {
                    List<CheckBox> accessories = new List<CheckBox>();
                    accessories.Add(checkBox1);
                    accessories.Add(checkBox2);
                    accessories.Add(checkBox3);
                    accessories.Add(checkBox4);
                    accessories.Add(checkBox5);
                    accessories.Add(checkBox6);
                    accessories.Add(checkBox7);
                    accessories.Add(checkBox8);
                    accessories.Add(checkBox9);
                    accessories.Add(checkBox10);
                    accessories.Add(checkBox11);
                    accessories.Add(checkBox12);
                    accessories.Add(checkBox13);
                    accessories.Add(checkBox14);
                    accessories.Add(checkBox15);
                    accessories.Add(checkBox16);
                    accessories.Add(checkBox17);
                    accessories.Add(checkBox18);
                    accessories.Add(checkBox19);
                    accessories.Add(checkBox20);
                    accessories.Add(checkBox21);
                    accessories.Add(checkBox22);
                    accessories.Add(checkBox23);
                    accessories.Add(checkBox24);
                    accessories.Add(checkBox25);
                    accessories.Add(checkBox26);
                    accessories.Add(checkBox27);
                    accessories.Add(checkBox28);
                    accessories.Add(checkBox29);
                    accessories.Add(checkBox30);
                    accessories.Add(checkBox31);
                    accessories.Add(checkBox32);
                    accessories.Add(checkBox33);
                    accessories.Add(checkBox34);
                    accessories.Add(checkBox35);
                    accessories.Add(checkBox36);
                    accessories.Add(checkBox37);
                    accessories.Add(checkBox38);

                    var reader = new StreamReader(@"T:\\! LOANER SR FOLDERS\\" + txtOpLoaner.Text + "\\" + txtOpLoaner.Text + "_OutgoingInventory.txt");

                    var line = reader.ReadLine();
                    var values = line.Split('<');

                    txtOpRepair.Text = values[1];
                    txtOpAccnt.Text = values[2];
                    txtOpCustomer.Text = values[3];

                    line = reader.ReadLine();
                    var dataval = line.Split('<');

                    for (int i = 0; i < dataval[0].Length; i++)
                    {
                        if (dataval[0][i] == '1')
                        {
                            string ser = "";
                            if(accessories[i].Text!="SpO2 Cable" && accessories[i].Text != "SpO2 Sensor"&& accessories[i].Text != "EtCO2 Sensor"&& accessories[i].Text != "Battery 1" && accessories[i].Text != "Battery 2" && accessories[i].Text != "ROC Adaptor" && accessories[i].Text != "Pads 1" && accessories[i].Text != "Pads 2" && accessories[i].Text != "Paddles" && accessories[i].Text != "Data Card")
                            {
                                ser = "n/a";
                            }
                            else
                            {
                                ser = "?";
                            }
                            string[] row = { accessories[i].Text, ser };
                            dgvAccessories.Rows.Add(row);
                        }
                    }

                    reader.Close();
                }
                catch
                {
                    txtOpAccnt.Clear();
                    txtOpCustomer.Clear();
                    dgvAccessories.Rows.Clear();
                }
            }

            if (radReceiving.Checked == true)
            {
                try
                {
                    List<CheckBox> accessories = new List<CheckBox>();
                    accessories.Add(checkBox1);
                    accessories.Add(checkBox2);
                    accessories.Add(checkBox3);
                    accessories.Add(checkBox4);
                    accessories.Add(checkBox5);
                    accessories.Add(checkBox6);
                    accessories.Add(checkBox7);
                    accessories.Add(checkBox8);
                    accessories.Add(checkBox9);
                    accessories.Add(checkBox10);
                    accessories.Add(checkBox11);
                    accessories.Add(checkBox12);
                    accessories.Add(checkBox13);
                    accessories.Add(checkBox14);
                    accessories.Add(checkBox15);
                    accessories.Add(checkBox16);
                    accessories.Add(checkBox17);
                    accessories.Add(checkBox18);
                    accessories.Add(checkBox19);
                    accessories.Add(checkBox20);
                    accessories.Add(checkBox21);
                    accessories.Add(checkBox22);
                    accessories.Add(checkBox23);
                    accessories.Add(checkBox24);
                    accessories.Add(checkBox25);
                    accessories.Add(checkBox26);
                    accessories.Add(checkBox27);
                    accessories.Add(checkBox28);
                    accessories.Add(checkBox29);
                    accessories.Add(checkBox30);
                    accessories.Add(checkBox31);
                    accessories.Add(checkBox32);
                    accessories.Add(checkBox33);
                    accessories.Add(checkBox34);
                    accessories.Add(checkBox35);
                    accessories.Add(checkBox36);
                    accessories.Add(checkBox37);
                    accessories.Add(checkBox38);

                    var reader = new StreamReader(@"T:\\! LOANER SR FOLDERS\\" + txtOpLoaner.Text + "\\" + txtOpLoaner.Text + "_OutgoingInventory.txt");

                    var line = reader.ReadLine();
                    var values = line.Split('<');

                    txtOpAccnt.Text = values[2];
                    txtOpCustomer.Text = values[3];

                    line = reader.ReadLine();
                    var dataval = line.Split('<');

                    for (int i = 0; i < dataval[0].Length; i++)
                    {
                        if (dataval[0][i] == '1')
                        {
                            string ser = "n/a";
                            if (accessories[i].Text == "SpO2 Cable") //accessories[i].Text != "Data Card")
                            {
                                ser = dataval[1];
                            }
                            if(accessories[i].Text == "SpO2 Sensor")
                            {
                                ser = dataval[2];
                            }
                            if (accessories[i].Text == "EtCO2 Sensor")
                            {
                                ser = dataval[3];
                            }
                            if (accessories[i].Text == "Battery 1")
                            {
                                ser = dataval[4];
                            }
                            if (accessories[i].Text == "Battery 2")
                            {
                                ser = dataval[5];
                            }
                            if (accessories[i].Text == "ROC Adaptor")
                            {
                                ser = dataval[6];
                            }
                            if (accessories[i].Text == "Pads 1")
                            {
                                ser = dataval[7];
                            }
                            if (accessories[i].Text == "Pads 2")
                            {
                                ser = dataval[8];
                            }
                            if (accessories[i].Text == "Paddles")
                            {
                                ser = dataval[9];
                            }
                            if (accessories[i].Text == "Data Card")
                            {
                                ser = dataval[10];
                            }

                            string[] row = { accessories[i].Text, ser };
                            dgvAccessories.Rows.Add(row);
                        }
                    }

                    reader.Close();
                }
                catch
                {
                    txtOpAccnt.Clear();
                    txtOpCustomer.Clear();
                    dgvAccessories.Rows.Clear();
                }
            }
        }

        private void txtGenCompSerial_TextChanged(object sender, EventArgs e)
        {
            if (txtGenCompSerial.Text.Length > 2)
            {
                if (txtGenPN.Text == "")
                {
                    foreach (CheckBox i in grpRequest.Controls.OfType<CheckBox>())
                    {
                        i.Checked = false;
                    }
                }

                if (txtGenCompSerial.Text[0].ToString() + txtGenCompSerial.Text[1].ToString() == "AF")
                {
                    //R Series accessories
                    checkBox8.Checked = true;
                    checkBox21.Text = "One-Step Cable for R-SERIES";
                    return;
                }
                else
                {
                    checkBox21.Text = "MFC Multi-Function Cable";
                }

                if (txtGenCompSerial.Text[0].ToString() + txtGenCompSerial.Text[1].ToString() == "AB")
                {
                    //E Series accessories
                    checkBox21.Checked = true;
                    checkBox22.Checked = true;
                    checkBox10.Checked = true;
                    checkBox12.Checked = true;
                    checkBox8.Checked = true;
                    checkBox28.Checked = true;

                    if (txtGenPN.Text.Length > 5)
                    {
                        if (txtGenPN.Text[txtGenPN.Text.Length - 3].ToString() + txtGenPN.Text[txtGenPN.Text.Length - 2].ToString() == "26")
                        {
                            checkBox1.Checked = true;
                            checkBox2.Checked = true;
                        }
                        else
                        {
                            checkBox1.Checked = false;
                            checkBox2.Checked = false;
                        }
                    }
                    else
                    {
                        checkBox1.Checked = false;
                        checkBox2.Checked = false;
                    }
                    return;
                }

                if (txtGenCompSerial.Text[0].ToString() == "T")
                {
                    //M Series accessories
                    checkBox21.Checked = true;
                    checkBox22.Checked = true;
                    checkBox8.Checked = true;
                    return;
                }

                

                if (txtGenCompSerial.Text[0].ToString() + txtGenCompSerial.Text[1].ToString() == "AR" || txtGenCompSerial.Text[0].ToString() + txtGenCompSerial.Text[1].ToString() == "AI")
                {
                    //X Series and Propaq accessories
                    checkBox10.Checked = true;
                    checkBox12.Checked = true;
                    checkBox8.Checked = true;
                    checkBox28.Checked = true;
                    return;
                }


            }
            foreach (CheckBox i in grpAcc.Controls.OfType<CheckBox>())
            {
                i.Checked = false;
            }
        }

        private void txtGenPN_TextChanged(object sender, EventArgs e)
        {
            if (txtGenPN.Text.Length > 3)
            {
                if (txtGenPN.Text[txtGenPN.Text.Length - 3].ToString() + txtGenPN.Text[txtGenPN.Text.Length - 2].ToString() == "26")
                {
                    if (txtGenCompSerial.Text.Length > 2)
                    {
                        if (txtGenCompSerial.Text[0].ToString() + txtGenCompSerial.Text[1].ToString() == "AB")
                        {
                            checkBox1.Checked = true;
                            checkBox2.Checked = true;
                        }
                        else
                        {
                            checkBox1.Checked = false;
                            checkBox2.Checked = false;
                        }
                    }
                }
                else
                {
                    checkBox1.Checked = false;
                    checkBox2.Checked = false;
                }
            }
        }

        private void chkReportDates_CheckedChanged(object sender, EventArgs e)
        {
            EnableTrendChart();
        }

        private void chkChartESeries_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkStack100_CheckedChanged(object sender, EventArgs e)
        {
            EnableStatusChart();
        }

        private void lstPending_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstPending.SelectedIndex > -1)
            {
                txtOpLoaner.Text = lstPending.Items[lstPending.SelectedIndex].ToString();
            }
        }
    }
}
