using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraBars.Helpers;
using DevExpress.Skins;
using DevExpress.LookAndFeel;
using DevExpress.UserSkins;
using DevExpress.XtraEditors;


namespace DXWindowsApplication
{
    public partial class Form1 : RibbonForm
    {
        protected DataSet dsStudents = new DataSet();

        public Form1()
        {
            InitializeComponent();
            InitSkinGallery();
            //InitGrid();

        }
        void InitSkinGallery()
        {
            SkinHelper.InitSkinGallery(rgbiSkins, true);
        }
        BindingList<Person> gridDataList = new BindingList<Person>();
        void InitGrid()
        {

            /*
            gridDataList.Add(new Person("John", "Smith"));
            gridDataList.Add(new Person("Gabriel", "Smith"));
            gridDataList.Add(new Person("Ashley", "Smith", "some comment"));
            gridDataList.Add(new Person("Adrian", "Smith", "some comment"));
            gridDataList.Add(new Person("Gabriella", "Smith", "some comment"));
            gridControl.DataSource = gridDataList;
             */
           
        }


        private void btnExit_ItemClick(object sender, ItemClickEventArgs e)
        {
            //show confirm message box
            if (MessageBox.Show("Do you want exit application ?", "Quit", MessageBoxButtons.YesNo) != DialogResult.OK)
            {
                this.Close();
            }
               
        }
        //event click button Attendance
        private void btnAttendance_ItemClick(object sender, ItemClickEventArgs e)
        {

        }

        private void btnOpen_ItemClick(object sender, ItemClickEventArgs e)
        {
            //open file dialog
            OpenFileDialog dlgOpen = new OpenFileDialog();
            dlgOpen.Title = "Choose students list from excel file ... ";
            //filter excel file
            dlgOpen.Filter = "Microsoft Excel File(*.xls,*xlsx)|*.xls;*.xlsx";

            //get my document path
            String mydocumentPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dlgOpen.InitialDirectory = mydocumentPath;
            //show dialog
            if (dlgOpen.ShowDialog() == DialogResult.OK)
            {
                //load data to gridview
                String filePath = dlgOpen.FileName;
                String fileExtension = System.IO.Path.GetExtension(filePath);
                DataSet ds = UltilityClass.loadDataFromExcelToDataSet(filePath, fileExtension);

                this.dsStudents = ds;
                //gridControl.DataSource = ds;
                this.dataGridView1.DataSource = this.dsStudents.Tables[0].DefaultView;
                this.gridControl.DataSource = this.dsStudents.Tables[0].DefaultView;
            }
        }




    }
}