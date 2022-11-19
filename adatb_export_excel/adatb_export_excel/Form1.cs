using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; // ad a fajlnak egy nevet, az excel helyére lehetne teljesen mast írni
using System.Reflection;

namespace adatb_export_excel
{
    public partial class Form1 : Form

    {
        List<Flat> Flats;
        RealEstateEntities context = new RealEstateEntities();

        Excel.Application xlApp; // A Microsoft Excel alkalmazás
        Excel.Workbook xlWB; // A létrehozott munkafüzet
        Excel.Worksheet xlSheet; // Munkalap a munkafüzeten belül

        void LoadData()
        {
            Flats = context.Flat.ToList();
        }

        void CreateExcel()
        {
            try
            {
                xlApp = new Excel.Application();
                xlWB = xlApp.Workbooks.Add();
                xlSheet = xlWB.ActiveSheet();

                CreateTable();

                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Source + '\n' + ex.Message);
                xlWB.Close(false);
                xlApp.Quit();
                xlApp = null;
                xlWB = null;
            }
        }

        private void CreateTable()
        {
            
        }

        public Form1()
        {
            InitializeComponent();
            LoadData();

        }
    }
}
