using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; // ad neki egy nevet, az excel helyére lehetne teljesen mast írni
using System.Reflection;

namespace adatb_export_excel
{
    public partial class Form1 : Form

    {
        List<Flat> Flats;
        RealEstateEntities context = new RealEstateEntities();

        void LoadData()
        {
            Flats = context.Flat.ToList();
        }
        public Form1()
        {
            InitializeComponent();
            LoadData();

        }
    }
}
