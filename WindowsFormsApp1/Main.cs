using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();

            clss.Spread_Init(fpSpread1, true);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            print.gsSSPrint(
                true,
                this,
                fpSpread1,
                this.Text,
                false,
                TitlePosition: print.ssTitlePosition.ssCenter,
                PrintOrientType: print.ssPrintOrientType.ssPrintLandscape,
                ssPrintStyle: print.ssPrintType.ssSmartPrint,
                zoomFactor: 1F
            );
        }
    }
}