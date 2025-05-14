using FarPoint.Win;
using FarPoint.Win.Spread;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public static class print
    {
        public enum ssPrintOrientType
        {
            ssPrintDefault = 0,
            ssPrintPortrait = 1,
            ssPrintLandscape = 2
        }

        public enum ssPrintType
        {
            ssNormalPrint = 13,
            ssSmartPrint = 32
        }

        public enum ssTitlePosition
        {
            ssLeft = 1,
            ssCenter = 2,
            ssRight = 3
        }

        public static ssPrintOrientType gudtPrintOrientType;
        public static ssPrintType gudtssPrintStyle;
        public static bool gblnPrintRunType;
        public static string gstrFooter;

        public static FpSpread gobjTargetPrintSS;
        public static Form gobjTargetPrintForm;
        public static int gintPrintPageCount;
        public static int gintPrintType;
        public static int gintPrintPageStart;
        public static int gintPrintPageEnd;
        public static bool gblnStandardReport;

        public static int gorgPMTop = 0;
        public static int gorgPMBottom = 0;
        public static int gorgPMLeft = 0;
        public static int gorgPMRight = 0;

        private static int SheetColumnHeaderColorCount = 0;
        private static PrintHeader ColumnHeaderVisible;

        private const int WM_SETICON = 0x0080;
        private static string sPreviewWindowTitle;
        private static Timer timer;

        public static async void gsSSPrint(
            bool preview,
            Form frm,
            FpSpread SS,
            string Title,
            bool StandardReport,
            string Sch_Cnd1 = "",
            string Sch_Cnd2 = "",
            string Sch_Desc1 = "",
            string Sch_Desc2 = "",
            bool ShowFooter = true,
            string Footer = "",
            int FooterHeight = -1,
            int FooterMargin = 10,
            ssTitlePosition TitlePosition = ssTitlePosition.ssLeft,
            ssPrintOrientType PrintOrientType = ssPrintOrientType.ssPrintLandscape,
            ssPrintType ssPrintStyle = ssPrintType.ssSmartPrint,
            int ssMarginTop = 1200,
            int ssMarginBottom = 400,
            int ssMarginLeft = 900,
            int ssMarginRight = 750,
            bool ssHeaderColor = false,
            bool ssPrintColor = false,
            bool ssPrintRowHeader = false,
            object ssPrintBorder = null,
            object ssPrintGrid = null,
            bool columnHeaderVisible = true,
            bool firstColumnLeftLineVisible = false,
            bool firstRowTopLineVisible = false,
            int endColumnIndex = -1,
            int endRowIndex = -1,
            bool isPDF = false,
            string PDFFilePath = "",
            float zoomFactor = 1f,
            Centering ssPrintCentering = Centering.None
            )
        {

            SheetView ss = SS.ActiveSheet;

            StyleInfo headerStyle = new StyleInfo
            {
                Border = new LineBorder(Color.Black, 1, false, false, false, true)
            };

            ss.ColumnHeader.DefaultStyle = headerStyle;

            object varShadowColor;
            varShadowColor = ss.ColumnHeader.Columns[0].BackColor;

            ss.PrintInfo.ShowTitle = PrintTitle.Hide;
            ss.PrintInfo.ShowSubtitle = PrintTitle.Hide;
            ss.PrintInfo.Centering = ssPrintCentering;

            try
            {
                gstrFooter = Footer;
                if (Sch_Cnd1.Length != 0 && Sch_Cnd2.Length == 0)
                {
                    Sch_Cnd2 = Sch_Cnd1;
                    Sch_Cnd1 = "";
                }

                if (Sch_Desc1.Length != 0 && Sch_Desc2.Length == 0)
                {
                    Sch_Desc2 = Sch_Desc1;
                    Sch_Desc1 = "";
                }

                ss.PrintInfo.JobName = "자료 인쇄......";
                ss.PrintInfo.AbortMessage = "인쇄중 취소시는 [Cancel]버튼을 누르세요.";

                ss.PrintInfo.ZoomFactor = zoomFactor;

                gudtPrintOrientType = PrintOrientType;
                gudtssPrintStyle = ssPrintStyle;

                gorgPMTop = (int)(ssMarginTop * 0.05); // 1twips = 0.05 point, 1 point = 0.01 inches
                gorgPMBottom = (int)(ssMarginBottom * 0.05); // 1twips = 0.05 point, 1 point = 0.01 inches
                gorgPMLeft = (int)(ssMarginLeft * 0.05); // 1twips = 0.05 point, 1 point = 0.01 inches
                gorgPMRight = (int)(ssMarginRight * 0.05); // 1twips = 0.05 point, 1 point = 0.01 inches

                if (ssPrintBorder == null) ssPrintBorder = StandardReport ? false : true;
                if (ssPrintGrid == null) ssPrintGrid = StandardReport ? false : true;

                ss.PrintInfo.Margin.Header = 10;
                ss.PrintInfo.Margin.Top = gorgPMTop;
                ss.PrintInfo.Margin.Bottom = gorgPMBottom;
                ss.PrintInfo.Margin.Left = gorgPMLeft;
                ss.PrintInfo.Margin.Right = gorgPMRight;
                ss.PrintInfo.Margin.Footer = FooterMargin;

                if (FooterHeight >= 0)
                    ss.PrintInfo.FooterHeight = FooterHeight;

                string strHeader = "";
                object strSsPosStr = "";
                gblnStandardReport = StandardReport;
                if (ss.Columns[0].CellType?.ToString() == "CheckBoxCellType")
                    ss.Columns[0].Visible = false;

                if (!preview)
                {
                    SS.Refresh();
                }

                switch (TitlePosition)
                {
                    case ssTitlePosition.ssLeft:
                        strSsPosStr = "/l";
                        break;
                    case ssTitlePosition.ssCenter:
                        strSsPosStr = "/c";
                        break;
                    case ssTitlePosition.ssRight:
                        strSsPosStr = "/r";
                        break;
                }

                if (StandardReport)
                {
                    ss.PrintInfo.Header = "";
                    ss.PrintInfo.Footer = gstrFooter;
                    goto Step_Print;
                }
                else
                {
                    if (Title != "") strHeader = strSsPosStr + "/fz\"20\"" + "/fb1/fu1" + Title + "/fb0/fu0/n/n/fz\"10\"";

                    if (Sch_Cnd1 != "") strHeader = strHeader + "/l" + Sch_Cnd1;
                    if (Sch_Desc1 != "")
                        strHeader = strHeader + "/r" + Sch_Desc1 + "  " + "/n";
                    else
                        strHeader = strHeader + "/n";
                    if (Sch_Cnd2 != "") strHeader = strHeader + "/fz\"10\"/l" + Sch_Cnd2;
                    if (Sch_Desc2 != "") strHeader = strHeader + "/r" + Sch_Desc2 + " ";

                    ss.PrintInfo.Header = strHeader + "/n";
                }

                if (ss.FrozenRowCount == 0)
                    ss.PrintInfo.ShowColumnHeader = PrintHeader.Show;
                else
                    ss.PrintInfo.ShowColumnHeader = PrintHeader.Hide;
                ss.PrintInfo.ShowRowHeader = ssPrintRowHeader ? PrintHeader.Show : PrintHeader.Hide;

            Step_Print:

                if (StandardReport)
                {
                    ss.PrintInfo.ShowBorder = ssPrintBorder == null ? false : (bool)ssPrintBorder;
                    ss.PrintInfo.ShowGrid = ssPrintGrid == null ? false : (bool)ssPrintGrid;
                    ss.PrintInfo.UseMax = false;
                }
                else
                {
                    ss.PrintInfo.ShowBorder = ssPrintBorder == null ? false : (bool)ssPrintBorder;
                    ss.PrintInfo.ShowGrid = ssPrintGrid == null ? false : (bool)ssPrintGrid;
                    ss.PrintInfo.UseMax = false;
                }

                if (preview)
                    ss.PrintInfo.EnhancePreview = true;

                if (ss.GetSelections() != null && ss.GetSelections().Length > 0 && (ss.GetSelections()[0].ColumnCount > 0 || ss.GetSelections()[0].RowCount > 0))
                {
                    Cell cell;
                    if (ss.GetSelections()[0].RowCount == -1)
                        cell = ss.Cells[0, ss.GetSelections()[0].Column, ss.RowCount - 1, ss.GetSelections()[0].Column];
                    else
                        cell = ss.Cells[ss.GetSelections()[0].Row, ss.GetSelections()[0].Column];
                    if (cell.RowSpan == ss.GetSelections()[0].RowCount && cell.ColumnSpan == ss.GetSelections()[0].ColumnCount)
                    {
                        ss.PrintInfo.PrintType = PrintType.CellRange;
                        ss.PrintInfo.ColStart = 0;
                        ss.PrintInfo.ColEnd = ss.Columns.Count - 1;
                        ss.PrintInfo.RowStart = 0;
                        ss.PrintInfo.RowEnd = ss.RowCount - 1;
                    }
                    else
                    {
                        ss.PrintInfo.PrintType = PrintType.CellRange;
                        ss.PrintInfo.ColStart = ss.GetSelections()[0].Column;
                        ss.PrintInfo.ColEnd = ss.GetSelections()[0].Column + (ss.GetSelections()[0].ColumnCount - 1);
                        ss.PrintInfo.RowStart = ss.GetSelections()[0].Row;
                        ss.PrintInfo.RowEnd = ss.GetSelections()[0].Row + (ss.GetSelections()[0].RowCount - 1);
                    }
                }
                else
                {
                    if (endColumnIndex >= 0 || endRowIndex >= 0) ss.PrintInfo.UseMax = false;
                    ss.PrintInfo.PrintType = PrintType.CellRange;
                    ss.PrintInfo.ColStart = 0;
                    ss.PrintInfo.ColEnd = endColumnIndex >= 0 ? endColumnIndex : ss.Columns.Count - 1;
                    ss.PrintInfo.RowStart = 0;
                    ss.PrintInfo.RowEnd = endRowIndex >= 0 ? endRowIndex : ss.RowCount - 1;
                }

                switch (gudtPrintOrientType)
                {
                    case ssPrintOrientType.ssPrintDefault:
                        ss.PrintInfo.Orientation = PrintOrientation.Auto;
                        break;
                    case ssPrintOrientType.ssPrintLandscape:
                        ss.PrintInfo.Orientation = PrintOrientation.Landscape;
                        break;
                    case ssPrintOrientType.ssPrintPortrait:
                        ss.PrintInfo.Orientation = PrintOrientation.Portrait;
                        break;
                }

                SheetColumnHeaderColorCount++;

                gobjTargetPrintForm = frm;
                gobjTargetPrintSS = SS;
                if (ShowFooter)
                    gsDisplayFooter(gstrFooter);

                ss.PrintInfo.ShowColor = ssPrintColor;

                if (!ssHeaderColor)
                {
                    for (int col = 0; col < ss.ColumnCount; col++)
                    {
                        ss.ColumnHeader.Columns[col].BackColor = Color.White;
                        ss.PrintInfo.ShowColor = true;
                    }
                }

                if (firstColumnLeftLineVisible)
                {
                    ss.AddColumns(0, 1);
                    ss.Columns[0].Width = 1;
                }

                if (firstRowTopLineVisible)
                {
                    ss.AddRows(0, 1);
                    ss.Rows[0].Height = 5;
                }

                if (!columnHeaderVisible)
                {
                    ss.PrintInfo.ShowColumnHeader = PrintHeader.Hide;
                }

                ColumnHeaderVisible = ss.PrintInfo.ShowColumnHeader;

                if (preview)
                {
                    ss.PrintInfo.Preview = preview;
                    ss.PrintInfo = ss.PrintInfo;
                    SS.PrintSheet(-1);
                }
                else if (isPDF)
                {
                    ss.PrintInfo.PrintToPdf = true;
                    ss.PrintInfo.PdfFileName = PDFFilePath;
                    ss.PrintInfo = ss.PrintInfo;
                    SS.PrintSheet(-1);
                }
                else
                {
                    ss.PrintInfo = ss.PrintInfo;
                    SS.PrintSheet(-1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                SheetColumnHeaderColorCount--;

                if (ss.Columns[0].CellType?.ToString() == "CheckBoxCellType")
                    ss.Columns[0].Visible = true;

                if (SheetColumnHeaderColorCount == 0)
                {
                    for (int col = 0; col < ss.ColumnCount; col++)
                    {
                        ss.ColumnHeader.Columns[col].ResetBackColor();
                    }
                }

                ss.PrintInfo.ShowColumnHeader = ColumnHeaderVisible;

                if (firstColumnLeftLineVisible)
                {
                    ss.RemoveColumns(0, 1);
                }

                if (firstRowTopLineVisible)
                {
                    ss.RemoveRows(0, 1);
                }

                await Task.Delay(3000);
            }
        }

        public static void gsDisplayFooter(object Footer)
        {
            if (gblnStandardReport) return;
            try
            {
                string strLogInDate;
                string strLogInTime;

                strLogInDate = DateTime.Now.ToString("yyyy/MM/dd");
                strLogInTime = DateTime.Now.ToString("HH:mm:ss");

                gobjTargetPrintSS.ActiveSheet.PrintInfo.Footer = "/n/fz\"8\"" + "/l/fu1" + "/fu0/l" + Footer + " [" + "/p" + " / " + "/pc" + "]" + "/c" + "Company" + "/r";
                if (strLogInDate != "")
                    gobjTargetPrintSS.ActiveSheet.PrintInfo.Footer = gobjTargetPrintSS.ActiveSheet.PrintInfo.Footer + strLogInDate + " [" + strLogInTime + "] - " + gobjTargetPrintForm.Name;
            }
            catch { }
        }
    }
}
