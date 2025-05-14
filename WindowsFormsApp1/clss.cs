using FarPoint.Win;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using FarPoint.Win.Spread.Model;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public static class clss
    {
        public static void Spread_Init(FpSpread spread, bool bLock, int colCheck = -1, bool isEditAllSelected = true, bool blnRowAdd = false, bool blnSpaceNum = true, bool isWordWrap = false)
        {
            SheetView sv = spread.ActiveSheet;

            // spread column resize 시 컬럼 색상 변경되는 문제 적용
            FlatColumnHeaderRenderer ch = new FlatColumnHeaderRenderer();
            if (isWordWrap)
                ch.WordWrap2 = false;
            else
                ch.WordWrap2 = true;
            ch.VerticalAlignment = VerticalAlignment.Center;
            ch.BackgroundStyle = BackStyle.Default;
            ch.GridLineNormalColor = Color.FromArgb(153, 153, 153);
            ch.GridLineActiveColor = Color.FromArgb(153, 153, 153);
            spread.ActiveSheet.ColumnHeader.Rows.Default.Renderer = ch;

            // column header
            //for (int rowIndex = 0; rowIndex < spread.ActiveSheet.ColumnHeader.Rows.Count; rowIndex++)
            //{
            //    spread.ActiveSheet.ColumnHeader.Rows[rowIndex].Height = 40;
            //    spread.ActiveSheet.ColumnHeader.Rows[rowIndex].Renderer = ch;
            //}

            spread.ScrollBarTrackPolicy = ScrollBarTrackPolicy.Both;
            spread.ShowActiveCell(VerticalPosition.Center, HorizontalPosition.Center);

            // 출력시 라인 중간이 하얗게 끊어지는 오류 대응
            spread.BorderCollapse = BorderCollapse.Collapse;

            spread.AllowDragDrop = false;
            spread.AllowDragFill = false;

            // row header width 조정
            spread.ActiveSheet.RowHeader.Columns[0].Width = 40;

            if (isEditAllSelected == true)
                spread.EditModeReplace = true;
            else
                spread.EditModeReplace = false;

            for (int col = 0; col < spread.ActiveSheet.Columns.Count; col++)
            {
                spread.ActiveSheet.Columns[col].VerticalAlignment = CellVerticalAlignment.Center;
                spread.ActiveSheet.Columns[col].Locked = bLock;
            }

            if (blnRowAdd == true)
            {
                spread.DialogKey += (sender, e) =>
                {
                    if (e.KeyCode == Keys.Down && sv.ActiveRowIndex + 1 == sv.Rows.Count)
                    {
                        sv.Rows.Add(sv.Rows.Count, 1);
                    }
                };
            }

            spread.DialogKey += (sender, e) =>
            {
                if (sv.Columns[sv.ActiveColumnIndex].CellType != null && sv.Columns[sv.ActiveColumnIndex].Locked == false && sv.Columns[sv.ActiveColumnIndex].CellType.ToString() == "NumberCellType")
                {
                    if (e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Add)
                    {
                        sv.Cells[sv.ActiveRowIndex, sv.ActiveColumnIndex].Value = Convert.ToDouble(sv.Cells[sv.ActiveRowIndex, sv.ActiveColumnIndex].Value) * 1000;
                    }
                }

                if (e.KeyCode == Keys.Enter)
                    spread.StopCellEditing();
            };

            if (blnSpaceNum)
            {
                spread.EditModeOn += (sender, e) =>
                {
                    if (sv.Columns[sv.ActiveColumnIndex].CellType != null && sv.Columns[sv.ActiveColumnIndex].CellType.ToString() == "NumberCellType")
                    {
                        string cellValueString = sv.GetValue(sv.ActiveRowIndex, sv.ActiveColumnIndex)?.ToString();

                        try
                        {
                            // 값이 null이 아니고, 숫자로 변환 가능한지 확인
                            if (string.IsNullOrEmpty(cellValueString) || Convert.ToDouble(sv.GetValue(sv.ActiveRowIndex, sv.ActiveColumnIndex)?.ToString()) == 0)
                            {
                                sv.SetText(sv.ActiveRowIndex, sv.ActiveColumnIndex, "0");
                            }
                        }
                        catch { }
                    }
                    else if (sv.Columns[sv.ActiveColumnIndex].CellType != null && sv.Columns[sv.ActiveColumnIndex].CellType.ToString() == "CurrencyCellType")
                    {
                        string cellValueString = sv.GetValue(sv.ActiveRowIndex, sv.ActiveColumnIndex)?.ToString();

                        // 값이 null이 아니고, 숫자로 변환 가능한지 확인
                        if (string.IsNullOrEmpty(cellValueString) || Convert.ToDouble(sv.GetValue(sv.ActiveRowIndex, sv.ActiveColumnIndex)?.ToString()) == 0)
                        {
                            sv.SetText(sv.ActiveRowIndex, sv.ActiveColumnIndex, "0");
                        }
                    }
                };

                spread.EditModeOff += (sender, e) =>
                {
                    if (sv.Columns[sv.ActiveColumnIndex].CellType != null && sv.Columns[sv.ActiveColumnIndex].CellType.ToString() == "NumberCellType")
                    {
                        try
                        {
                            if (Convert.ToDouble(sv.GetValue(sv.ActiveRowIndex, sv.ActiveColumnIndex)?.ToString()) == 0)
                                sv.SetText(sv.ActiveRowIndex, sv.ActiveColumnIndex, "");
                        }
                        catch { }
                    }
                    else if (sv.Columns[sv.ActiveColumnIndex].CellType != null && sv.Columns[sv.ActiveColumnIndex].CellType.ToString() == "CurrencyCellType")
                    {
                        if (Convert.ToDouble(sv.GetValue(sv.ActiveRowIndex, sv.ActiveColumnIndex)?.ToString()) == 0)
                            sv.SetText(sv.ActiveRowIndex, sv.ActiveColumnIndex, "");
                    }
                };
            }


            spread.KeyDown += (sender, e) =>
            {
                if (e.Control && e.KeyCode == Keys.C)
                {
                    e.Handled = true;
                    spread.ActiveSheet.ClipboardCopy(ClipboardCopyOptions.AsStringSkipHidden);
                }
                else if (e.Control && e.KeyCode == Keys.X)
                {
                    e.Handled = true;

                    spread.ActiveSheet.ClipboardCopy(ClipboardCopyOptions.AsStringSkipHidden);

                    CellRange cellRange = spread.ActiveSheet.GetSelection(0);
                    if (cellRange == null) return;
                    for (int row = cellRange.Row; row < cellRange.Row + cellRange.RowCount; row++)
                    {
                        bool isChanged = false;

                        for (int col = cellRange.Column; col < cellRange.Column + cellRange.ColumnCount; col++)
                        {
                            if (spread.ActiveSheet.Cells[row, col].Locked == false &&
                                spread.ActiveSheet.Rows[row].Locked == false &&
                                spread.ActiveSheet.Columns[col].Locked == false &&
                                spread.ActiveSheet.OperationMode != OperationMode.ReadOnly &&
                                spread.ActiveSheet.Columns[col].CellType?.ToString().ToLower() != "comboboxcelltype" &&
                                spread.ActiveSheet.Cells[row, col].CellType?.ToString().ToLower() != "comboboxcelltype")
                            {
                                spread.ActiveSheet.SetActiveCell(row, col);
                                spread.EditModePermanent = true;
                                spread.ActiveSheet.SetValue(row, col, "");
                                spread.EditModePermanent = false;
                                isChanged = true;
                            }

                            ComboBoxCellType comboboxcelltype = null;
                            if (spread.ActiveSheet.Columns[col].CellType?.ToString().ToLower() == "comboboxcelltype")
                            {
                                comboboxcelltype = (spread.ActiveSheet.Columns[col].CellType) as ComboBoxCellType;
                            }
                            if (spread.ActiveSheet.Cells[row, col].CellType?.ToString().ToLower() == "comboboxcelltype")
                            {
                                comboboxcelltype = (spread.ActiveSheet.Cells[row, col].CellType) as ComboBoxCellType;
                            }
                            if (comboboxcelltype != null)
                            {
                                if (spread.ActiveSheet.Cells[row, col].Locked == false &&
                                    spread.ActiveSheet.Rows[row].Locked == false &&
                                    spread.ActiveSheet.Columns[col].Locked == false &&
                                    spread.ActiveSheet.OperationMode != OperationMode.ReadOnly)
                                {
                                    spread.ActiveSheet.SetActiveCell(row, col);
                                    spread.EditModePermanent = true;
                                    spread.ActiveSheet.Cells[row, col].Text = comboboxcelltype.Items[0];
                                    spread.EditModePermanent = false;
                                }
                            }
                        }

                        if (isChanged && colCheck >= 0)
                        {
                            spread.ActiveSheet.SetValue(row, colCheck, 1);
                        }

                        isChanged = false;
                    }
                }
                else if (e.Control && e.KeyCode == Keys.V)
                {
                    e.Handled = true;

                    if (!spread.ActiveSheet.ActiveColumn.Locked && spread.ActiveSheet.OperationMode != OperationMode.ReadOnly)
                    {
                        spread.ActiveSheet.ClipboardPaste(ClipboardPasteOptions.AsStringSkipHidden);

                        if (colCheck >= 0)
                        {
                            CellRange cellRange = spread.ActiveSheet.GetSelection(0);
                            if (cellRange != null)
                            {
                                if (cellRange.RowCount == -1)
                                {
                                    for (int row = 0; row < spread.ActiveSheet.RowCount; row++)
                                    {
                                        spread.ActiveSheet.SetValue(row, colCheck, 1);
                                    }
                                }
                                else
                                {
                                    for (int row = cellRange.Row; row < cellRange.Row + cellRange.RowCount; row++)
                                    {
                                        spread.ActiveSheet.SetValue(row, colCheck, 1);
                                    }
                                }
                            }
                        }
                    }
                }
                else if (e.KeyCode == Keys.Back)
                {
                    e.Handled = true;

                    CellRange cellRange = spread.ActiveSheet.GetSelection(0);
                    if (cellRange == null) return;
                    for (int row = cellRange.Row; row < cellRange.Row + cellRange.RowCount; row++)
                    {
                        bool isChanged = false;

                        for (int col = cellRange.Column; col < cellRange.Column + cellRange.ColumnCount; col++)
                        {
                            if (spread.ActiveSheet.Cells[row, col].Locked == false &&
                                spread.ActiveSheet.Rows[row].Locked == false &&
                                spread.ActiveSheet.Columns[col].Locked == false &&
                                spread.ActiveSheet.OperationMode != OperationMode.ReadOnly &&
                                spread.ActiveSheet.Columns[col].CellType?.ToString().ToLower() != "comboboxcelltype" &&
                                spread.ActiveSheet.Cells[row, col].CellType?.ToString().ToLower() != "comboboxcelltype")
                            {
                                spread.ActiveSheet.SetActiveCell(row, col);
                                spread.EditModePermanent = true;
                                spread.ActiveSheet.SetValue(row, col, "");
                                spread.EditModePermanent = false;
                                isChanged = true;
                            }

                            ComboBoxCellType comboboxcelltype = null;
                            if (spread.ActiveSheet.Columns[col].CellType?.ToString().ToLower() == "comboboxcelltype")
                            {
                                comboboxcelltype = (spread.ActiveSheet.Columns[col].CellType) as ComboBoxCellType;
                            }
                            if (spread.ActiveSheet.Cells[row, col].CellType?.ToString().ToLower() == "comboboxcelltype")
                            {
                                comboboxcelltype = (spread.ActiveSheet.Cells[row, col].CellType) as ComboBoxCellType;
                            }
                            if (comboboxcelltype != null)
                            {
                                if (spread.ActiveSheet.Cells[row, col].Locked == false &&
                                    spread.ActiveSheet.Rows[row].Locked == false &&
                                    spread.ActiveSheet.Columns[col].Locked == false &&
                                    spread.ActiveSheet.OperationMode != OperationMode.ReadOnly)
                                {
                                    spread.ActiveSheet.SetActiveCell(row, col);
                                    spread.EditModePermanent = true;
                                    spread.ActiveSheet.Cells[row, col].Text = comboboxcelltype.Items[0];
                                    spread.EditModePermanent = false;
                                }
                            }
                        }

                        if (isChanged && colCheck >= 0)
                        {
                            spread.ActiveSheet.SetValue(row, colCheck, 1);
                        }

                        isChanged = false;
                    }
                }
            };

            // 헤더 라인 추가
            LineBorder headerLineBorder = new LineBorder(Color.FromArgb(255, 128, 128, 128), 1, true, true, true, true);
            spread.ActiveSheet.ColumnHeader.DefaultStyle.Border = headerLineBorder;
            spread.ActiveSheet.RowHeader.DefaultStyle.Border = headerLineBorder;
            spread.ActiveSheet.SheetCorner.DefaultStyle.Border = headerLineBorder;

        }
    }
}
