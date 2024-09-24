using DevExpress.Spreadsheet;
using QuanLyMayMoc.Common.Constant;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMayMoc.Common.Helper
{
    public static class SpreadsheetHelper
    {
        public static void ReplaceAllFormulaAfterImport(CellRange range)
        {
            ////          range.Worksheet.Workbook.History.IsEnabled = false;
            range.Worksheet.Workbook.BeginUpdate();

            SearchOptions MySearchOptions = new SearchOptions()
            {
                SearchIn = SearchIn.Values,
                MatchEntireCellContents = false,
            };

            var res = range.Search(MyConstant.PrefixFormula, MySearchOptions);
            foreach (var cell in res)
            {
                cell.Formula = cell.Value.ToString().Replace(MyConstant.PrefixFormula, "");
            }
            range.Worksheet.Calculate();

            res = range.Search(MyConstant.PrefixMerge, MySearchOptions);

            List<CellRange> ranges = new List<CellRange>();
            var ws = range.Worksheet;
            foreach (var cell in res)
            {
                string val = cell.Value.ToString().Replace(MyConstant.PrefixMerge, "");
                bool isParse = int.TryParse(val, out int ind);
                if (isParse)
                {
                    ws.Range.FromLTRB(cell.ColumnIndex, ind, cell.ColumnIndex, cell.RowIndex).Merge();
                }

            }

            res = range.Search(MyConstant.PrefixDate, MySearchOptions);

            foreach (var cell in res)
            {
                cell.SetValueFromText(cell.Value.ToString().Replace(MyConstant.PrefixDate, ""));
            }

            range.Worksheet.Workbook.EndUpdate();


            foreach (var cell in res)
            {
                cell.Formula = cell.Value.ToString().Replace(MyConstant.PrefixFormula, "");
            }
            try
            {
                ////          range.Worksheet.Workbook.History.IsEnabled = true;

            }
            catch (Exception) { }

        }

        public static void FormatRowsInRange(CellRange range, string typeRowHeading, string colNgay = null)
        {
            List<string> ValidMaCT = new List<string>();

            var ws = range.Worksheet;
            ws.OutlineOptions.SummaryRowsBelow = false;
            ////          ws.Workbook.History.IsEnabled = false;
            ws.Workbook.BeginUpdate();
            string OldDate = string.Empty;
            for (int i = range.TopRowIndex; i < range.BottomRowIndex + 1; i++)
            {
                Row crRow = ws.Rows[i];
                string typeRow = crRow[typeRowHeading].Value.ToString();

                //if (typesRowNeedGroup.Contains(typeRow))
                //{
                //    var nextInd = SpreadsheetHelper.FindNextGreaterTypeInd(range, i, rowChaHeading, typeRowHeading, typeRow);
                //    if (nextInd > i + 1)
                //        ws.Rows.Group(i, nextInd - 1, false);                
                //}
                switch (typeRow)
                {                 
                    case MyConstant.TYPEROW_CVCha:
                        crRow.Font.Color = Color.Black;
                        crRow.Font.Bold = false;
                        string Date = crRow[colNgay].Value.TextValue;
                        if (string.IsNullOrEmpty(Date))
                            continue;
                        if (Date != OldDate)
                        {
                            OldDate = Date;
                            int indRowCha = ws.Range.GetColumnIndexByName(colNgay);
                            var indsCon = ws.Range.FromLTRB(indRowCha, i, indRowCha, range.BottomRowIndex).Search(Date);
                            ws.Range[$"{colNgay}{i+1}:{colNgay}{i+indsCon.Count()}"].Merge();
                            //ws.Rows[i + indsCon.Count()].Visible = false;
                        }
                        break;
                    default:
                        break;
                }
            };

            ws.Workbook.EndUpdate();

            try
            {
                ////          ws.Workbook.History.IsEnabled = true;

            }
            catch (Exception) { }
        }



    }
}
