using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Common;
using Syncfusion.XlsIO;

namespace Ead.Summary.Generator
{
    class Program : BaseExcelGenerator
    {
        static void Main(string[] args)
        {
            var nums = new List<int> { 5, 2, 15 };
            var numsAsc = nums.OrderBy(n => n); // 2 5 15
            var numdsc = nums.OrderByDescending(n => n); // 15 5 2

            Console.WriteLine(Math.Round(1.02));
            Console.WriteLine(Math.Round(1.46));
            Console.WriteLine(Math.Round(1.6));
            CCRE7007();
        }

        private IDbConnection GetReportingConnection()
        {
             //var connStr = @"Data Source=SQLAU501MEL0057.globaltest.anz.com\CCRE_OLTP,49168;Initial Catalog=CCRE_Reporting;Integrated Security=SSPI;Connection Timeout=300;"; // PP - SQL 2016
             var connStr = @"Data Source=SQLAU101MEL0389.globaltest.anz.com\CCRE_OLTP,49168;Initial Catalog=CCRE_Reporting;Integrated Security=SSPI;Connection Timeout=300;";   // QA - SQL 2016
            return DatabaseConnectivity.GetConnection<SqlConnection>(connStr);
        }

        private static void CCRE7007()
        {
            // SACCR.GetEADSummaryReport 28,23 - QA
            // SACCR.GetEADSummaryReport 9,8 - PP
            EadReportInfo eadReportInfo = new EadReportInfo
            {
                FirstBatchRunID = 12,
                FirstBatchRunBusinessDate = "31-Mar-20",
                SecondBatchRunID = 16,
                SecondBatchRunBusinessDate = "31-Mar-20",
                GroupType = "Group",
            };
            var file = new Program().DownloadEadSummaryFile(eadReportInfo); // PP - SQL 2016
            //var file = new Program().DownloadEADDashBoardFile(9, 8, "Group"); // QA - SQL 2016
        }

        private byte[] DownloadEadSummaryFile(EadReportInfo eadReportInfo)
        {
            var eadSummaryReportData = GetEadSummaryReportData(eadReportInfo);
            return GenerateEadSummaryFile(eadSummaryReportData, eadReportInfo);
        }

        private byte[] GenerateEadSummaryFile(EadSummaryReportData eadSummaryReportData, EadReportInfo eadReportInfo)
        {
            const string EadSummaryFileName = "SACCR EAD Summary_{0}.xlsx";

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2010;
                application.StandardFont = "Verdana";
                application.StandardFontSize = 10;
                IWorkbook workbook = application.Workbooks.Create(1);

                #region Common Styles

                #region Table

                IStyle noDataHeaderStyle = workbook.Styles.Add("noDataHeaderStyle");
                noDataHeaderStyle.BeginUpdate();
                noDataHeaderStyle.Color = Color.DarkBlue;
                noDataHeaderStyle.Font.Bold = true;
                noDataHeaderStyle.Font.RGBColor = Color.White;
                noDataHeaderStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                noDataHeaderStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                noDataHeaderStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                noDataHeaderStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                noDataHeaderStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                noDataHeaderStyle.EndUpdate();

                IStyle tableFirstHeaderStyle = workbook.Styles.Add("tableFirstHeaderStyle");
                tableFirstHeaderStyle.BeginUpdate();
                tableFirstHeaderStyle.Color = Color.FromArgb(48, 51, 54);
                tableFirstHeaderStyle.Font.Bold = true;
                tableFirstHeaderStyle.Font.RGBColor = Color.White;
                tableFirstHeaderStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                tableFirstHeaderStyle.EndUpdate();

                IStyle tableSecondHeaderStyle = workbook.Styles.Add("tableSecondHeaderStyle");
                tableSecondHeaderStyle.BeginUpdate();
                tableSecondHeaderStyle.Color = Color.LightGray;
                tableSecondHeaderStyle.Font.Bold = true;
                tableSecondHeaderStyle.Font.RGBColor = Color.Black;
                tableSecondHeaderStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                tableSecondHeaderStyle.EndUpdate();

                #endregion

                #region Body style

                IStyle bodyStyle = workbook.Styles.Add("BodyStyle");
                bodyStyle.BeginUpdate();
                bodyStyle.Color = Color.FromArgb(239, 243, 247);
                bodyStyle.EndUpdate();

                IStyle cellGreyStyle = workbook.Styles.Add("cellGreyStyle");
                cellGreyStyle.BeginUpdate();
                cellGreyStyle.Color = Color.FromArgb(183, 187, 191);
                cellGreyStyle.EndUpdate();

                #endregion

                #endregion

                #region Summary        -- Tab 1

                IWorksheet summaryWorksheet = workbook.Worksheets[0];
                summaryWorksheet.Name = "Summary";
                summaryWorksheet.Zoom = 70;

                Log.InfoFormat("Writing data into worksheet1 started for : {0}", summaryWorksheet.Name);

                int tab1RowCounter = 0;

                if (eadSummaryReportData.CounterpartyTypes.Any())
                {
                    #region Table Header

                    summaryWorksheet.Range["F2:AS2"].Merge();
                    summaryWorksheet.Range["F2"].Text = "SA-CCR: EAD EXPLAIN SUMMARY";
                    summaryWorksheet.Range["F2"].CellStyle.Font.Bold = true;
                    summaryWorksheet.Range["F2"].CellStyle.Font.Size = 24;
                    summaryWorksheet.Range["F2"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                    summaryWorksheet.Range["F1"].ColumnWidth = 20;
                    summaryWorksheet.Range["G1:AS1"].ColumnWidth = 14;
                    summaryWorksheet.Range["AT1"].ColumnWidth = 35;

                    #endregion

                    #region Table1 - Counteparty Type

                    Log.Info("Writing data into worksheet1-Table1(Counteparty Type) started1");

                    tab1RowCounter = 5;

                    AddFirstTableHeader(summaryWorksheet, 1, "By Counteparty Type", ++tab1RowCounter, eadReportInfo, tableFirstHeaderStyle);
                    
                    AddSecondTableHeader(summaryWorksheet, 1, "Netting / CSA", ++tab1RowCounter, eadReportInfo, tableSecondHeaderStyle);

                    int _tab1RowCounter = tab1RowCounter;

                    ++tab1RowCounter;

                    var countepartyTypes = eadSummaryReportData.CounterpartyTypes.Where(cp => cp.TableSeq == 1);

                    #region NettingCollateral - TradeStatus

                    foreach (var cpGrpNettingCollateral in countepartyTypes.OrderBy(n => n.LegalDocStatusOrder).GroupBy(cp => cp.LegalDocStatus))
                    {
                        int __tab1RowCounter = tab1RowCounter;

                        foreach (var _cpNettingCollateral in cpGrpNettingCollateral.OrderBy(n => n.TradeStatusOrder))
                        {
                            summaryWorksheet.Range["F" + tab1RowCounter].Value = string.Format("{0} - {1}", cpGrpNettingCollateral.Key, _cpNettingCollateral.TradeStatus);

                            AddBodyRow(summaryWorksheet, tab1RowCounter, _cpNettingCollateral);

                            ++tab1RowCounter;
                        }

                        summaryWorksheet.Range[string.Format("B{0}:AT{1}", __tab1RowCounter, tab1RowCounter - 1)].CellStyle = bodyStyle;

                        #region Row Group

                        summaryWorksheet.Range[string.Format("A{0}:A{1}", __tab1RowCounter, tab1RowCounter - 1)].Group(ExcelGroupBy.ByRows, true);

                        #endregion

                        summaryWorksheet.Range["F" + tab1RowCounter].Value = cpGrpNettingCollateral.Key;

                        AddTotalRow(summaryWorksheet, tab1RowCounter, __tab1RowCounter);

                        ++tab1RowCounter;
                    }

                    #endregion

                    #region Total - TradeStatus

                    int ___tab1RowCounter = tab1RowCounter;
                    foreach (var cpGrpTradeStatus in countepartyTypes.GroupBy(cp => cp.TradeStatus))
                    {
                        summaryWorksheet.Range["F" + tab1RowCounter].Value = string.Format("{0} - {1}", "TOTAL", cpGrpTradeStatus.Key);

                        AddTotalRow(summaryWorksheet, tab1RowCounter, cpGrpTradeStatus);

                        ++tab1RowCounter;
                    }

                    summaryWorksheet.Range[string.Format("B{0}:AT{1}", ___tab1RowCounter, tab1RowCounter - 1)].CellStyle = bodyStyle;

                    #region Row Group

                    summaryWorksheet.Range[string.Format("A{0}:A{1}", ___tab1RowCounter, tab1RowCounter - 1)].Group(ExcelGroupBy.ByRows, true);

                    #endregion

                    summaryWorksheet.Range["F" + tab1RowCounter].Value = "TOTAL";

                    AddTotalRow(summaryWorksheet, tab1RowCounter, ___tab1RowCounter);

                    #endregion

                    #region Row/Column/Border Styles
                    
                    if (countepartyTypes.Any())
                    {
                        summaryWorksheet.Range[string.Format("F{0}:F{1}", _tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                        summaryWorksheet.Range[string.Format("G{0}:S{1}", _tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                        summaryWorksheet.Range[string.Format("T{0}:AF{1}", _tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                        summaryWorksheet.Range[string.Format("AG{0}:AS{1}", _tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                        summaryWorksheet.Range[string.Format("B{0}:AT{0}", _tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                        summaryWorksheet.Range[string.Format("B{0}:AT{1}", _tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);

                        summaryWorksheet.Range[string.Format("B{0}:AT{1}", ___tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                    }

                    #endregion

                    ++tab1RowCounter;

                    Log.Info("Writing data into worksheet1-Table1(Counteparty Type) completed.");

                    #endregion

                    #region Table2 - Product Type

                    Log.Info("Writing data into worksheet1-Table2(Product Type) started.");

                    tab1RowCounter = tab1RowCounter + 2;

                    AddFirstTableHeader(summaryWorksheet, 2, "By Product Type", tab1RowCounter, eadReportInfo, tableFirstHeaderStyle);

                    var productTypes = eadSummaryReportData.CounterpartyTypes.Where(cp => cp.TableSeq == 2);

                    _tab1RowCounter = tab1RowCounter;
                    ++tab1RowCounter;

                    foreach (var productType in productTypes.OrderBy(p => p.LegalDocStatusOrder).GroupBy(p => p.LegalDocStatus))
                    {
                        AddSecondTableHeader(summaryWorksheet, 2, productType.Key, tab1RowCounter, eadReportInfo, tableSecondHeaderStyle);
                        ++tab1RowCounter;

                        List<int> assetClassSumRows = new List<int>();

                        int __tab1RowCounter = tab1RowCounter;
                        foreach (var _productType in productType.OrderBy(p => p.AssetClassOrder).GroupBy(p => p.AssetClass))
                        {
                            var top10ProductTypes = _productType.OrderByDescending(p => Math.Abs(ToDouble(p.BATCH1_TotalRWA) - ToDouble(p.BATCH2_TotalRWA))).Take(10);

                            ___tab1RowCounter = tab1RowCounter;
                            foreach (var _top10ProductType in top10ProductTypes.OrderBy(p => p.AssetClassOrder))
                            {
                                summaryWorksheet.Range["F" + tab1RowCounter].Value = string.Format("{0} - {1}", _productType.Key, _top10ProductType.SingleFactor);

                                AddBodyRow(summaryWorksheet, tab1RowCounter, _top10ProductType);
                                
                                ++tab1RowCounter;
                            }

                            summaryWorksheet.Range[string.Format("B{0}:AT{1}", ___tab1RowCounter, tab1RowCounter - 1)].CellStyle = bodyStyle;

                            #region Row Group

                            summaryWorksheet.Range[string.Format("A{0}:A{1}", ___tab1RowCounter, tab1RowCounter - 1)].Group(ExcelGroupBy.ByRows, true);

                            #endregion

                            summaryWorksheet.Range["F" + tab1RowCounter].Value = _productType.Key;

                            AddTotalRow(summaryWorksheet, tab1RowCounter, _productType);

                            assetClassSumRows.Add(tab1RowCounter);

                            ++tab1RowCounter;
                        }

                        summaryWorksheet.Range["F" + tab1RowCounter].Value = "TOTAL";

                        AddTotalRow(summaryWorksheet, tab1RowCounter, assetClassSumRows);

                        #region Row/Column/Border Styles

                        if (productTypes.Any())
                        {
                            summaryWorksheet.Range[string.Format("F{0}:F{1}", _tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                            summaryWorksheet.Range[string.Format("G{0}:S{1}", _tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                            summaryWorksheet.Range[string.Format("T{0}:AF{1}", _tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                            summaryWorksheet.Range[string.Format("AG{0}:AS{1}", _tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                            summaryWorksheet.Range[string.Format("B{0}:AT{0}", tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);

                            summaryWorksheet.Range[string.Format("B{0}:AT{1}", _tab1RowCounter, tab1RowCounter)].BorderAround(ExcelLineStyle.Thin);
                        }

                        #endregion

                        ++tab1RowCounter;
                    }

                    Log.Info("Writing data into worksheet1-Table2(Product Type) completed.");

                    #endregion

                    #region Table3 - "Top 10 by RWA increase"

                    Log.Info("Writing data into worksheet1-Table3(Top 10 by RWA increase) started.");

                    tab1RowCounter = tab1RowCounter + 2;

                    AddFirstTableHeader(summaryWorksheet, 3, "Top 10 by RWA increase", tab1RowCounter, eadReportInfo, tableFirstHeaderStyle);

                    AddSecondTableHeader(summaryWorksheet, 3, "Razor ID", ++tab1RowCounter, eadReportInfo, tableSecondHeaderStyle);

                    _tab1RowCounter = tab1RowCounter;

                    ++tab1RowCounter;

                    var ctpyCodeTop10Increase = eadSummaryReportData.CounterpartyTypes
                        .Where(cp => cp.TableSeq == 3)
                        .GroupBy(c => c.PartyCd)
                        .OrderByDescending(p => p.Sum(t => ToDouble(t.BATCH1_TotalRWA)) - p.Sum(t => ToDouble(t.BATCH2_TotalRWA)))
                        .Take(10);

                    int index = 1;
                    foreach (var ctpyCode in ctpyCodeTop10Increase)
                    {
                        int __tab1RowCounter = tab1RowCounter;

                        foreach (var _ctpyCode in ctpyCode.GroupBy(c => c.TradeStatus))
                        {
                            foreach (var __ctpyCode in _ctpyCode)
                            {
                                summaryWorksheet.Range["A" + tab1RowCounter].Text = ctpyCode.Key;
                                summaryWorksheet.Range["B" + tab1RowCounter].Text = __ctpyCode.PartyName;
                                summaryWorksheet.Range["D" + tab1RowCounter].Text = __ctpyCode.CCR;
                                summaryWorksheet.Range["E" + tab1RowCounter].Text = __ctpyCode.BASEL_Treatment;
                                summaryWorksheet.Range["F" + tab1RowCounter].Text = string.Format("{0} - {1}", index, _ctpyCode.Key);

                                AddBodyRow(summaryWorksheet, tab1RowCounter, __ctpyCode);

                                ++tab1RowCounter;
                            }
                        }

                        summaryWorksheet.Range[string.Format("A{0}:AT{1}", __tab1RowCounter, tab1RowCounter - 1)].CellStyle = bodyStyle;
                        summaryWorksheet.Range[string.Format("C{0}:C{1}", __tab1RowCounter, tab1RowCounter - 1)].CellStyle = cellGreyStyle;

                        #region Row Group

                        summaryWorksheet.Range[string.Format("A{0}:A{1}", __tab1RowCounter, tab1RowCounter - 1)].Group(ExcelGroupBy.ByRows, true);

                        #endregion

                        summaryWorksheet.Range["A" + tab1RowCounter].Text = ctpyCode.Key;
                        summaryWorksheet.Range["B" + tab1RowCounter].Text = ctpyCode.FirstOrDefault().PartyName;
                        summaryWorksheet.Range["C" + tab1RowCounter].Text = ctpyCode.FirstOrDefault().LegalDocStatus;
                        summaryWorksheet.Range["D" + tab1RowCounter].Text = ctpyCode.FirstOrDefault().CCR;
                        summaryWorksheet.Range["E" + tab1RowCounter].Text = ctpyCode.FirstOrDefault().BASEL_Treatment;
                        summaryWorksheet.Range["F" + tab1RowCounter].Value = index.ToString();

                        AddTotalRow(summaryWorksheet, tab1RowCounter, __tab1RowCounter);

                        index++;
                        ++tab1RowCounter;
                    }

                    Log.Info("Writing data into worksheet1-Table3(Top 10 by RWA increase) completed.");

                    #endregion

                    #region Table4 - "Top 10 by RWA decrease"

                    Log.Info("Writing data into worksheet1-Table4(Top 10 by RWA decrease) started.");

                    tab1RowCounter = tab1RowCounter + 2;

                    AddFirstTableHeader(summaryWorksheet, 4, "Top 10 by RWA decrease", tab1RowCounter, eadReportInfo, tableFirstHeaderStyle);

                    AddSecondTableHeader(summaryWorksheet, 4, "Razor ID", ++tab1RowCounter, eadReportInfo, tableSecondHeaderStyle);

                    _tab1RowCounter = tab1RowCounter;

                    ++tab1RowCounter;

                    var ctpyCodeTop10Decrease = eadSummaryReportData.CounterpartyTypes
                        .Where(cp => cp.TableSeq == 3)
                        .GroupBy(c => c.PartyCd)
                        .OrderBy(p => p.Sum(t => ToDouble(t.BATCH1_TotalRWA)) - p.Sum(t => ToDouble(t.BATCH2_TotalRWA)))
                        .Take(10);

                    index = 1;
                    foreach (var ctpyCode in ctpyCodeTop10Decrease)
                    {
                        int __tab1RowCounter = tab1RowCounter;

                        foreach (var _ctpyCode in ctpyCode.GroupBy(c => c.TradeStatus))
                        {
                            foreach (var __ctpyCode in _ctpyCode)
                            {
                                summaryWorksheet.Range["A" + tab1RowCounter].Text = ctpyCode.Key;
                                summaryWorksheet.Range["B" + tab1RowCounter].Text = __ctpyCode.PartyName;
                                summaryWorksheet.Range["D" + tab1RowCounter].Text = __ctpyCode.CCR;
                                summaryWorksheet.Range["E" + tab1RowCounter].Text = __ctpyCode.BASEL_Treatment;
                                summaryWorksheet.Range["F" + tab1RowCounter].Value = string.Format("{0} - {1}", index, _ctpyCode.Key);

                                AddBodyRow(summaryWorksheet, tab1RowCounter, __ctpyCode);

                                ++tab1RowCounter;
                            }
                        }

                        summaryWorksheet.Range[string.Format("A{0}:AT{1}", __tab1RowCounter, tab1RowCounter - 1)].CellStyle = bodyStyle;
                        summaryWorksheet.Range[string.Format("C{0}:C{1}", __tab1RowCounter, tab1RowCounter - 1)].CellStyle = cellGreyStyle;

                        #region Row Group

                        summaryWorksheet.Range[string.Format("A{0}:A{1}", __tab1RowCounter, tab1RowCounter - 1)].Group(ExcelGroupBy.ByRows, true);

                        #endregion

                        summaryWorksheet.Range["A" + tab1RowCounter].Text = ctpyCode.Key;
                        summaryWorksheet.Range["B" + tab1RowCounter].Text = ctpyCode.FirstOrDefault().PartyName;
                        summaryWorksheet.Range["C" + tab1RowCounter].Text = ctpyCode.FirstOrDefault().LegalDocStatus;
                        summaryWorksheet.Range["D" + tab1RowCounter].Text = ctpyCode.FirstOrDefault().CCR;
                        summaryWorksheet.Range["E" + tab1RowCounter].Text = ctpyCode.FirstOrDefault().BASEL_Treatment;
                        summaryWorksheet.Range["F" + tab1RowCounter].Value = index.ToString();

                        AddTotalRow(summaryWorksheet, tab1RowCounter, __tab1RowCounter);

                        index++;
                        ++tab1RowCounter;
                    }

                    Log.Info("Writing data into worksheet1-Table4(Top 10 by RWA decrease) completed.");

                    #endregion

                    #region Column Group

                    summaryWorksheet.Range["B1:E1"].Group(ExcelGroupBy.ByColumns, true);
                    summaryWorksheet.Range["G1:H1"].Group(ExcelGroupBy.ByColumns, true);
                    summaryWorksheet.Range["J1:N1"].Group(ExcelGroupBy.ByColumns, true);
                    summaryWorksheet.Range["Q1:R1"].Group(ExcelGroupBy.ByColumns, true);
                    summaryWorksheet.Range["T1:U1"].Group(ExcelGroupBy.ByColumns, true);
                    summaryWorksheet.Range["W1:AA1"].Group(ExcelGroupBy.ByColumns, true);
                    summaryWorksheet.Range["AD1:AE1"].Group(ExcelGroupBy.ByColumns, true);
                    summaryWorksheet.Range["AG1:AH1"].Group(ExcelGroupBy.ByColumns, true);
                    summaryWorksheet.Range["AJ1:AN1"].Group(ExcelGroupBy.ByColumns, true);
                    summaryWorksheet.Range["AQ1:AR1"].Group(ExcelGroupBy.ByColumns, true);

                    #endregion

                    #region Formatting for -ve values

                    IConditionalFormats condition = summaryWorksheet.Range["G8:AS" + tab1RowCounter].ConditionalFormats;
                    IConditionalFormat negativeValue = condition.AddCondition();
                    negativeValue.FormatType = ExcelCFType.CellValue;
                    negativeValue.Operator = ExcelComparisonOperator.Less;
                    negativeValue.FirstFormula = "0";
                    negativeValue.FontColor = ExcelKnownColors.Red;

                    #endregion

                    #region Amount Format

                    summaryWorksheet.Range["G8:AS" + tab1RowCounter].NumberFormat = "###,##0";

                    #endregion
                }
                else 
                {
                    summaryWorksheet.Range["B2:F2"].Merge();
                    summaryWorksheet.Range["B2"].Text = NoDataMessage;
                    summaryWorksheet.Range["B2:F2"].CellStyle = noDataHeaderStyle;
                }

                #endregion

                string path = @"C:\temp\7007";
                workbook.SaveAs(Path.Combine(path, string.Format(EadSummaryFileName, DateTime.Now.Ticks)));

                workbook.Close();
                excelEngine.Dispose();
            }

            return null;
        }

        #region Private

        #region Excel Table Headers

        private static void AddFirstTableHeader(IWorksheet worksheet, int tableSeq, string typeName, int tabRowCounter, EadReportInfo eadDashboardReportInfo, IStyle style)
        {
            switch (tableSeq)
            {
                case 1:
                case 2:
                case 3:
                case 4:
                    worksheet.Range["F" + tabRowCounter].Text = typeName;
                    worksheet.Range["G" + tabRowCounter].Text = string.Format("Difference (Batch({0})_{1} - Batch({2})_{3}", eadDashboardReportInfo.FirstBatchRunID, FormatBusinessDate(eadDashboardReportInfo.FirstBatchRunBusinessDate), eadDashboardReportInfo.SecondBatchRunID, FormatBusinessDate(eadDashboardReportInfo.SecondBatchRunBusinessDate));
                    worksheet.Range[string.Format("G{0}:S{0}", tabRowCounter)].Merge();
                    worksheet.Range["T" + tabRowCounter].Text = string.Format("Batch({0})_{1}", eadDashboardReportInfo.FirstBatchRunID, FormatBusinessDate(eadDashboardReportInfo.FirstBatchRunBusinessDate));
                    worksheet.Range[string.Format("T{0}:AF{0}", tabRowCounter)].Merge();
                    worksheet.Range["AG" + tabRowCounter].Text = string.Format("Batch({0})_{1}", eadDashboardReportInfo.SecondBatchRunID, FormatBusinessDate(eadDashboardReportInfo.SecondBatchRunBusinessDate));
                    worksheet.Range[string.Format("AG{0}:AS{0}", tabRowCounter)].Merge();
                    worksheet.Range["AT" + tabRowCounter].Text = "Comments";
                    break;
                default:
                    break;
            }

            worksheet.Range[string.Format("B{0}:AT{0}", tabRowCounter)].CellStyle = style;
            worksheet.Range[string.Format("B{0}:AT{0}", tabRowCounter)].RowHeight = 25;
            worksheet.Range[string.Format("B{0}:AT{0}", tabRowCounter)].VerticalAlignment = ExcelVAlign.VAlignCenter;

            worksheet.Range["F" + tabRowCounter].WrapText = true;

            worksheet.Range[string.Format("G{0}:S{0}", tabRowCounter)].Merge();
            worksheet.Range[string.Format("T{0}:AF{0}", tabRowCounter)].Merge();
            worksheet.Range[string.Format("AG{0}:AS{0}", tabRowCounter)].Merge();

            worksheet.Range["AT" + tabRowCounter].WrapText = true;

        }

        private static void AddSecondTableHeader(IWorksheet worksheet, int tableSeq, string nettingCollateralTypeName, int tabRowCounter, EadReportInfo eadDashboardReportInfo, IStyle style)
        {
            switch (tableSeq)
            {
                case 1:
                case 2:
                    worksheet.Range[string.Format("B{0}:AT{0}", tabRowCounter)].CellStyle = style;
                    break;

                case 3:
                case 4:
                    worksheet.Range["A" + tabRowCounter].Text = "Ctpy Code";
                    worksheet.Range["B" + tabRowCounter].Text = "Ctpy Name";
                    worksheet.Range["C" + tabRowCounter].Text = "Netting";
                    worksheet.Range["D" + tabRowCounter].Text = "CCR";
                    worksheet.Range["E" + tabRowCounter].Text = "Basel treatment";

                    worksheet.Range[string.Format("A{0}:AT{0}", tabRowCounter)].CellStyle = style;
                    break;
            }

            worksheet.Range["F" + tabRowCounter].Text = nettingCollateralTypeName;

            // B1 - B2
            worksheet.Range["G" + tabRowCounter].Text = "MtM";
            worksheet.Range["H" + tabRowCounter].Text = "C";
            worksheet.Range["I" + tabRowCounter].Text = "RC";
            worksheet.Range["J" + tabRowCounter].Text = "PFE IR";
            worksheet.Range["K" + tabRowCounter].Text = "PFE FX";
            worksheet.Range["L" + tabRowCounter].Text = "PFE Commo";
            worksheet.Range["M" + tabRowCounter].Text = "PFE Cr";
            worksheet.Range["N" + tabRowCounter].Text = "PFE Eq";
            worksheet.Range["O" + tabRowCounter].Text = "PFE";
            worksheet.Range["P" + tabRowCounter].Text = "EAD";
            worksheet.Range["Q" + tabRowCounter].Text = "RWA";
            worksheet.Range["R" + tabRowCounter].Text = "CVA_RWA";
            worksheet.Range["S" + tabRowCounter].Text = "Total RWA";

            // B1
            worksheet.Range["T" + tabRowCounter].Text = "MtM";
            worksheet.Range["U" + tabRowCounter].Text = "C";
            worksheet.Range["V" + tabRowCounter].Text = "RC";
            worksheet.Range["W" + tabRowCounter].Text = "PFE IR";
            worksheet.Range["X" + tabRowCounter].Text = "PFE FX";
            worksheet.Range["Y" + tabRowCounter].Text = "PFE COMM";
            worksheet.Range["Z" + tabRowCounter].Text = "PFE CR";
            worksheet.Range["AA" + tabRowCounter].Text = "PFE EQ";
            worksheet.Range["AB" + tabRowCounter].Text = "PFE";
            worksheet.Range["AC" + tabRowCounter].Text = "EAD";
            worksheet.Range["AD" + tabRowCounter].Text = "RWA";
            worksheet.Range["AE" + tabRowCounter].Text = "CVA_RWA";
            worksheet.Range["AF" + tabRowCounter].Text = "Total RWA";

            // B2
            worksheet.Range["AG" + tabRowCounter].Text = "MtM";
            worksheet.Range["AH" + tabRowCounter].Text = "C";
            worksheet.Range["AI" + tabRowCounter].Text = "RC";
            worksheet.Range["AJ" + tabRowCounter].Text = "PFE IR";
            worksheet.Range["AK" + tabRowCounter].Text = "PFE FX";
            worksheet.Range["AL" + tabRowCounter].Text = "PFE COMM";
            worksheet.Range["AM" + tabRowCounter].Text = "PFE CR";
            worksheet.Range["AN" + tabRowCounter].Text = "PFE EQ";
            worksheet.Range["AO" + tabRowCounter].Text = "PFE";
            worksheet.Range["AP" + tabRowCounter].Text = "EAD";
            worksheet.Range["AQ" + tabRowCounter].Text = "RWA";
            worksheet.Range["AR" + tabRowCounter].Text = "CVA_RWA";
            worksheet.Range["AS" + tabRowCounter].Text = "Total RWA";

            worksheet.Range["AT" + tabRowCounter].Text = "Comments";
        }
        
        #endregion

        #region Table Row and Total

        private static void AddBodyRow(IWorksheet worksheet, int tab1RowCounter, EadSummaryByCounterpartyType eadSummaryByCounterpartyType)
        {
            #region B1-B2

            worksheet.Range["G" + tab1RowCounter].Formula = string.Format("T{0}-AG{0}", tab1RowCounter);
            worksheet.Range["H" + tab1RowCounter].Formula = string.Format("U{0}-AH{0}", tab1RowCounter);
            worksheet.Range["I" + tab1RowCounter].Formula = string.Format("V{0}-AI{0}", tab1RowCounter);
            worksheet.Range["J" + tab1RowCounter].Formula = string.Format("W{0}-AJ{0}", tab1RowCounter);
            worksheet.Range["K" + tab1RowCounter].Formula = string.Format("X{0}-AK{0}", tab1RowCounter);
            worksheet.Range["L" + tab1RowCounter].Formula = string.Format("Y{0}-AL{0}", tab1RowCounter);
            worksheet.Range["M" + tab1RowCounter].Formula = string.Format("Z{0}-AM{0}", tab1RowCounter);
            worksheet.Range["N" + tab1RowCounter].Formula = string.Format("AA{0}-AN{0}", tab1RowCounter);
            worksheet.Range["O" + tab1RowCounter].Formula = string.Format("AB{0}-AO{0}", tab1RowCounter);
            worksheet.Range["P" + tab1RowCounter].Formula = string.Format("AC{0}-AP{0}", tab1RowCounter);
            worksheet.Range["Q" + tab1RowCounter].Formula = string.Format("AD{0}-AQ{0}", tab1RowCounter);
            worksheet.Range["R" + tab1RowCounter].Formula = string.Format("AE{0}-AR{0}", tab1RowCounter);
            worksheet.Range["S" + tab1RowCounter].Formula = string.Format("AF{0}-AS{0}", tab1RowCounter);

            #endregion

            #region B1

            worksheet.Range["T" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_MTM, true);
            worksheet.Range["U" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_C, true);
            worksheet.Range["V" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_RC, true);
            worksheet.Range["W" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_PFE_IR, true);
            worksheet.Range["X" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_PFE_FX, true);
            worksheet.Range["Y" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_PFE_COMM, true);
            worksheet.Range["Z" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_PFE_CR, true);
            worksheet.Range["AA" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_PFE_EQ, true);
            worksheet.Range["AB" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_PFE, true);
            worksheet.Range["AC" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_EAD, true);
            worksheet.Range["AD" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_RWA, true);
            worksheet.Range["AE" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_CVA_RWA, true);
            worksheet.Range["AF" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH1_TotalRWA, true);

            #endregion

            #region B2

            worksheet.Range["AG" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_MTM, true);
            worksheet.Range["AH" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_C, true);
            worksheet.Range["AI" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_RC, true);
            worksheet.Range["AJ" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_PFE_IR, true);
            worksheet.Range["AK" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_PFE_FX, true);
            worksheet.Range["AL" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_PFE_COMM, true);
            worksheet.Range["AM" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_PFE_CR, true);
            worksheet.Range["AN" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_PFE_Eq, true);
            worksheet.Range["AO" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_PFE, true);
            worksheet.Range["AP" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_EAD, true);
            worksheet.Range["AQ" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_RWA, true);
            worksheet.Range["AR" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_CVA_RWA, true);
            worksheet.Range["AS" + tab1RowCounter].Value = FormatDoubleValueForExcel(eadSummaryByCounterpartyType.BATCH2_TotalRWA, true);

            #endregion
        }

        private static void AddTotalRow(IWorksheet worksheet, int tab1RowCounter, int tab1ChildRowCounter)
        {
            #region B1-B2

            worksheet.Range["G" + tab1RowCounter].Formula = string.Format("=Sum(G{0}:G{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["H" + tab1RowCounter].Formula = string.Format("=Sum(H{0}:H{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["I" + tab1RowCounter].Formula = string.Format("=Sum(I{0}:I{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["J" + tab1RowCounter].Formula = string.Format("=Sum(J{0}:J{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["K" + tab1RowCounter].Formula = string.Format("=Sum(K{0}:K{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["L" + tab1RowCounter].Formula = string.Format("=Sum(L{0}:L{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["M" + tab1RowCounter].Formula = string.Format("=Sum(M{0}:M{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["N" + tab1RowCounter].Formula = string.Format("=Sum(N{0}:N{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["O" + tab1RowCounter].Formula = string.Format("=Sum(O{0}:O{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["P" + tab1RowCounter].Formula = string.Format("=Sum(P{0}:P{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["Q" + tab1RowCounter].Formula = string.Format("=Sum(Q{0}:Q{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["R" + tab1RowCounter].Formula = string.Format("=Sum(R{0}:R{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["S" + tab1RowCounter].Formula = string.Format("=Sum(S{0}:S{1})", tab1ChildRowCounter, tab1RowCounter - 1);

            #endregion

            #region B1

            worksheet.Range["T" + tab1RowCounter].Formula = string.Format("=Sum(T{0}:T{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["U" + tab1RowCounter].Formula = string.Format("=Sum(U{0}:U{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["V" + tab1RowCounter].Formula = string.Format("=Sum(V{0}:V{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["W" + tab1RowCounter].Formula = string.Format("=Sum(W{0}:W{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["X" + tab1RowCounter].Formula = string.Format("=Sum(X{0}:X{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["Y" + tab1RowCounter].Formula = string.Format("=Sum(Y{0}:Y{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["Z" + tab1RowCounter].Formula = string.Format("=Sum(Z{0}:Z{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AA" + tab1RowCounter].Formula = string.Format("=Sum(AA{0}:AA{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AB" + tab1RowCounter].Formula = string.Format("=Sum(AB{0}:AB{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AC" + tab1RowCounter].Formula = string.Format("=Sum(AC{0}:AC{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AD" + tab1RowCounter].Formula = string.Format("=Sum(AD{0}:AD{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AE" + tab1RowCounter].Formula = string.Format("=Sum(AE{0}:AE{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AF" + tab1RowCounter].Formula = string.Format("=Sum(AF{0}:AF{1})", tab1ChildRowCounter, tab1RowCounter - 1);

            #endregion

            #region B2

            worksheet.Range["AG" + tab1RowCounter].Formula = string.Format("=Sum(AG{0}:AG{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AH" + tab1RowCounter].Formula = string.Format("=Sum(AH{0}:AH{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AI" + tab1RowCounter].Formula = string.Format("=Sum(AI{0}:AI{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AJ" + tab1RowCounter].Formula = string.Format("=Sum(AJ{0}:AJ{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AK" + tab1RowCounter].Formula = string.Format("=Sum(AK{0}:AK{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AL" + tab1RowCounter].Formula = string.Format("=Sum(AL{0}:AL{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AM" + tab1RowCounter].Formula = string.Format("=Sum(AM{0}:AM{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AN" + tab1RowCounter].Formula = string.Format("=Sum(AN{0}:AN{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AO" + tab1RowCounter].Formula = string.Format("=Sum(AO{0}:AO{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AP" + tab1RowCounter].Formula = string.Format("=Sum(AP{0}:AP{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AQ" + tab1RowCounter].Formula = string.Format("=Sum(AQ{0}:AQ{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AR" + tab1RowCounter].Formula = string.Format("=Sum(AR{0}:AR{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            worksheet.Range["AS" + tab1RowCounter].Formula = string.Format("=Sum(AS{0}:AS{1})", tab1ChildRowCounter, tab1RowCounter - 1);
            
            #endregion
        }

        private static void AddTotalRow(IWorksheet summaryWorksheet, int tab1RowCounter, IGrouping<string, EadSummaryByCounterpartyType> productType)
        {
            #region B1-B2

            summaryWorksheet.Range["G" + tab1RowCounter].Formula = string.Format("T{0}-AG{0}", tab1RowCounter);
            summaryWorksheet.Range["H" + tab1RowCounter].Formula = string.Format("U{0}-AH{0}", tab1RowCounter);
            summaryWorksheet.Range["I" + tab1RowCounter].Formula = string.Format("V{0}-AI{0}", tab1RowCounter);
            summaryWorksheet.Range["J" + tab1RowCounter].Formula = string.Format("W{0}-AJ{0}", tab1RowCounter);
            summaryWorksheet.Range["K" + tab1RowCounter].Formula = string.Format("X{0}-AK{0}", tab1RowCounter);
            summaryWorksheet.Range["L" + tab1RowCounter].Formula = string.Format("Y{0}-AL{0}", tab1RowCounter);
            summaryWorksheet.Range["M" + tab1RowCounter].Formula = string.Format("Z{0}-AM{0}", tab1RowCounter);
            summaryWorksheet.Range["N" + tab1RowCounter].Formula = string.Format("AA{0}-AN{0}", tab1RowCounter);
            summaryWorksheet.Range["O" + tab1RowCounter].Formula = string.Format("AB{0}-AO{0}", tab1RowCounter);
            summaryWorksheet.Range["P" + tab1RowCounter].Formula = string.Format("AC{0}-AP{0}", tab1RowCounter);
            summaryWorksheet.Range["Q" + tab1RowCounter].Formula = string.Format("AD{0}-AQ{0}", tab1RowCounter);
            summaryWorksheet.Range["R" + tab1RowCounter].Formula = string.Format("AE{0}-AR{0}", tab1RowCounter);
            summaryWorksheet.Range["S" + tab1RowCounter].Formula = string.Format("AF{0}-AS{0}", tab1RowCounter);

            #endregion

            #region B1

            summaryWorksheet.Range["T" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_MTM)).ToString(), true);
            summaryWorksheet.Range["U" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_C)).ToString(), true);
            summaryWorksheet.Range["V" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_RC)).ToString(), true);
            summaryWorksheet.Range["W" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_PFE_IR)).ToString(), true);
            summaryWorksheet.Range["X" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_PFE_FX)).ToString(), true);
            summaryWorksheet.Range["Y" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_PFE_COMM)).ToString(), true);
            summaryWorksheet.Range["Z" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_PFE_CR)).ToString(), true);
            summaryWorksheet.Range["AA" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_PFE_EQ)).ToString(), true);
            summaryWorksheet.Range["AB" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_PFE)).ToString(), true);
            summaryWorksheet.Range["AC" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_EAD)).ToString(), true);
            summaryWorksheet.Range["AD" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_RWA)).ToString(), true);
            summaryWorksheet.Range["AE" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_CVA_RWA)).ToString(), true);
            summaryWorksheet.Range["AF" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH1_TotalRWA)).ToString(), true);

            #endregion

            #region B2

            summaryWorksheet.Range["AG" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_MTM)).ToString(), true);
            summaryWorksheet.Range["AH" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_C)).ToString(), true);
            summaryWorksheet.Range["AI" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_RC)).ToString(), true);
            summaryWorksheet.Range["AJ" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_PFE_IR)).ToString(), true);
            summaryWorksheet.Range["AK" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_PFE_FX)).ToString(), true);
            summaryWorksheet.Range["AL" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_PFE_COMM)).ToString(), true);
            summaryWorksheet.Range["AM" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_PFE_CR)).ToString(), true);
            summaryWorksheet.Range["AN" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_PFE_Eq)).ToString(), true);
            summaryWorksheet.Range["AO" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_PFE)).ToString(), true);
            summaryWorksheet.Range["AP" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_EAD)).ToString(), true);
            summaryWorksheet.Range["AQ" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_RWA)).ToString(), true);
            summaryWorksheet.Range["AR" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_CVA_RWA)).ToString(), true);
            summaryWorksheet.Range["AS" + tab1RowCounter].Value = FormatDoubleValueForExcel(productType.Sum(ts => ToDouble(ts.BATCH2_TotalRWA)).ToString(), true);

            #endregion
        }

        /// <summary>
        /// Table 2 - Total Row
        /// </summary>
        private static void AddTotalRow(IWorksheet worksheet, int tab1RowCounter, List<int> rowCounters)
        {
            #region B1-B2

            worksheet.Range["G" + tab1RowCounter].Formula = string.Format("T{0}-AG{0}", tab1RowCounter);
            worksheet.Range["H" + tab1RowCounter].Formula = string.Format("U{0}-AH{0}", tab1RowCounter);
            worksheet.Range["I" + tab1RowCounter].Formula = string.Format("V{0}-AI{0}", tab1RowCounter);
            worksheet.Range["J" + tab1RowCounter].Formula = string.Format("W{0}-AJ{0}", tab1RowCounter);
            worksheet.Range["K" + tab1RowCounter].Formula = string.Format("X{0}-AK{0}", tab1RowCounter);
            worksheet.Range["L" + tab1RowCounter].Formula = string.Format("Y{0}-AL{0}", tab1RowCounter);
            worksheet.Range["M" + tab1RowCounter].Formula = string.Format("Z{0}-AM{0}", tab1RowCounter);
            worksheet.Range["N" + tab1RowCounter].Formula = string.Format("AA{0}-AN{0}", tab1RowCounter);
            worksheet.Range["O" + tab1RowCounter].Formula = string.Format("AB{0}-AO{0}", tab1RowCounter);
            worksheet.Range["P" + tab1RowCounter].Formula = string.Format("AC{0}-AP{0}", tab1RowCounter);
            worksheet.Range["Q" + tab1RowCounter].Formula = string.Format("AD{0}-AQ{0}", tab1RowCounter);
            worksheet.Range["R" + tab1RowCounter].Formula = string.Format("AE{0}-AR{0}", tab1RowCounter);
            worksheet.Range["S" + tab1RowCounter].Formula = string.Format("AF{0}-AS{0}", tab1RowCounter);

            #endregion

            #region B1

            worksheet.Range["T" + tab1RowCounter].Value = GetCellSumFormula("T", rowCounters);
            worksheet.Range["U" + tab1RowCounter].Value = GetCellSumFormula("U", rowCounters);
            worksheet.Range["V" + tab1RowCounter].Value = GetCellSumFormula("V", rowCounters);
            worksheet.Range["W" + tab1RowCounter].Value = GetCellSumFormula("W", rowCounters);
            worksheet.Range["X" + tab1RowCounter].Value = GetCellSumFormula("X", rowCounters);
            worksheet.Range["Y" + tab1RowCounter].Value = GetCellSumFormula("Y", rowCounters);
            worksheet.Range["Z" + tab1RowCounter].Value = GetCellSumFormula("Z", rowCounters);
            worksheet.Range["AA" + tab1RowCounter].Value = GetCellSumFormula("AA", rowCounters);
            worksheet.Range["AB" + tab1RowCounter].Value = GetCellSumFormula("AB", rowCounters);
            worksheet.Range["AC" + tab1RowCounter].Value = GetCellSumFormula("AC", rowCounters);
            worksheet.Range["AD" + tab1RowCounter].Value = GetCellSumFormula("AD", rowCounters);
            worksheet.Range["AE" + tab1RowCounter].Value = GetCellSumFormula("AE", rowCounters);
            worksheet.Range["AF" + tab1RowCounter].Value = GetCellSumFormula("AF", rowCounters);

            #endregion

            #region B2

            worksheet.Range["AG" + tab1RowCounter].Value = GetCellSumFormula("AG", rowCounters);
            worksheet.Range["AH" + tab1RowCounter].Value = GetCellSumFormula("AH", rowCounters);
            worksheet.Range["AI" + tab1RowCounter].Value = GetCellSumFormula("AI", rowCounters);
            worksheet.Range["AJ" + tab1RowCounter].Value = GetCellSumFormula("AJ", rowCounters);
            worksheet.Range["AK" + tab1RowCounter].Value = GetCellSumFormula("AK", rowCounters);
            worksheet.Range["AL" + tab1RowCounter].Value = GetCellSumFormula("AL", rowCounters);
            worksheet.Range["AM" + tab1RowCounter].Value = GetCellSumFormula("AM", rowCounters);
            worksheet.Range["AN" + tab1RowCounter].Value = GetCellSumFormula("AN", rowCounters);
            worksheet.Range["AO" + tab1RowCounter].Value = GetCellSumFormula("AO", rowCounters);
            worksheet.Range["AP" + tab1RowCounter].Value = GetCellSumFormula("AP", rowCounters);
            worksheet.Range["AQ" + tab1RowCounter].Value = GetCellSumFormula("AQ", rowCounters);
            worksheet.Range["AR" + tab1RowCounter].Value = GetCellSumFormula("AR", rowCounters);
            worksheet.Range["AS" + tab1RowCounter].Value = GetCellSumFormula("AS", rowCounters);

            #endregion
        }

        private static string GetCellSumFormula(string columnName, List<int> rowCounters)
        {
            var cells = string.Join(",", rowCounters.Select(counter => columnName + counter));
            return "=SUM(" + cells + ")";
        }

        #endregion

        #endregion

        #region DB Related - Get

        private EadSummaryReportData GetEadSummaryReportData(EadReportInfo eadReportInfo)
        {
            List<EadSummaryByCounterpartyType> _eadSummaryData = null;

            Parallel.Invoke(() => { _eadSummaryData = GetEadSummaryData(eadReportInfo.FirstBatchRunID, eadReportInfo.SecondBatchRunID, eadReportInfo.GroupType); });

            return new EadSummaryReportData(_eadSummaryData);
        }

        private List<EadSummaryByCounterpartyType> GetEadSummaryData(long firstBatchRunID, long secondBatchRunID, string groupType)
        {
            using (var cn = GetReportingConnection())
            {
                var cm = cn.CreateCommand();
                cm.CommandText = "[SACCR].[GetEADSummaryReport]";
                cm.CommandType = CommandType.StoredProcedure;
                cm.CommandTimeout = 300;

                var param1 = new SqlParameter("@BatchRunId1", firstBatchRunID);
                param1.Direction = ParameterDirection.Input;
                param1.DbType = DbType.Int64;

                var param2 = new SqlParameter("@BatchRunId2", secondBatchRunID);
                param2.Direction = ParameterDirection.Input;
                param2.DbType = DbType.Int64;

                var param3 = new SqlParameter("@GroupType", groupType);
                param3.Direction = ParameterDirection.Input;
                param3.DbType = DbType.String;

                cm.Parameters.Add(param1);
                cm.Parameters.Add(param2);
                cm.Parameters.Add(param3);

                SqlDataAdapter adapter = new SqlDataAdapter(cm as SqlCommand);

                DataSet ds = new DataSet();
                adapter.Fill(ds);
                cn.Close();

                var res = new List<EadSummaryByCounterpartyType>();
                int tableSeq = 0;

                foreach (DataTable table in ds.Tables)
                {
                    ++tableSeq;
                    foreach (DataRow row in table.Rows)
                    {
                        EadSummaryByCounterpartyType eadSummaryByCounterPartyType = new EadSummaryByCounterpartyType();
                        eadSummaryByCounterPartyType.TableSeq = tableSeq;

                        switch (tableSeq)
                        {
                            case 1:
                                eadSummaryByCounterPartyType.TradeStatus = Convert.ToString(row["TradeStatus"]);
                                eadSummaryByCounterPartyType.TradeStatusOrder = Convert.ToInt32(row["TradeStatus_Order"]);

                                eadSummaryByCounterPartyType.LegalDocStatus = Convert.ToString(row["LegalDocStatus"]);
                                eadSummaryByCounterPartyType.LegalDocStatusOrder = Convert.ToInt32(row["LegalDocStatus_Order"]);

                                break;
                            case 2:
                                eadSummaryByCounterPartyType.AssetClass = Convert.ToString(row["AssetClass"]);
                                eadSummaryByCounterPartyType.AssetClassOrder = Convert.ToInt32(row["AssetClass_Order"]);

                                eadSummaryByCounterPartyType.LegalDocStatus = Convert.ToString(row["LegalDocStatus"]);
                                eadSummaryByCounterPartyType.LegalDocStatusOrder = Convert.ToInt32(row["LegalDocStatus_Order"]);

                                eadSummaryByCounterPartyType.SingleFactor = Convert.ToString(row["SingleFactor"]);

                                break;

                            case 3:
                            case 4:
                                eadSummaryByCounterPartyType.PartyCd = Convert.ToString(row["PartyCd"]);
                                eadSummaryByCounterPartyType.PartyName = Convert.ToString(row["PartyName"]);
                                eadSummaryByCounterPartyType.LegalDocStatus = Convert.ToString(row["LegalDocStatus"]);
                                eadSummaryByCounterPartyType.CCR = Convert.ToString(row["CCR"]);
                                eadSummaryByCounterPartyType.BASEL_Treatment = Convert.ToString(row["BASEL_Treatment"]);
                                eadSummaryByCounterPartyType.TradeStatus = Convert.ToString(row["TradeStatus"]);

                                break;
                        }

                        eadSummaryByCounterPartyType.BATCH1_MTM = Convert.ToString(row["BATCH1_MTM"]);
                        eadSummaryByCounterPartyType.BATCH1_C = Convert.ToString(row["BATCH1_C"]);
                        eadSummaryByCounterPartyType.BATCH1_RC = Convert.ToString(row["BATCH1_RC"]);
                        eadSummaryByCounterPartyType.BATCH1_PFE_IR = Convert.ToString(row["BATCH1_PFE_IR"]);
                        eadSummaryByCounterPartyType.BATCH1_PFE_FX = Convert.ToString(row["BATCH1_PFE_FX"]);
                        eadSummaryByCounterPartyType.BATCH1_PFE_COMM = Convert.ToString(row["BATCH1_PFE_COMM"]);
                        eadSummaryByCounterPartyType.BATCH1_PFE_CR = Convert.ToString(row["BATCH1_PFE_CR"]);
                        eadSummaryByCounterPartyType.BATCH1_PFE_EQ = Convert.ToString(row["BATCH1_PFE_EQ"]);
                        eadSummaryByCounterPartyType.BATCH1_PFE = Convert.ToString(row["BATCH1_PFE"]);
                        eadSummaryByCounterPartyType.BATCH1_EAD = Convert.ToString(row["BATCH1_EAD"]);
                        eadSummaryByCounterPartyType.BATCH1_RWA = Convert.ToString(row["BATCH1_RWA"]);
                        eadSummaryByCounterPartyType.BATCH1_CVA_RWA = Convert.ToString(row["BATCH1_CVA_RWA"]);
                        eadSummaryByCounterPartyType.BATCH1_TotalRWA = Convert.ToString(row["BATCH1_TotalRWA"]);

                        eadSummaryByCounterPartyType.BATCH2_MTM = Convert.ToString(row["BATCH2_MTM"]);
                        eadSummaryByCounterPartyType.BATCH2_C = Convert.ToString(row["BATCH2_C"]);
                        eadSummaryByCounterPartyType.BATCH2_RC = Convert.ToString(row["BATCH2_RC"]);
                        eadSummaryByCounterPartyType.BATCH2_PFE_IR = Convert.ToString(row["BATCH2_PFE_IR"]);
                        eadSummaryByCounterPartyType.BATCH2_PFE_FX = Convert.ToString(row["BATCH2_PFE_FX"]);
                        eadSummaryByCounterPartyType.BATCH2_PFE_COMM = Convert.ToString(row["BATCH2_PFE_COMM"]);
                        eadSummaryByCounterPartyType.BATCH2_PFE_CR = Convert.ToString(row["BATCH2_PFE_CR"]);
                        eadSummaryByCounterPartyType.BATCH2_PFE_Eq = Convert.ToString(row["BATCH2_PFE_Eq"]);
                        eadSummaryByCounterPartyType.BATCH2_PFE = Convert.ToString(row["BATCH2_PFE"]);
                        eadSummaryByCounterPartyType.BATCH2_EAD = Convert.ToString(row["BATCH2_EAD"]);
                        eadSummaryByCounterPartyType.BATCH2_RWA = Convert.ToString(row["BATCH2_RWA"]);
                        eadSummaryByCounterPartyType.BATCH2_CVA_RWA = Convert.ToString(row["BATCH2_CVA_RWA"]);
                        eadSummaryByCounterPartyType.BATCH2_TotalRWA = Convert.ToString(row["BATCH2_TotalRWA"]);

                        res.Add(eadSummaryByCounterPartyType);
                    }
                }
                return res;
            }
        }

        #endregion
    }
}
