/*
 * Decompiled with CFR 0.152.
 */
package facade;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.ILegend;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.LegendPosition;
import com.grapecity.documents.excel.drawing.RowCol;
import facade.DocHelper;
import java.io.File;

public class GcExcelHelper {
    private File graphs;

    public GcExcelHelper(File graphs) {
        this.graphs = graphs;
    }

    public GcExcelHelper() {
    }

    public void genGraphs() {
        Workbook workbook = new Workbook();
        workbook.open(this.graphs.getAbsolutePath());
        IWorksheet worksheet0 = workbook.getWorksheets().get(0);
        IShape shape0 = worksheet0.getShapes().addChart(ChartType.BarClustered, 250.0, 20.0, 360.0, 230.0);
        shape0.getChart().getSeriesCollection().add(worksheet0.getRange("A1:C4"), RowCol.Columns);
        shape0.getChart().getChartTitle().setText("Gender disclosed for all Wessex trainees and PSW referrals");
        shape0.getChart().setHasLegend(true);
        int nextYear = DocHelper.getStartingYear() + 1;
        IWorksheet worksheet1 = workbook.getWorksheets().get(1);
        IShape shape1 = worksheet1.getShapes().addChart(ChartType.ColumnClustered, 250.0, 20.0, 360.0, 230.0);
        shape1.getChart().getSeriesCollection().add(worksheet1.getRange("A1:U3"), RowCol.Rows);
        shape1.getChart().getChartTitle().setText("Referral reason financial year ");
        shape1.getChart().setHasLegend(true);
        IWorksheet worksheet2 = workbook.getWorksheets().get(2);
        IShape shape2 = worksheet2.getShapes().addChart(ChartType.ColumnClustered, 250.0, 20.0, 360.0, 230.0);
        shape2.getChart().getSeriesCollection().add(worksheet2.getRange("A1:D3"), RowCol.Rows);
        shape2.getChart().getChartTitle().setText("Health Referrals (Breakdown) ");
        shape2.getChart().setHasLegend(true);
        IWorksheet worksheet3 = workbook.getWorksheets().get(3);
        IShape shape3 = worksheet3.getShapes().addChart(ChartType.Pie, 250.0, 20.0, 360.0, 230.0);
        shape3.getChart().getSeriesCollection().add(worksheet3.getRange("A1:B5"), RowCol.Columns, true, true);
        shape3.getChart().getChartTitle().setText("Referrals broken down by Stage of Training ");
        IWorksheet worksheet4 = workbook.getWorksheets().get(4);
        IShape shape4 = worksheet4.getShapes().addChart(ChartType.Pie, 250.0, 20.0, 360.0, 230.0);
        shape4.getChart().getSeriesCollection().add(worksheet4.getRange("A1:B5"), RowCol.Columns, true, true);
        shape4.getChart().getChartTitle().setText("Referrals broken down by programme grade for new referrals ");
        shape4.getChart().setHasLegend(true);
        IWorksheet worksheet5 = workbook.getWorksheets().get(5);
        IShape shape5 = worksheet5.getShapes().addChart(ChartType.ColumnClustered, 250.0, 20.0, 360.0, 230.0);
        shape5.getChart().getSeriesCollection().add(worksheet5.getRange("A1:F3"), RowCol.Rows);
        shape5.getChart().getChartTitle().setText("Duration of PSW input for current open cases");
        shape5.getChart().setHasLegend(true);
        IWorksheet worksheet6 = workbook.getWorksheets().get(6);
        IShape shape6 = worksheet6.getShapes().addChart(ChartType.Pie, 250.0, 20.0, 500.0, 450.0);
        shape6.getChart().getSeriesCollection().add(worksheet6.getRange("A1:B21"), RowCol.Columns, true, true);
        shape6.getChart().getChartTitle().setText("Reasons for referrals open for longer than 24 months");
        shape6.getChart().setHasLegend(true);
        ILegend leg = shape6.getChart().getLegend();
        leg.setIncludeInLayout(false);
        leg.setPosition(LegendPosition.Bottom);
        IWorksheet worksheet7 = workbook.getWorksheets().get(7);
        IShape shape7 = worksheet7.getShapes().addChart(ChartType.Pie, 250.0, 20.0, 360.0, 230.0);
        shape7.getChart().getSeriesCollection().add(worksheet7.getRange("A1:B8"), RowCol.Columns, true, true);
        shape7.getChart().getChartTitle().setText("Rolling Analysis of case closures");
        shape7.getChart().setHasLegend(true);
        IWorksheet worksheet8 = workbook.getWorksheets().get(8);
        IShape shape8 = worksheet8.getShapes().addChart(ChartType.ColumnStacked, 250.0, 20.0, 360.0, 230.0);
        shape8.getChart().getSeriesCollection().add(worksheet8.getRange("A1:C13"), RowCol.Columns, true, true);
        shape8.getChart().getChartTitle().setText("Open case load overall by month");
        shape8.getChart().setHasLegend(true);
        IWorksheet worksheet9 = workbook.getWorksheets().get(9);
        IShape shape9 = worksheet9.getShapes().addChart(ChartType.LineMarkers, 250.0, 20.0, 360.0, 230.0);
        shape9.getChart().getSeriesCollection().add(worksheet9.getRange("A1:B13"), RowCol.Columns, true, true);
        shape9.getChart().getChartTitle().setText("Total\u00a0Hours for PSW\u00a0CMs\u00a0and SSG Experts\u00a0by month");
        IWorksheet worksheet10 = workbook.getWorksheets().get(10);
        IShape shape10 = worksheet10.getShapes().addChart(ChartType.LineMarkers, 250.0, 20.0, 360.0, 230.0);
        shape10.getChart().getSeriesCollection().add(worksheet10.getRange("A1:B13"), RowCol.Columns, true, true);
        shape10.getChart().getChartTitle().setText("Monthly invoice trend for PSW and SSG experts");
        IWorksheet worksheet11 = workbook.getWorksheets().get(11);
        IShape shape11 = worksheet11.getShapes().addChart(ChartType.LineMarkers, 250.0, 20.0, 360.0, 230.0);
        shape11.getChart().getSeriesCollection().add(worksheet11.getRange("A1:C13"), RowCol.Columns, true, true);
        shape11.getChart().getChartTitle().setText("Monthly breakdown of CM and SSG invoices ");
        shape11.getChart().setHasLegend(true);
        workbook.save(this.graphs.getPath());
    }
}
