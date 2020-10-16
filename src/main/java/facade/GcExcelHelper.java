/*
* To change this license header, choose License Headers in Project Properties.
* To change this template file, choose Tools | Templates
* and open the template in the editor.
*/
package facade;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.ILegend;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.sun.javafx.charts.Legend;
import java.io.File;
import res.Strings;

/**
 *
 * @author diego
 */
public class GcExcelHelper {
    
    private File graphs;
    
    public GcExcelHelper(File graphs){
        this.graphs=graphs;
    }
    
    public void genGraphs(){
        
        
        Workbook workbook = new Workbook();
        workbook.open(graphs.getAbsolutePath());
        
        //Graph1
        IWorksheet worksheet0 = workbook.getWorksheets().get(0);
        IShape shape0 = worksheet0.getShapes().addChart(ChartType.BarClustered, 250, 20, 360, 230);
        shape0.getChart().getSeriesCollection().add(worksheet0.getRange("A1:C4"), RowCol.Columns);
        shape0.getChart().getChartTitle().setText(Strings.GRAPH_1_TITLE);
        shape0.getChart().setHasLegend(true);
        
        //Graph2
        int nextYear=DocHelper.getYear()+1;
        IWorksheet worksheet1 = workbook.getWorksheets().get(1);
        IShape shape1 = worksheet1.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        shape1.getChart().getSeriesCollection().add(worksheet1.getRange("A1:G3"), RowCol.Rows);
        shape1.getChart().getChartTitle().setText(Strings.GRAPH_2_TITLE+DocHelper.getYear()+"/"+nextYear);
        shape1.getChart().setHasLegend(true);
        
        //Graph3
        IWorksheet worksheet2 = workbook.getWorksheets().get(2);
        IShape shape2 = worksheet2.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        shape2.getChart().getSeriesCollection().add(worksheet2.getRange("A1:D3"), RowCol.Rows);
        shape2.getChart().getChartTitle().setText(Strings.GRAPH_3_TITLE+DocHelper.getYear()+"/"+nextYear);
        shape2.getChart().setHasLegend(true);
        
        //Graph4
        IWorksheet worksheet3 = workbook.getWorksheets().get(3);
        IShape shape3 = worksheet3.getShapes().addChart(ChartType.Pie, 250, 20, 360, 230);
        shape3.getChart().getSeriesCollection().add(worksheet3.getRange("A1:B5"), RowCol.Columns, true, true);
        shape3.getChart().getChartTitle().getTextFrame().getTextRange().getParagraphs().add(Strings.GRAPH_4_TITLE+DocHelper.getYear()+"/"+nextYear);
        
        //Graph5
        IWorksheet worksheet4 = workbook.getWorksheets().get(4);
        IShape shape4 = worksheet4.getShapes().addChart(ChartType.Pie, 250, 20, 360, 230);
        shape4.getChart().getSeriesCollection().add(worksheet4.getRange("A1:B5"), RowCol.Columns, true, true);
        shape4.getChart().getChartTitle().getTextFrame().getTextRange().getParagraphs().add(Strings.GRAPH_5_TITLE+DocHelper.getYear()+"/"+nextYear);
        shape4.getChart().setHasLegend(true);
        
        ILegend legend = shape4.getChart().getLegend();
        legend.setIncludeInLayout(true);
        
        //Graph6
        IWorksheet worksheet5 = workbook.getWorksheets().get(5);
        IShape shape5 = worksheet5.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        shape5.getChart().getSeriesCollection().add(worksheet5.getRange("A1:F3"), RowCol.Rows);
        shape5.getChart().getChartTitle().setText(Strings.GRAPH_6_TITLE);
        shape5.getChart().setHasLegend(true);
        
        //Graph 8
        IWorksheet worksheet7 = workbook.getWorksheets().get(7);
        IShape shape6 = worksheet7.getShapes().addChart(ChartType.Pie, 250, 20, 360, 230);
        shape6.getChart().getSeriesCollection().add(worksheet7.getRange("A1:B8"), RowCol.Columns, true, true);
        shape6.getChart().getChartTitle().getTextFrame().getTextRange().getParagraphs().add(Strings.GRAPH_8_TITLE);
        shape6.getChart().setHasLegend(true);
        
        //Graph 9
        IWorksheet worksheet8 = workbook.getWorksheets().get(8);
        IShape shape7 = worksheet8.getShapes().addChart(ChartType.LineMarkers, 250, 20, 360, 230);
        shape7.getChart().getSeriesCollection().add(worksheet8.getRange("A2:B13"), RowCol.Columns, true, true);
        shape7.getChart().getChartTitle().getTextFrame().getTextRange().getParagraphs().add(Strings.GRAPH_9_TITLE);
        
        
        //Graph 10
        IWorksheet worksheet9 = workbook.getWorksheets().get(9);
        IShape shape8 = worksheet9.getShapes().addChart(ChartType.LineMarkers, 250, 20, 360, 230);
        shape8.getChart().getSeriesCollection().add(worksheet9.getRange("A2:B13"), RowCol.Columns, true, true);
        shape8.getChart().getChartTitle().getTextFrame().getTextRange().getParagraphs().add(Strings.GRAPH_9_TITLE);
        
        //Graph 11
        IWorksheet worksheet10 = workbook.getWorksheets().get(9);
        IShape shape9 = worksheet10.getShapes().addChart(ChartType.LineMarkers, 250, 20, 360, 230);
        shape9.getChart().getSeriesCollection().add(worksheet9.getRange("A1:C13"), RowCol.Columns, true, true);
        shape9.getChart().getChartTitle().getTextFrame().getTextRange().getParagraphs().add(Strings.GRAPH_9_TITLE);
        shape9.getChart().setHasLegend(true);
        
        
        workbook.save(graphs.getPath());
    }
    
    
    
    
}
