package com.youyan.adapterpoi.word;

import com.youyan.adapterpoi.util.XWPFHelper;
import com.youyan.adapterpoi.util.XWPFHelperTable;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;


/**
 * @Description 导出word文档
 * @Author  Huangxiaocong
 * 2018年12月1日  下午12:12:15
 */
public class ExportWordMgsyl {
    private XWPFHelperTable xwpfHelperTable = null;
    private XWPFHelper xwpfHelper = null;
    public ExportWordMgsyl() {
        xwpfHelperTable = new XWPFHelperTable();
        xwpfHelper = new XWPFHelper();
    }
    /**
     * 创建好文档的基本 标题，表格  段落等部分
     * @return
     * @Author Huangxiaocong 2018年12月16日
     */
    public XWPFDocument createXWPFDocument(int sensorNum,int alarmNum,int signNum) {
        XWPFDocument doc = new XWPFDocument();

        createTitleParagraph(doc);
        createDateParagraph(doc);

        createTableParagraph(doc, 11+sensorNum+alarmNum, 11,sensorNum,alarmNum,signNum);

        return doc;
    }
    public void createBzt(XWPFDocument doc,XWPFRun run){

        //region饼状图
        try {
            XWPFChart chart = doc.createChart(run,19 * Units.EMU_PER_CENTIMETER, 3 * Units.EMU_PER_CENTIMETER);
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);
            XDDFCategoryDataSource countries = XDDFDataSourcesFactory.fromArray(new String[]{"类型1", "类型2", "类型3", "类型4", "类型5", "类型6", "类型7"}, "jiang");
            XDDFNumericalDataSource<Integer> values = XDDFDataSourcesFactory.fromArray(new Integer[]{1, 2, 3, 4, 5, 6, 7}, "jiangye");
            XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);
            data.addSeries(countries, values);
            chart.plot(data);
        }catch (Exception e){

        }
        //endregion



        //region柱状图
        /*
        try {
            XWPFDocument document = new XWPFDocument();

            // create the data
            String[] categories = new String[]{"类别1", "类别2", "类别3"};
            Double[] valuesA = new Double[]{10d, 20d, 30d};//y轴
            Double[] valuesB = new Double[]{15d, 25d, 155d};

            // create the chart
            XWPFChart chart = doc.createChart(run,19 * Units.EMU_PER_CENTIMETER, 6 * Units.EMU_PER_CENTIMETER);

            // create data sources
            int numOfPoints = categories.length;
            String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
            String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
            String valuesDataRangeB = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
            XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
            XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesA, valuesDataRangeA, 1);
            XDDFNumericalDataSource<Double> valuesDataB = XDDFDataSourcesFactory.fromArray(valuesB, valuesDataRangeB, 2);

            // create axis
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
            // Else first and last category is exactly on cross points and the bars are only half visible.
            leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

            // create chart data
//            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBar3DChartData data = (XDDFBar3DChartData) chart.createData(ChartTypes.BAR3D, bottomAxis, leftAxis);
            ((XDDFBar3DChartData) data).setBarDirection(BarDirection.COL);

            // create series
            // if only one series do not vary colors for each bar
            ((XDDFBar3DChartData) data).setVaryColors(false);
            XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
            // XDDFChart.setSheetTitle is buggy. It creates a Table but only half way and incomplete.
            // Excel cannot opening the workbook after creatingg that incomplete Table.
            // So updating the chart data in Word is not possible.
            //series.setTitle("a", chart.setSheetTitle("a", 1));
            series.setTitle("", setTitleInDataSheet(chart, "", 0));

			   // if more than one series do vary colors of the series
			   //((XDDFBarChartData)data).setVaryColors(true);
			   //series = data.addSeries(categoriesData, valuesDataB);
			   //series.setTitle("b", chart.setSheetTitle("b", 2));
			   //series.setTitle("b", setTitleInDataSheet(chart, "b", 2));

            // plot chart data
            chart.plot(data);

            // create legend
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);
            legend.setOverlay(false);

            // Write the output to a file
//            FileOutputStream fileOut = new FileOutputStream("E:/CreateWordXDDFChart.docx");
//            document.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
        */
        //endregion

    }
    static CellReference setTitleInDataSheet(XWPFChart chart, String title, int column) throws Exception {
        XSSFWorkbook workbook = chart.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        if (row == null)
            row = sheet.createRow(0);
        XSSFCell cell = row.getCell(column);
        if (cell == null)
            cell = row.createCell(column);
        cell.setCellValue(title);
        return new CellReference(sheet.getSheetName(), 0, column, true, true);
    }

    /**
     * 创建表格的标题样式
     * @param document
     * @Author Huangxiaocong 2018年12月16日 下午5:28:38
     */
    public void createDateParagraph(XWPFDocument document) {
        XWPFParagraph dateParagraph = document.createParagraph();    //新建一个标题段落对象（就是一段文字）
        dateParagraph.setAlignment(ParagraphAlignment.LEFT);//样式左
        XWPFRun dateFun = dateParagraph.createRun();    //创建文本对象
//        titleFun.setText(titleName); //设置标题的名字
//        dateFun.setBold(true); //加粗
        dateFun.setColor("000000");//设置颜色
        dateFun.setFontSize(10);    //字体大小
//        titleFun.setFontFamily("");//设置字体
        //...
//        titleFun.addBreak();    //换行
    }

    public void createTitleParagraph(XWPFDocument document) {
        XWPFParagraph titleParagraph = document.createParagraph();    //新建一个标题段落对象（就是一段文字）
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);//样式居中
        XWPFRun titleFun = titleParagraph.createRun();    //创建文本对象
//        titleFun.setText(titleName); //设置标题的名字
        titleFun.setBold(true); //加粗
        titleFun.setColor("000000");//设置颜色
        titleFun.setFontSize(14);    //字体大小
//        titleFun.setFontFamily("");//设置字体
        //...
//        titleFun.addBreak();    //换行
    }
    /**
     * 设置表格
     * @param document
     * @param rows
     * @param cols
     * @Author Huangxiaocong 2018年12月16日
     */
    public void createTableParagraph(XWPFDocument document, int rows, int cols,int sensorNum,int alarmNum,int signNum) {
        XWPFTable infoTable = document.createTable(rows, cols);
        xwpfHelperTable.setTableWidthAndHAlign(infoTable, "10072", STJc.CENTER);
        //合并表格
//        xwpfHelperTable.mergeCellsHorizontal(infoTable, 1, 1, 5);
//        xwpfHelperTable.mergeCellsVertically(infoTable, 0, 3, 6);
        //第1行
        xwpfHelperTable.mergeCellsHorizontal(infoTable, 0, 1, 2);
        xwpfHelperTable.mergeCellsHorizontal(infoTable, 0, 5, 6);
        xwpfHelperTable.mergeCellsHorizontal(infoTable, 0, 9, 10);
        //第2-4行
        for(int col = 1; col < 4; col++) {
            xwpfHelperTable.mergeCellsHorizontal(infoTable, col, 1, 2);
        }
        xwpfHelperTable.mergeCellsVertically(infoTable, 0, 1, 3);
        xwpfHelperTable.mergeCellsVertically(infoTable, 1, 1, 3);
        xwpfHelperTable.mergeCellsVertically(infoTable, 2, 1, 3);
        xwpfHelperTable.mergeCellsVertically(infoTable, 5, 1, 3);
        xwpfHelperTable.mergeCellsVertically(infoTable, 8, 1, 3);
        //第5行,监测数据统计报表表头
        xwpfHelperTable.mergeCellsHorizontal(infoTable, 4, 0, 10);

        //第8行,监测数据柱状图表头
        xwpfHelperTable.mergeCellsHorizontal(infoTable, 5+sensorNum, 0, 10);
//        xwpfHelperTable.mergeCellsHorizontal(infoTable, 7, 0, 10);

        //第9行,监测数据柱状图
        xwpfHelperTable.mergeCellsHorizontal(infoTable, 6+sensorNum, 0, 10);
//        xwpfHelperTable.mergeCellsHorizontal(infoTable, 8, 0, 10);
        //第10行，报警统计表头
        xwpfHelperTable.mergeCellsHorizontal(infoTable, 7+sensorNum, 0, 10);
//        xwpfHelperTable.mergeCellsHorizontal(infoTable, 9, 0, 10);
        //第11行
        //第12行
        //第13行,监测分析结论
        xwpfHelperTable.mergeCellsHorizontal(infoTable, 8+sensorNum+alarmNum, 1, 10);
//        xwpfHelperTable.mergeCellsHorizontal(infoTable, 12, 1, 10);
        //第14行,处理措施
        xwpfHelperTable.mergeCellsHorizontal(infoTable, 9+sensorNum+alarmNum, 1, 10);
//        xwpfHelperTable.mergeCellsHorizontal(infoTable, 13, 1, 10);

        //设置表格样式
        List<XWPFTableRow> rowList = infoTable.getRows();
        for(int i = 0; i < rowList.size(); i++) {
            XWPFTableRow infoTableRow = rowList.get(i);
            List<XWPFTableCell> cellList = infoTableRow.getTableCells();
            for(int j = 0; j < cellList.size(); j++) {
                XWPFParagraph cellParagraph = cellList.get(j).getParagraphArray(0);
                cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun cellParagraphRun = cellParagraph.createRun();
                cellParagraphRun.setFontSize(8);
                if(i % 2 != 0) {
//                    cellParagraphRun.setBold(true);
                }
            }
        }
        xwpfHelperTable.setTableHeight(infoTable, 560, STVerticalJc.CENTER);
    }

    /**
     * 往表格中填充数据
     * @param dataList
     * @param document
     * @throws IOException
     * @Author Huangxiaocong 2018年12月16日
     */
    @SuppressWarnings("unchecked")
    public void exportCheckWord(Map<String, Object> dataList, XWPFDocument document, String savePath,int sensorNum,int alarmNum,int signNum) throws Exception {
        //标题
        XWPFParagraph paragraph = document.getParagraphArray(0);
        XWPFRun titleFun = paragraph.getRuns().get(0);
        titleFun.setText(String.valueOf(dataList.get("TITLE")));
        //日期
        XWPFParagraph dateParagraph = document.getParagraphArray(1);
        XWPFRun dateFun = dateParagraph.getRuns().get(0);
        dateFun.setText(String.valueOf(dataList.get("DATE")));

        List<List<Object>> tableData = (List<List<Object>>) dataList.get("TABLEDATA");
        XWPFTable table = document.getTableArray(0);
        fillTableData(table, tableData, sensorNum, alarmNum,signNum);
        xwpfHelper.saveDocument(document, savePath);
    }
    /**
     * 往表格中填充数据
     * @param table
     * @param tableData
     * @Author Huangxiaocong 2018年12月16日
     */
    public void fillTableData(XWPFTable table, List<List<Object>> tableData,int sensorNum,int alarmNum,int signNum)throws Exception  {
        List<XWPFTableRow> rowList = table.getRows();
        for(int i = 0; i < tableData.size(); i++) {
            List<Object> list = tableData.get(i);
            List<XWPFTableCell> cellList = rowList.get(i).getTableCells();
            for(int j = 0; j < list.size(); j++) {
                XWPFTableCell xwpfTableCell = cellList.get(j);


                XWPFParagraph paragraphArray = xwpfTableCell.getParagraphArray(0);
                XWPFDocument document = paragraphArray.getDocument();


                XWPFParagraph cellParagraph = cellList.get(j).getParagraphArray(0);
                XWPFRun cellParagraphRun = cellParagraph.getRuns().get(0);

                if (i==(6+sensorNum)&&j==0){
                    createBzt(document,cellParagraphRun);
                }

                cellParagraphRun.setText(String.valueOf(list.get(j)));
            }
        }
    }
}