package com.youyan.adapterpoi;


import com.youyan.adapterpoi.word.ExportWord;
import com.youyan.adapterpoi.word.ExportWordMgsyl;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RunWith(SpringRunner.class)
@SpringBootTest
public class AdapterPoiApplicationTests {

    @Test
    public void contextLoads() {
        ExportWord ew = new ExportWord();
        XWPFDocument document = ew.createXWPFDocument();
        List<List<Object>> list = new ArrayList<List<Object>>();

        List<Object> tempList = new ArrayList<Object>();
        tempList.add("姓名");
        tempList.add("黄xx");
        tempList.add("性别");
        tempList.add("男");
        tempList.add("出生日期");
        tempList.add("2018-10-10");
        list.add(tempList);

        tempList = new ArrayList<Object>();
        tempList.add("身份证号");
        tempList.add("36073xxxxxxxxxxx");
        list.add(tempList);

        tempList = new ArrayList<Object>();
        tempList.add("出生地");
        tempList.add("江西");
        tempList.add("名族");
        tempList.add("汉");
        tempList.add("婚否");
        tempList.add("否");
        list.add(tempList);

        tempList = new ArrayList<Object>();
        tempList.add("既往病史");
        tempList.add("无");
        list.add(tempList);

        Map<String, Object> dataList = new HashMap<String, Object>();
        dataList.put("TITLE", "个人体检表");
        dataList.put("TABLEDATA", list);
        try {
            ew.exportCheckWord(dataList, document, "E:/expWordTest.docx");
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("文档生成成功");

    }

    @Test
    public void createMgsylExport() {
        Map<String,Object> argsMap = new HashMap<>();
        argsMap.put("wfName","3301工作面");
        argsMap.put("savePath","E:/expWordMgsyl.docx");
        argsMap.put("startDate","2020-05-20 10:20:05");
        argsMap.put("endDate","2020-05-20 10:20:05");
        argsMap.put("currentDate","2020-05-20 10:20:05");
        argsMap.put("sensorNum",4);
        argsMap.put("alarmNum",4);
        argsMap.put("signNum",4);




        String wfName = (String) argsMap.get("wfName");
        String savePath = (String) argsMap.get("savePath");
        String startDate = (String) argsMap.get("startDate");
        String endDate = (String) argsMap.get("endDate");
        String currentDate = (String) argsMap.get("currentDate");

        int sensorNum = (int) argsMap.get("sensorNum");
        int alarmNum = (int) argsMap.get("alarmNum");
        int signNum = (int) argsMap.get("signNum");


        try {
            ExportWordMgsyl ew = new ExportWordMgsyl();
            //监测数据统计表的行数sensorNum
            //报警信息统计的行数alarmNum
            //签名栏的个数signNum
            XWPFDocument document = ew.createXWPFDocument(sensorNum,alarmNum,signNum);
            List<List<Object>> list = new ArrayList<>();

            //region第1行
            List<Object> temList = new ArrayList<>();
            temList.add("工作面/巷道");
            temList.add("3500工作面");
            temList.add("");
            temList.add("面长");
            temList.add("1000m");
            temList.add("设计推采长度");
            temList.add("");
            temList.add("800m");
            temList.add("采煤工艺");
            temList.add("露天挖掘");
            list.add(temList);
            //endregion

            //region第2行
            temList = new ArrayList<>();
            temList.add("推采进度");
            temList.add("本日进尺");
            temList.add("");
            temList.add("上巷");
            temList.add("10");
            temList.add("总进尺");
            temList.add("上巷");
            temList.add("20");
            temList.add("剩余进尺");
            temList.add("上巷");
            temList.add("30");
            list.add(temList);
            //endregion

            //region第3行
            temList = new ArrayList<>();
            temList.add("");
            temList.add("");
            temList.add("");
            temList.add("下巷");
            temList.add("10");
            temList.add("");
            temList.add("下巷");
            temList.add("20");
            temList.add("");
            temList.add("下巷");
            temList.add("30");
            list.add(temList);
            //endregion

            //region第4行
            temList = new ArrayList<>();
            temList.add("");
            temList.add("");
            temList.add("");
            temList.add("平均");
            temList.add("10");
            temList.add("");
            temList.add("平均");
            temList.add("20");
            temList.add("");
            temList.add("平均");
            temList.add("30");
            list.add(temList);
            //endregion

            //region第5行
            temList = new ArrayList<>();
            temList.add("监测数据统计表");
            list.add(temList);
            //endregion

            //region第6行
            temList = new ArrayList<>();
            temList.add("测点名");
            temList.add("编号");
            temList.add("类型");
            temList.add("巷道");
            temList.add("安装位置");
            temList.add("距工作面距离");
            temList.add("监测对象");
            temList.add("安装方式");
            temList.add("当前值");
            temList.add("当日增量");
            temList.add("备注");
            list.add(temList);
            //endregion
            for (int i=1;i<sensorNum;i++){
                //region第7行
                temList = new ArrayList<>();
                temList.add("测点1");
                temList.add("1");
                temList.add("锚杆");
                temList.add("3500轨道顺槽");
                temList.add("200m");
                temList.add("50m");
                temList.add("锚杆");
                temList.add("帮部");
                temList.add("58");
                temList.add("2");
                temList.add("");
                list.add(temList);
                //endregion
            }

            //region第8行
            temList = new ArrayList<>();
            temList.add("监测数据柱状图");
            list.add(temList);
            //endregion

            //region第9行
            temList = new ArrayList<>();
            temList.add("");
            list.add(temList);
            //endregion

            //region第10行
            temList = new ArrayList<>();
            temList.add("报警信息统计");
            list.add(temList);
            //endregion

            //region第11行
            temList = new ArrayList<>();
            temList.add("时间");
            temList.add("测点名称");
            temList.add("传感器编号");
            temList.add("所属巷道");
            temList.add("安装位置");
            temList.add("距工作面距离");
            temList.add("监测对象");
            temList.add("安装方位");
            temList.add("离层量");
            temList.add("预警预案");
            temList.add("预警等级");
            list.add(temList);
            //endregion
            for (int i=1;i<alarmNum;i++){
                //region第12行
                temList = new ArrayList<>();
                temList.add("");
                list.add(temList);
                //endregion
            }

            //region第13行
            temList = new ArrayList<>();
            temList.add("监测分析结论");
            temList.add("当日最大值为XX测点XXMPa/kN：当日最大增量值（正）为XX测点XXMPa/kN：\n" +
                    "该时间段内出现0次预警。\n");
            list.add(temList);
            //endregion

            //region第14行
            temList = new ArrayList<>();
            temList.add("处理措施");
            temList.add("如有报警信息，可导入消警备注中处理措施。若无报警信息，此处空,可人工输入。");
            list.add(temList);
            //endregion

            //region第15行
            temList = new ArrayList<>();
            temList.add("签字栏");
            temList.add("正处");
            temList.add("王大小");
            temList.add("副处");
            temList.add("王二小");
            temList.add("正科");
            temList.add("王三小");
            temList.add("副科");
            temList.add("王四小");
            temList.add("安全员");
            temList.add("王五小");
            list.add(temList);
            //endregion

            //region
            Map<String, Object> dataList = new HashMap<>();
            dataList.put("TITLE", wfName+"锚杆(索)应力监测报表");
            dataList.put("DATE", "监测日期:"+startDate+" -- "+endDate+"             上报日期: "+currentDate);
            dataList.put("TABLEDATA", list);
            //endregion

            ew.exportCheckWord(dataList, document, savePath,sensorNum,alarmNum,signNum);
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    @Test
    public void createZzt() {
        try {
            XWPFDocument document = new XWPFDocument();

            // create the data
            String[] categories = new String[]{"类别1", "类别2", "类别3"};
            Double[] valuesA = new Double[]{10d, 20d, 30d};//y轴
            Double[] valuesB = new Double[]{15d, 25d, 155d};

            // create the chart
            XWPFChart chart = document.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

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
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
//            XDDFBar3DChartData data = (XDDFBar3DChartData) chart.createData(ChartTypes.BAR3D, bottomAxis, leftAxis);
            ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);

            // create series
            // if only one series do not vary colors for each bar
            ((XDDFBarChartData) data).setVaryColors(false);
            XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
            // XDDFChart.setSheetTitle is buggy. It creates a Table but only half way and incomplete.
            // Excel cannot opening the workbook after creatingg that incomplete Table.
            // So updating the chart data in Word is not possible.
            //series.setTitle("a", chart.setSheetTitle("a", 1));
            series.setTitle("", setTitleInDataSheet(chart, "", 0));

			/*
			   // if more than one series do vary colors of the series
			   ((XDDFBarChartData)data).setVaryColors(true);
			   series = data.addSeries(categoriesData, valuesDataB);
			   //series.setTitle("b", chart.setSheetTitle("b", 2));
			   series.setTitle("b", setTitleInDataSheet(chart, "b", 2));
			*/

            // plot chart data
            chart.plot(data);

            // create legend
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);
            legend.setOverlay(false);

            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream("E:/CreateWordXDDFChart.docx");
            document.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
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

    @Test
    public void createZxt() {
        try {
            XWPFDocument document = new XWPFDocument();
            // create the chart
            XWPFChart chart = document.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

            //标题
//            chart.setTitleText("地区排名前七的国家");
            //标题覆盖
//            chart.setTitleOverlay(false);

            //图例位置
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP);

            //分类轴标(X轴),标题位置
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("传感器");
            //值(Y轴)轴,标题位置
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("钻孔应力");

            //CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）
            //分类轴标(X轴)数据，单元格范围位置[0, 0]到[0, 6]
//			XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, 6));
            XDDFCategoryDataSource countries = XDDFDataSourcesFactory.fromArray(new String[]{"测点1", "测点2", "测点3", "测点4", "测点5", "测点6", "测点7"}, "jiang");
            //数据1，单元格范围位置[1, 0]到[1, 6]
//			XDDFNumericalDataSource<Double> area = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 6));
            XDDFNumericalDataSource<Integer> area = XDDFDataSourcesFactory.fromArray(new Integer[]{9, 7, 1, 3, 5, 8, 4}, "jiangye");

            //数据1，单元格范围位置[2, 0]到[2, 6]
//			XDDFNumericalDataSource<Double> population = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, 6));

            //LINE：折线图，
            XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
//            XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.PIE, bottomAxis, leftAxis);

            //图表加载数据，折线1
            XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(countries, area);
            //折线图例标题
            series1.setTitle("应力值", null);
            //直线
            series1.setSmooth(false);
            //设置标记大小
            series1.setMarkerSize((short) 6);
            //设置标记样式，星星
            series1.setMarkerStyle(MarkerStyle.STAR);


            //绘制
            chart.plot(data);

            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream("E:/CreateWordXDDFChart.docx");
            document.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    @Test
    public void createBzt() {
        try (XWPFDocument document = new XWPFDocument()) {

            // create the chart
            XWPFChart chart = document.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

            //标题
            chart.setTitleText("地区排名前七的国家");
            //标题是否覆盖图表
            chart.setTitleOverlay(false);

            //图例位置
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);

            //CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）
            //分类轴标数据，
//			XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, 6));
            XDDFCategoryDataSource countries = XDDFDataSourcesFactory.fromArray(new String[]{"俄罗斯", "加拿大", "美国", "中国", "巴西", "澳大利亚", "印度"}, "jiang");
            //数据1，
//			XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 6));
            XDDFNumericalDataSource<Integer> values = XDDFDataSourcesFactory.fromArray(new Integer[]{17098242, 9984670, 9826675, 9596961, 8514877, 7741220, 3287263}, "jiangye");
            //XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);

            //分类轴标(X轴),标题位置

//            XDDFChartData createData(ChartTypes type, XDDFChartAxis category, XDDFValueAxis values)
            XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
            //设置为可变颜色
            data.setVaryColors(true);
            //图表加载数据
            data.addSeries(countries, values);


            //绘制
            chart.plot(data);

            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream("E:/CreateWordXDDFChart.docx");
            document.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }



}
