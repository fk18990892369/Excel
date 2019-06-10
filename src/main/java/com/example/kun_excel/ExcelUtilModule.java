package com.example.kun_excel;

import android.app.Activity;
import android.os.Environment;

import android.util.Log;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.taobao.weex.WXSDKEngine;
import com.taobao.weex.annotation.JSMethod;
import com.taobao.weex.bridge.JSCallback;
import com.taobao.weex.utils.WXLogUtils;
import com.taobao.weex.utils.WXResourceUtils;

public class ExcelUtilModule extends WXSDKEngine.DestroyableModule {

    private ArrayList<ArrayList<String>> recordList;
    private List<Student> students;
    private static String[] title = { "编号","单位名称","所属集团","是否合格","类型","性质","问题描述","安全防护要求","治理措施","问题详情","人员/联系方式" };
    private File file;
    private String fileName;

    public static WritableFont arial14font = null;

    public static WritableCellFormat arial14format = null;
    public static WritableFont arial10font = null;
    public static WritableCellFormat arial10format = null;
    public static WritableFont arial12font = null;
    public static WritableCellFormat arial12format = null;

    public final static String UTF8_ENCODING = "UTF-8";
    public final static String GBK_ENCODING = "GBK";

    /**
     * 导出excel
     */
    @JSMethod(uiThread = true)
    public void exportExcel(JSONArray options, String TableName, JSCallback jsCallback) {

//        Log.i("debug", "导出excel");

        if (mWXSDKInstance.getContext() instanceof Activity) {
            //模拟数据集合
            students = new ArrayList<>();
            for (int i=0;i<options.size();i++){
                JSONObject jsonObject=(JSONObject)options.get(i);

                String orderNumber = jsonObject.getString("orderNumber"); //编号
                String unitName = jsonObject.getString("unitName"); //单位名称
                String affiliatedGroup = jsonObject.getString("affiliatedGroup"); //所属集团
                String isQualified = jsonObject.getString("isQualified"); //是否合格
                String type = jsonObject.getString("type"); //类型
                String rank = jsonObject.getString("rank"); //性质
                String checkop = jsonObject.getString("checkop"); //问题描述
                String checkcontent = jsonObject.getString("checkcontent"); //安全防护要求
                String biaozhun = jsonObject.getString("biaozhun"); //治理措施
                String problemDetails = jsonObject.getString("problemDetails"); //问题详情
                String contactInformation = jsonObject.getString("contactInformation"); //人员/联系方式
                students.add(new Student(orderNumber,unitName,affiliatedGroup,isQualified,type,rank,checkop,checkcontent,biaozhun,problemDetails, contactInformation));
            }

            file = new File(getSDPath() + "/Record");
            makeDir(file);

            if(TableName == null){
                TableName = "电厂统计表";
            }

            initExcel(file.toString() + "/"+ TableName +".xls", title);
            fileName = getSDPath() + "/Record/"+ TableName +".xls";
            writeObjListToExcel(getRecordData(), fileName, jsCallback);
        }

    }

    /**
     * 将数据集合 转化成ArrayList<ArrayList<String>>
     * @return
     */
    private  ArrayList<ArrayList<String>> getRecordData() {
        recordList = new ArrayList<>();
        for (int i = 0; i <students.size(); i++) {
            Student student = students.get(i);
            ArrayList<String> beanList = new ArrayList<String>();
            beanList.add(student.orderNumber);
            beanList.add(student.unitName);
            beanList.add(student.affiliatedGroup);
            beanList.add(student.isQualified);
            beanList.add(student.type);
            beanList.add(student.rank);
            beanList.add(student.checkop);
            beanList.add(student.checkcontent);
            beanList.add(student.biaozhun);
            beanList.add(student.problemDetails);
            beanList.add(student.contactInformation);

            recordList.add(beanList);
        }
        return recordList;
    }

    private  String getSDPath() {
        File sdDir = null;
        boolean sdCardExist = Environment.getExternalStorageState().equals(
                android.os.Environment.MEDIA_MOUNTED);
        if (sdCardExist) {
            sdDir = Environment.getExternalStorageDirectory();
        }
        String dir = sdDir.toString();
        return dir;
    }

    public  void makeDir(File dir) {
        if (!dir.getParentFile().exists()) {
            makeDir(dir.getParentFile());
        }
        dir.mkdir();
    }


    /**
     * 单元格的格式设置 字体大小 颜色 对齐方式、背景颜色等...
     */
    public static void format() {
        try {
            arial14font = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
            arial14font.setColour(jxl.format.Colour.LIGHT_BLUE);
            arial14format = new WritableCellFormat(arial14font);
            arial14format.setAlignment(jxl.format.Alignment.CENTRE);
            arial14format.setBorder(jxl.format.Border.ALL,jxl.format.BorderLineStyle.THIN);
            arial14format.setBackground(jxl.format.Colour.VERY_LIGHT_YELLOW);

            arial10font = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
            arial10format = new WritableCellFormat(arial10font);
            arial10format.setAlignment(jxl.format.Alignment.CENTRE);
            arial10format.setBorder(jxl.format.Border.ALL,jxl.format.BorderLineStyle.THIN);
            arial10format.setBackground(Colour.GRAY_25);

            arial12font = new WritableFont(WritableFont.ARIAL, 10);
            arial12format = new WritableCellFormat(arial12font);
            arial10format.setAlignment(jxl.format.Alignment.CENTRE);//对齐格式
            arial12format.setBorder(jxl.format.Border.ALL,jxl.format.BorderLineStyle.THIN); //设置边框

        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    /**
     * 初始化Excel
     * @param fileName
     * @param colName
     */
    @JSMethod(uiThread = true)
    public static void initExcel(String fileName, String[] colName) {
        format();
        WritableWorkbook workbook = null;
        try {
            File file = new File(fileName);
            if (!file.exists()) {
                file.createNewFile();
            }
            workbook = Workbook.createWorkbook(file);
            WritableSheet sheet = workbook.createSheet("任务统计", 0);
            //创建标题栏
            sheet.addCell((WritableCell) new Label(0, 0, fileName,arial14format));
            for (int col = 0; col < colName.length; col++) {
                sheet.addCell(new Label(col, 0, colName[col], arial10format));
            }
            sheet.setRowView(0,340); //设置行高

            workbook.write();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    @JSMethod(uiThread = true)
    @SuppressWarnings("unchecked")
    public static <T> void writeObjListToExcel(List<T> objList,String fileName, final JSCallback jsCallback) {
        if (objList != null && objList.size() > 0) {
            WritableWorkbook writebook = null;
            InputStream in = null;
            try {
                WorkbookSettings setEncode = new WorkbookSettings();
                setEncode.setEncoding(UTF8_ENCODING);
                in = new FileInputStream(new File(fileName));
                Workbook workbook = Workbook.getWorkbook(in);
                writebook = Workbook.createWorkbook(new File(fileName),workbook);
                WritableSheet sheet = writebook.getSheet(0);

//				sheet.mergeCells(0,1,0,objList.size()); //合并单元格
//				sheet.mergeCells()

                for (int j = 0; j < objList.size(); j++) {
                    ArrayList<String> list = (ArrayList<String>) objList.get(j);
                    for (int i = 0; i < list.size(); i++) {
                        sheet.addCell(new Label(i, j + 1, list.get(i),arial12format));
                        if (list.get(i).length() <= 5){
                            sheet.setColumnView(i,list.get(i).length()+8); //设置列宽
                        }else {
                            sheet.setColumnView(i,list.get(i).length()+5); //设置列宽
                        }
                    }
                    sheet.setRowView(j+1,350); //设置行高
                }

                writebook.write();
                Log.i("debug", "导出到手机存储中文件夹Record成功:" + fileName);
                jsCallback.invoke("导出成功");
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (writebook != null) {
                    try {
                        writebook.close();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }

                }
                if (in != null) {
                    try {
                        in.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }

        }
    }

    @Override
    public void destroy() {

    }
}
