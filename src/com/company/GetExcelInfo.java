package com.company;
//package jxl.zhanhj;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.Number;
import jxl.write.DateTime;
import java.util.*;
public class GetExcelInfo {

    public static Set<String> ignoreSuppliersSet = new HashSet();
    public static Set<String> ignoreWords = new HashSet();
    public static String[] ignoreSuppliers = {"上海科东自控工程有限公司","宁波永祥铸造有限公司","上海建蓓铸造有限公司","上海力泽机械制造有限公司",
            "上海明燕机械制造有限公司","上海铨莳机械有限公司","上海松江振中五金厂","泰兴市凯越机械配件厂","上海晚初自动化科技有限公司",
            "上海泽怡电气设备有限公司","上海长潼机电科技有限公司","昆山灿坤机械制造有限公司","上海观邑机电有限公司","上海宏赛自动化电气有限公司",
            "上海弘秀机械有限公司","保全(上海)建材有限公司","东莞市翔河机械设备有限公司","东莞市星锐机械设备有限公司","东莞市长安毅创五金模具加工店",
            "广州市保全普美建筑材料有限公司","恺逊自动化科技（上海）有限公司","昆山市研通模具有限公司","萨普精密金属（太仓）有限公司",
            "上海倍固实业有限公司","上海诚枫金属制品有限公司","上海初锐机械有限公司","上海沸莱德表面处理有限公司","上海焕泰实业有限公司",
            "上海珈宇模塑科技有限公司","上海科晶精密制造有限公司","上海铃锐模具有限公司","上海铭嵌模塑科技有限公司","上海腾跞机械厂",
            "上海颖则机械有限公司","上海章彪金属制品有限公司","太仓夏鑫电镀有限公司","一胜百模具技术(上海)有限公司","永康市庄塔工贸有限公司",
            "上海衡灵自动化设备有限公司"};

    public static Queue<Cell[]> usefulData = new LinkedList();
    public static void main(String[] args) {
	// write your code here
        GetExcelInfo obj = new GetExcelInfo();
        File file = new File("../PO Receiving 20170606-20180606_Processing.xls"); // Read files.
        obj.readExcel(file);
        obj.write2Excel();
    }

    public static void initialization(){
        for(int i = 0; i < ignoreSuppliers.length; i++){
            ignoreSuppliersSet.add(ignoreSuppliers[i]);
        }
        ignoreWords.add("材料");
        ignoreWords.add("物流");
    }

    public void readExcel(File file) {
        try {
            // 创建输入流，读取Excel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            // jxl提供的Workbook类
            Workbook wb = Workbook.getWorkbook(is);
            // Excel的页签数量
            int sheet_size = wb.getNumberOfSheets();
            Sheet dataSheet = null;
            for (int index = 0; index < sheet_size; index++) {
                // 每个页签创建一个Sheet对象
                if(wb.getSheet(index).getName().equals("Data")){
                    dataSheet = wb.getSheet(index);
                }
                /*
                // sheet.getRows()返回该页的总行数
                for (int i = 0; i < sheet.getRows(); i++) {
                    // sheet.getColumns()返回该页的总列数
                    for (int j = 0; j < sheet.getColumns(); j++) {
                        String cellinfo = sheet.getCell(j, i).getContents();
                        System.out.println(cellinfo);
                    }
                }*/
            }
            for(int i = 0; i < dataSheet.getRows(); i++){
                String suppliers = dataSheet.getCell(6, i).getContents();
                for(int j = 0; j < ignoreSuppliers.length; j++){
                    if(ignoreSuppliers[j].equals(suppliers)){
                        usefulData.add(dataSheet.getRow(i));
                    }
                }
                //System.out.println(suppliers);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void write2Excel(){
        try {
            WritableWorkbook myFile = Workbook.createWorkbook(new File("../usefulData.xls"));
            WritableSheet sheet = myFile.createSheet("Sheet1", 0);
            int currRowNum = 0;
            while(!usefulData.isEmpty()){
                Cell[] currRow = usefulData.poll();
                int currCol = 0;
                for(int i = 0; i < currRow.length; i++){
                    if(i == 0){
                        DateTime currCell = new DateTime(currCol, currRowNum, new Date(currRow[i].getContents()));
                        sheet.addCell(currCell);
                        currCol++;
                    }
                    else if(i == 2){
                        DateTime currCell = new DateTime(currCol, currRowNum, new Date(currRow[i].getContents()));
                        sheet.addCell(currCell);
                        currCol++;
                    }
                    else if(i == 6){
                        Label currCell = new Label(2, currRowNum, currRow[i].getContents());
                        sheet.addCell(currCell);
                        currCol++;
                    }
                }
                currRowNum++;
            }
            myFile.write();
            myFile.close();
        } catch (Exception e){
            System.out.println(e);
        }
    }
}
