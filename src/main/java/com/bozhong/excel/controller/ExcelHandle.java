package com.bozhong.excel.controller;

import com.bozhong.excel.entity.FileInfoModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.http.fileupload.FileUtils;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@RestController
@RequestMapping(value = "/excel")
public class ExcelHandle {

    static Map<String,Object> fileMap = new HashMap<String,Object>();
    static List<FileInfoModel> fileList = new ArrayList<FileInfoModel>();
    private String uploadPath = "d:/shiroro/upload";
    private String downloadPath = "d:/shiroro/download";
    @ResponseBody
    @RequestMapping(value = "/receive")
    Map receive(HttpServletRequest request, HttpServletResponse response, MultipartFile file) throws IOException {
        String fileName = file.getOriginalFilename();
        String isTotal = request.getParameter("isTotal");
        String colName = request.getParameter("colName");
        //String idCol = request.getParameter("idCol");
        String valueCol = request.getParameter("valueCol");
        File destFile= new File(uploadPath);
        if(!destFile.exists()) {
            destFile.mkdirs();
        }
        File uploadFile = new File(destFile, fileName);
        file.transferTo(uploadFile);
        if("true".equals(isTotal)){
            fileMap.put("totalFileName",fileName);
        }else{
            FileInfoModel fileInfoModel = new FileInfoModel();
            fileInfoModel.setFileName(fileName);
            fileInfoModel.setColName(colName);
            //fileInfoModel.setIdCol(idCol);
            fileInfoModel.setValueCol(valueCol);
            fileList.add(fileInfoModel);
        }

        Map map = new HashMap();
        map.put("success","true");
        map.put("fileName",fileName);
        map.put("isTotal",isTotal);
        return map;
    }

    @ResponseBody
    @RequestMapping(value = "/clean")
    Map clean(HttpServletRequest request, HttpServletResponse response){
        delete(uploadPath);
        fileMap.clear();
        fileList.clear();
        Map map = new HashMap();
        map.put("success","true");
        return map;
    }

    public void delete(String path) {
        // 为传进来的路径参数创建一个文件对象
        File file = new File(path);
        // 如果目标路径是一个文件，那么直接调用delete方法删除即可
        // file.delete();
        // 如果是一个目录，那么必须把该目录下的所有文件和子目录全部删除，才能删除该目标目录，这里要用到递归函数
        // 创建一个files数组，用来存放目标目录下所有的文件和目录的file对象
        File[] files = new File[50];
        // 将目标目录下所有的file对象存入files数组中
        files = file.listFiles();
        // 循环遍历files数组
        for(File temp : files){
            // 判断该temp对象是否为文件对象
            if (temp.isFile()) {
                temp.delete();
            }
            // 判断该temp对象是否为目录对象
            if (temp.isDirectory()) {
                // 将该temp目录的路径给delete方法（自己），达到递归的目的
                delete(temp.getAbsolutePath());
                // 确保该temp目录下已被清空后，删除该temp目录
                temp.delete();
            }
        }
    }

    @ResponseBody
    @RequestMapping(value = "/create")
    ResponseEntity<byte[]> create(HttpServletRequest request, HttpServletResponse response){
        List<Map<String,String>> list = readTotal();
        readOtherFile(list);

        File file = createExcel(list);
        byte[] body = null;
        InputStream is = null;
        ResponseEntity<byte[]> entity = null;
        try {
            is = new FileInputStream(file);
            body = new byte[is.available()];
            is.read(body);
            HttpHeaders headers = new HttpHeaders();
            headers.add("Content-Disposition", "attchement;filename=" + file.getName());
            HttpStatus statusCode = HttpStatus.OK;
            entity = new ResponseEntity<>(body, headers, statusCode);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return entity;
    }

    private List<Map<String,String>> readTotal(){
        List<Map<String,String>> list = new ArrayList<Map<String,String>>();
        String totalFileName = (String)fileMap.get("totalFileName");
        try{
            InputStream ips = new FileInputStream(uploadPath+"/"+totalFileName);
            //创建Excel，读取文件内容
            XSSFWorkbook workbook = new XSSFWorkbook(ips);
            //获取第一个工作表workbook.getSheet("sheet0");(方式一)
            // HSSFSheet sheet = workbook.getSheet("Sheet0");
            //读取默认第一个工作表sheet（方式二）
            XSSFSheet sheet = workbook.getSheetAt(0);
            //读取工作表中的数据
            int firstRowNum = 3;
            //获取sheet中最后一行行号
            int lastRowNum = sheet.getLastRowNum();
            for (int i = firstRowNum; i <= lastRowNum ; i++) {
                boolean flag = true;
                Map<String,String> map = new LinkedHashMap<String,String>();
                XSSFRow row = sheet.getRow(i);
                //获取当前行最后单元格的列号
                int lastCellNum = row.getLastCellNum();
                for (int j = 0; j < lastCellNum; j++) {
                    //读取cell
                    XSSFCell cell = row.getCell(j);
                    //定义value，读取值
                    cell.setCellType(CellType.STRING);
                    String value= cell.getStringCellValue();
                    if(checkChinese(value)&&j==0){
                        flag = false;
                        break;
                    }

                    switch (j){
                        case 0 :
                            map.put("productId",value);
                            break;
                        case 1 :
                            map.put("productName",value);
                            break;
                        case 2 :
                            map.put("productStandard",value);
                            break;
                        case 3 :
                            map.put("productUnit",value);
                            break;
                        default :
                    }

                }
                if(flag==true){
                    list.add(map);
                }

            }

        }catch (IOException e){
            e.printStackTrace();
        }

        return list;

    }

    public boolean checkChinese(String countname)
    {
        Pattern p = Pattern.compile("[\u4e00-\u9fa5]");
        Matcher m = p.matcher(countname);
        if (m.find()) {
            return true;
        }
        return false;
    }

    public List<Map<String,String>> readOtherFile(List<Map<String,String>> list){
        for(FileInfoModel model:fileList){
            String fileName = model.getFileName();
            String colName = model.getColName();
            String valueCol = model.getValueCol();
            if("".equals(valueCol)||valueCol==null){
                valueCol = "4";
            }
            try{
                InputStream ips = new FileInputStream(uploadPath+"/"+fileName);
                //创建Excel，读取文件内容
                XSSFWorkbook workbook = new XSSFWorkbook(ips);
                //获取第一个工作表workbook.getSheet("sheet0");(方式一)
                // HSSFSheet sheet = workbook.getSheet("Sheet0");
                //读取默认第一个工作表sheet（方式二）
                XSSFSheet sheet = workbook.getSheetAt(0);
                //读取工作表中的数据
                int firstRowNum = 3;
                //获取sheet中最后一行行号
                int lastRowNum = sheet.getLastRowNum();
                for (int i = firstRowNum; i <= lastRowNum ; i++) {
                    XSSFRow row = sheet.getRow(i);
                    //获取当前行最后单元格的列号
                    int lastCellNum = row.getLastCellNum();
                    for (int j = 0; j < lastCellNum; j++) {
                        //读取cell
                        XSSFCell cell = row.getCell(j);
                        //定义value，读取值
                        cell.setCellType(CellType.STRING);
                        String value= cell.getStringCellValue();
                        if(checkChinese(value)&&j==0){
                            break;
                        }
                        if((j+"").equals(valueCol)){
                            String id = row.getCell(0).getStringCellValue();
                            for(Map<String,String> map : list){
                                if(map.get("productId").equals(id)){
                                    map.put(colName,value);
                                    break;
                                }
                            }

                        }

                    }
                }

            }catch (IOException e){
                e.printStackTrace();
            }

        }
        for(FileInfoModel model:fileList){
            String colName = model.getColName();
            for(Map<String,String> map : list){
                if(map.get(colName)==null){
                    map.put(colName,"0");
                }
            }
        }

        System.out.println(list);

        return list;
    }

    public File createExcel(List<Map<String,String>> list){
        List<String> nameList = new ArrayList<String>();
        List<String> titleList = new ArrayList<String>();
        titleList.add("货品编码");
        titleList.add("货品名称");
        titleList.add("规格");
        titleList.add("单位");
        for(FileInfoModel model : fileList){
            titleList.add(model.getColName());
        }
        File destFile= new File(downloadPath);
        if(!destFile.exists()) {
            destFile.mkdirs();
        }
        String excelName = downloadPath +"/total.xlsx";
        File excel = new File(excelName);
        Workbook workbook = null;
        FileOutputStream fos = null;
        Sheet sheet = null;

        try {
            if(excel.exists()){
                excel.delete();
            }

            excel.createNewFile();
            //新建excel
            workbook = new SXSSFWorkbook();
            //新建sheet
            sheet = workbook.createSheet("sheet1");
            //创建行（标题行） 0代表第一行
            Row rowTitle = sheet.createRow(0);
            //定义单元格
            Cell cell1 = null;
            int k = 0;
            for(String titleName : titleList){
                //设置单元格类型
                cell1 = rowTitle.createCell(k, Cell.CELL_TYPE_STRING);
                //设置单元格的值
                cell1.setCellValue(titleName);
                k++;
            }

            Row row = null;
            Cell cell = null;
            String cellValueString = null;
            //循环写入数据

            for(int i=0;i<list.size();i++){
                row = sheet.createRow(i+1);
                if(list.get(i) == null){
                    continue;
                }
                int j = 0;
                for(String key:list.get(i).keySet()){
                    cellValueString = String.valueOf(list.get(i).get(key));
                    cell = row.createCell(j, Cell.CELL_TYPE_STRING);
                    cell.setCellValue(cellValueString);
                    j++;
                }
            }
            row = sheet.createRow(list.size()+1);
            cellValueString = String.valueOf("总计");
            cell = row.createCell(0, Cell.CELL_TYPE_STRING);
            cell.setCellValue(cellValueString);
            int i = 4;
            for(FileInfoModel model:fileList){
                String colName = model.getColName();
                int total = 0;
                for(Map<String,String> map : list){
                    total = total + Integer.valueOf(map.get(colName));
                }
                cell = row.createCell(i, Cell.CELL_TYPE_STRING);
                cell.setCellValue(total);
                i++;
            }
            fos = new FileOutputStream(excel);
            workbook.write(fos);
            System.out.print("结束"+"\n");
            nameList.add(excelName);
        } catch(Exception e){
            excel.delete();
            System.out.println("文件操作失败");
            e.printStackTrace();
        } finally {
            try{
                fos.close();
                workbook.close();
            }
            catch(Exception e2){
                e2.printStackTrace();
            }
        }
        return excel;
    }


}
