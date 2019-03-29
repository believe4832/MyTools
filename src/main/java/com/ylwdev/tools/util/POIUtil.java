package com.ylwdev.tools.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;  
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;  
import java.util.List;  
  
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileItemFactory;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

/** 
 * excel读写工具类 
 * 处理从客户端上传到服务器的Excel文件
 * */  
public class POIUtil {
    private static Logger logger  = Logger.getLogger(POIUtil.class);
    private final static String xls = "xls";  
    private final static String xlsx = "xlsx"; 
//    private final static String dataTableSpace = "NPAY_DATA";
//    private final static String indexTableSpace = "NPAY_IDX";
    private final static String dataTableSpace = "NPAYSITDATA";
    private final static String indexTableSpace = "NPAYSITDATA";
//    
    public static void main(String[] args) throws IOException {
		File[] file = new File("D:\\SQL-Creater\\weiTuoZhiFu").listFiles();
		for(int i=0;i< file.length;i++){
			if(file[i].getName().startsWith("~"))
				continue;
			getSQL(file[i].toString());
		}
	}
    
    public static void getSQL(String filePath) throws IOException {
    	//String filePath = "D:\\test\\T_EP_UN_TRUSTPAY_SIGN_INFO.xlsx";
    	//String filePath = "D:\\test\\T_EP_UN_TRUSTPAY_SIGN_SEQ.xlsx";
    	//String filePath = "D:\\test\\T_EP_MSG_SEND.xlsx";
    	List<String[]> resultList = readExcel(filePath);
    	StringBuffer tableBuf = new StringBuffer();
    	StringBuffer commentBuf = new StringBuffer();
    	StringBuffer pkBuf = new StringBuffer();
    	StringBuffer indexBuf = new StringBuffer();
    	int index = 0;
    	String tableName = "TABLE";
    	String usePart = "";
    	for(int i=0;i<resultList.size();i++){
    		String[] strArray = resultList.get(i);
    		if(index == 0){
    			tableName = strArray[0];
    			if(strArray.length>1)
    				usePart = strArray[1];
    			tableBuf.append("create table ").append(tableName).append("\n").append("(").append("\n");
    			index++;
    			continue;
    		}
    		if("tableindex".equals(strArray[0])){
    			if("唯一索引".equals(strArray[1])){
    				indexBuf.append("create unique index ");
        			indexBuf.append(strArray[2]).append(" on ").append(tableName)
    				.append(" (").append(strArray[3]).append(")  tablespace ").append(indexTableSpace).append(";").append("\n");
    			}else if("普通索引".equals(strArray[1])){
    				indexBuf.append("create index ");
        			indexBuf.append(strArray[2]).append(" on ").append(tableName)
    				.append(" (").append(strArray[3]).append(")  tablespace ").append(indexTableSpace).append(";").append("\n");
    			}else if("分区索引".equals(strArray[1])){
    				indexBuf.append("create index ");
        			indexBuf.append(strArray[2]).append(" on ").append(tableName)
    				.append(" (").append(strArray[3]).append(")  local").append(";").append("\n");
    			}

    		}else if("pk".equals(strArray[0])){
    			pkBuf.append("alter table ").append(tableName);
    			pkBuf.append(" add constraint ").append(strArray[2]).append(" primary key (").append(strArray[3]).append(") using index tablespace ").append(indexTableSpace).append(";");
    		}else {
    			tableBuf.append("  ").append(strArray[0]).append(" ").append(strArray[2]);
        		if("T".equals(strArray[3])){
        			tableBuf.append(" not null,").append("\n");
        		}else{
        			tableBuf.append(",\n");
        		}
        		
        		commentBuf.append("comment on column ").append(tableName).append(".")
        		.append(strArray[0]).append(" is '").append(strArray[1]).append("';\n");
    		}
    	}
    	tableBuf = new StringBuffer(tableBuf.substring(0,tableBuf.length()-2)+"\n)");
    	if(usePart.length()>1){
    		//使用分区|list|按天分区|31|CREATE_DAY|
    		String[] partion = usePart.split("\\|");
    		tableBuf.append("\n").append("partition by list (").append(partion[4]).append(")");
    		tableBuf.append("\n").append("(").append("\n");
    		for(int j = 1; j <=Integer.parseInt(partion[3]);j++){
    			String temp = j<10?"0"+j:j+"";
    			tableBuf.append("  partition PART_").append(temp).append(" values ('").append(temp).append("') tablespace ").append(dataTableSpace).append(",").append("\n");
    		}
    		tableBuf = new StringBuffer(tableBuf.substring(0,tableBuf.length()-2)+"\n);");
    	}else{
    		tableBuf.append(" tablespace ").append(dataTableSpace).append(";");
    	}
    	
    	System.out.println(tableBuf.toString());
    	System.out.println(commentBuf.toString());
    	System.out.println(pkBuf.toString());
    	System.out.println(indexBuf.toString());
    	System.out.println("\n\n\n");
//    	create index IDX_PARTI_RANGE_ID on T_PARTITION_RANGE(id) local
//    	partition by list (CREATE_DAY)
//    	(
//    	  partition PART_01 values ('01') tablespace NPAYSITDATA,
//    	  partition PART_02 values ('02') tablespace NPAYSITDATA
//    	);
    	
    	//create table T_EP_UNION_SIGN_INFO
    	//(
    	//comment on column T_EP_CREDIT_INFO.file_name is '2323323';
    	
    	//alter table T_EP_UN_ORDER_INFO
//    	  add constraint PK_UN_ORDER_INFO primary key (ORDER_ID)
//    	  using index 
//    	  tablespace NPAY_IDX
    	
    	//create unique index INDEX_UNION_SIGN_INFO_1 on T_EP_UNION_SIGN_INFO (MERCHANT_NO, ACCT_ALIAS)  tablespace NPAY_IDX;
//    	create index UN_ORDER_INFO_CREATE_DATE on T_EP_UN_ORDER_INFO (CREATE_DATE)  tablespace NPAY_IDX

	}
      
    /** 
     * 读入excel文件，解析后返回
     * @throws IOException  
     */  
    public static List<String[]> readExcel(String filePath) throws IOException{  
        //检查文件  
        checkFile(filePath);  
        //获得Workbook工作薄对象  
        //Workbook workbook = getWorkBook(filePath); 
        Workbook workbook = null;
        File file = new File(filePath);
    	FileInputStream fileIs = new FileInputStream(file);
        //获取excel文件的io流  
        //InputStream is = file.getInputStream();  
        //根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象  
        if(filePath.endsWith(xls)){  
            //2003  
            workbook = new HSSFWorkbook(fileIs);  
        }else if(filePath.endsWith(xlsx)){  
            //2007  
            workbook = new XSSFWorkbook(fileIs);  
        }  
        //创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回  
        List<String[]> list = new ArrayList<String[]>();  
        if(workbook != null){  
            for(int sheetNum = 0;sheetNum < workbook.getNumberOfSheets();sheetNum++){  
                //获得当前sheet工作表  
                Sheet sheet = workbook.getSheetAt(sheetNum);  
                if(sheet == null){  
                    continue;  
                }  
                //获得当前sheet的开始行  
                int firstRowNum  = sheet.getFirstRowNum();  
                //获得当前sheet的结束行  
                int lastRowNum = sheet.getLastRowNum();  
                //循环除了第一行的所有行  
                for(int rowNum = firstRowNum;rowNum <= lastRowNum;rowNum++){  
                    //获得当前行  
                    Row row = sheet.getRow(rowNum);  
                    if(row == null){  
                        continue;  
                    }  
                    //获得当前行的开始列  
                    int firstCellNum = row.getFirstCellNum();  
                    //获得当前行的列数  
                    int lastCellNum = row.getPhysicalNumberOfCells();  
                    String[] cells = new String[row.getPhysicalNumberOfCells()];  
                    //循环当前行  
                    for(int cellNum = firstCellNum; cellNum < lastCellNum;cellNum++){  
                        Cell cell = row.getCell(cellNum);  
                        cells[cellNum] = getCellValue(cell);  
                    }  
                    list.add(cells);  
                }  
            }  
            //workbook;  
        }  
        fileIs.close();
        return list;  
    }  
    public static void checkFile(String filePath) throws IOException{  
        //判断文件是否存在
    	File file = new File(filePath);
        if(!file.exists()){  
            logger.error("文件不存在！");  
            throw new FileNotFoundException("文件不存在！");  
        }  
        //判断文件是否是excel文件  
        if(!filePath.endsWith(xls) && !filePath.endsWith(xlsx)){  
            logger.error(filePath + "不是excel文件");  
            throw new IOException(filePath + "不是excel文件");  
        }  
    }  
    public static Workbook getWorkBook(String filePath) {  

        //创建Workbook工作薄对象，表示整个excel  
        Workbook workbook = null;  
        try {  
        	File file = new File(filePath);
        	FileInputStream fileIs = new FileInputStream(file);
            //获取excel文件的io流  
            //InputStream is = file.getInputStream();  
            //根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象  
            if(filePath.endsWith(xls)){  
                //2003  
                workbook = new HSSFWorkbook(fileIs);  
            }else if(filePath.endsWith(xlsx)){  
                //2007  
                workbook = new XSSFWorkbook(fileIs);  
            }  
        } catch (IOException e) {  
            logger.info(e.getMessage());  
        }  
        return workbook;  
    }  
    public static String getCellValue(Cell cell){  
        String cellValue = "";  
        if(cell == null){  
            return cellValue;  
        }  
        //把数字当成String来读，避免出现1读成1.0的情况  
        if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
            cell.setCellType(Cell.CELL_TYPE_STRING);
        }  
        //判断数据的类型  
        switch (cell.getCellType()){
            case Cell.CELL_TYPE_NUMERIC: //数字
                cellValue = String.valueOf(cell.getNumericCellValue());  
                break;  
            case Cell.CELL_TYPE_STRING: //字符串  
                cellValue = String.valueOf(cell.getStringCellValue());  
                break;  
            case Cell.CELL_TYPE_BOOLEAN: //Boolean  
                cellValue = String.valueOf(cell.getBooleanCellValue());  
                break;  
            case Cell.CELL_TYPE_FORMULA: //公式  
                cellValue = String.valueOf(cell.getCellFormula());  
                break;  
            case Cell.CELL_TYPE_BLANK: //空值   
                cellValue = "";  
                break;  
            case Cell.CELL_TYPE_ERROR: //故障  
                cellValue = "非法字符";  
                break;  
            default:  
                cellValue = "未知类型";  
                break;  
        }  
        return cellValue;  
    }  
    
    public FileItem createFileItem(String filePath) {
        FileItemFactory factory = new DiskFileItemFactory(16, null);
        String textFieldName = "textField";
        int num = filePath.lastIndexOf(".");
        String extFile = filePath.substring(num);
        FileItem item = factory.createItem(textFieldName, "text/plain", true, "MyFileName");
        File newfile = new File(filePath);
        int bytesRead = 0;
        byte[] buffer = new byte[8192];
        try {
            FileInputStream fis = new FileInputStream(newfile);
            OutputStream os = item.getOutputStream();
            while ((bytesRead = fis.read(buffer, 0, 8192)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
            os.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return item;
    }
}
