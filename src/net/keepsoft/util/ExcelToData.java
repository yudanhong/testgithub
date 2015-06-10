package net.keepsoft.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelToData {
	
	public static Integer INFO_LAN = 2;//信息栏
	
	public static Integer DATA_LAN = 7;//数据信息栏
	
	public static Integer DATA_NUM = 3;//数据循环个数
	
	public static Integer BIAO_TI = 3;//列名占用行数
	
	public static void main(String[] args) {
		try {
//			File
			readPcp("E://70405750-2011年佛岭水库站水库水文要素摘录表.xls",2,5,4,2);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public static void readPcp(String fileName,int infoLan,int dataLan,int dataNum,int biaoTi) throws Exception{
		 InputStream is = new FileInputStream(fileName);
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
        List list = new ArrayList();
     // 循环工作表Sheet
         HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
         if (hssfSheet == null) {
             return;
         }
         HSSFPrintSetup hff = hssfSheet.getPrintSetup();
         int[] num = hssfSheet.getRowBreaks();//获取分页符所在信息烂熟
         int lastNum = hssfSheet.getLastRowNum();//获取最后一行所在栏数
         int[][] lanSize = new int[num.length+1][2];
         for(int i = 0;i<num.length;i++){
         	if(i == 0){
         		lanSize[i] = new int[]{0,num[i]};
         	}else{
         		lanSize[i] = new int[]{num[i-1]+1,num[i]};
         	}
         }
         lanSize[lanSize.length-1] = new int[]{num[num.length-1]+1,lastNum};
         List<List<String>> listStr = new ArrayList<List<String>>();
         for(int i = 0;i<lanSize.length;i++){
         	//先数据信息表
         	HSSFRow row = hssfSheet.getRow(lanSize[i][0]+infoLan-1);
         	String year = getIntValue(row.getCell((short)2));
         	String stcd = getValue(row.getCell((short)(dataLan-0)));
         	for(int x = 0; x < dataNum; x++){
         		//拿到第一条数据的时间
         		String month =  getIntValue(hssfSheet.getRow(infoLan+biaoTi+lanSize[i][0]).getCell((short)(0+dataLan*x)));
         		String day =  getIntValue(hssfSheet.getRow(infoLan+biaoTi+lanSize[i][0]).getCell((short)(1+dataLan*x)));
         		month = nullToZero(month);
         		day = nullToZero(day);
         		for( int y = infoLan+biaoTi+lanSize[i][0]; y <= lanSize[i][1]; y++ ){
         			HSSFRow r = hssfSheet.getRow(y);
         			//测站编码,时间（2014/11/22 22:22:22）,坝上水位,蓄水量,出库流量
         				boolean isOver = false;
         				String mon =  getIntValue(r.getCell((short)(0+dataLan*x)));
	            		String d =  getIntValue(r.getCell((short)(1+dataLan*x)));
	            		String stt =  getIntValue(r.getCell((short)(2+dataLan*x)));
	            		String ett =  getIntValue(r.getCell((short)(3+dataLan*x)));
	            		
	            		if(StringUtils.isBlank(mon)&&StringUtils.isBlank(stt+"")&&StringUtils.isBlank(d)&&StringUtils.isBlank(ett+"")){
	            			break;
	            		}
	            		Integer st = Integer.parseInt( stt);
	            		Integer et = Integer.parseInt( ett);
	            		int dur = et -st;
	            		if(et<st){
	            			dur = et+24-st;
	            			isOver = true;
	            		}
	            		if(StringUtils.isNotBlank(mon)&&Integer.parseInt(mon) < Integer.parseInt(month)){
	            			break;
	            		}
	            		if(StringUtils.isBlank(mon)){
	            			mon = month;
	            		}else{
	            			month = mon;
	            		}
	            		if(StringUtils.isBlank(d)){
	            			d = day;
	            		}else{
	            			day = d;
	            		}
	            		Date date = DateUtil.setDate(Integer.parseInt(year), Integer.parseInt(mon), Integer.parseInt(d), et);
	            		if(isOver){
	            			date = DateUtil.increaseHour(date, 24);
	            		}
	            		List<String> l = new ArrayList<String>();
	            		l.add(stcd);
	            		l.add(DateUtil.format(date, "yyyy-MM-dd HH:mm:ss"));
	            		l.add(dur+"");
	            		l.add(getValue(r.getCell((short)(4+dataLan*x))));
	            		listStr .add(l);
         		}                                                                                                                                        
         	}
         }
         CreateSimpleExcelToDisk cs = new CreateSimpleExcelToDisk();
         List<String> head = new ArrayList<String>();
         head.add("站点编码");
         head.add("监测时间");
         head.add("时长");
         head.add("降雨量");
         cs.excelExp("e://xxx.xls", head, listStr);
//         FileUtils.writeLines(new File("e://70404100-2011年牛头山水库站水库水文要素摘录表.txt"), listStr);
	}
	
	@SuppressWarnings({ "rawtypes", "unused" })
	public static void readQobs(String fileName) throws Exception{
			 InputStream is = new FileInputStream(fileName);
		        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		        List list = new ArrayList();
	        // 循环工作表Sheet
	            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
	            if (hssfSheet == null) {
	                return;
	            }
	            HSSFPrintSetup hff = hssfSheet.getPrintSetup();
	            int[] num = hssfSheet.getRowBreaks();//获取分页符所在信息烂熟
	            int lastNum = hssfSheet.getLastRowNum();//获取最后一行所在栏数
	            int[][] lanSize = new int[num.length+1][2];
	            for(int i = 0;i<num.length;i++){
	            	if(i == 0){
	            		lanSize[i] = new int[]{0,num[i]};
	            	}else{
	            		lanSize[i] = new int[]{num[i-1]+1,num[i]};
	            	}
	            }
	            lanSize[lanSize.length-1] = new int[]{num[num.length-1]+1,lastNum};
	            List<String> listStr = new ArrayList<String>();
	            for(int i = 0;i<lanSize.length;i++){
	            	//先数据信息表
	            	HSSFRow row = hssfSheet.getRow(lanSize[i][0]+INFO_LAN-1);
	            	String year = getIntValue(row.getCell((short)3));
	            	String stcd = getValue(row.getCell((short)(DATA_LAN-0)));
	            	for(int x = 0; x < DATA_NUM; x++){
	            		//拿到第一条数据的时间
	            		String month =  getIntValue(hssfSheet.getRow(INFO_LAN+BIAO_TI+lanSize[i][0]).getCell((short)(0+DATA_LAN*x)));
	            		String day =  getIntValue(hssfSheet.getRow(INFO_LAN+BIAO_TI+lanSize[i][0]).getCell((short)(1+DATA_LAN*x)));
	            		month = nullToZero(month);
	            		day = nullToZero(day);
	            		for( int y = INFO_LAN+BIAO_TI+lanSize[i][0]; y <= lanSize[i][1]; y++ ){
	            			HSSFRow r = hssfSheet.getRow(y);
	            			//测站编码,时间（2014/11/22 22:22:22）,坝上水位,蓄水量,出库流量
	            			String mon =  getIntValue(hssfSheet.getRow(y).getCell((short)(0+DATA_LAN*x)));
		            		String d =  getIntValue(hssfSheet.getRow(y).getCell((short)(1+DATA_LAN*x)));
		            		String h =  getIntValue(hssfSheet.getRow(y).getCell((short)(2+DATA_LAN*x)));
		            		String m =  getIntValue(hssfSheet.getRow(y).getCell((short)(3+DATA_LAN*x)));
		            		if(StringUtils.isBlank(mon)&&StringUtils.isBlank(h)&&StringUtils.isBlank(d)&&StringUtils.isBlank(m)){
		            			break;
		            		}
		            		if(StringUtils.isNotBlank(mon)&&Integer.parseInt(mon) < Integer.parseInt(month)){
		            			break;
		            		}
		            		if(StringUtils.isBlank(mon)){
		            			mon = month;
		            		}else{
		            			month = mon;
		            		}
		            		if(StringUtils.isBlank(d)){
		            			d = day;
		            		}else{
		            			day = d;
		            		}
		            		h = nullToZero(h);
		            		m = nullToZero(m);
		            		StringBuffer sb = new StringBuffer(stcd);
		            		sb.append(","+year+"/"+mon+"/"+d+" "+h+":"+m+":00");
		            		sb.append(","+nullToZero(getValue(hssfSheet.getRow(y).getCell((short)(4+DATA_LAN*x)))));
		            		sb.append(","+nullToZero(getValue(hssfSheet.getRow(y).getCell((short)(5+DATA_LAN*x)))));
		            		sb.append(","+nullToZero(getValue(hssfSheet.getRow(y).getCell((short)(6+DATA_LAN*x)))));
		            		System.out.println(sb.toString());
		            		listStr.add(sb.toString());
	            		}                                                                                                                                        
	            	}
	            }
	            FileUtils.writeLines(new File("e://70404100-2011年牛头山水库站水库水文要素摘录表.txt"), listStr);
	}
	

	
		private static String getIntValue(HSSFCell hssfCell) {
		    if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
	            // 返回布尔类型的值
	            return String.valueOf(hssfCell.getBooleanCellValue());
	        } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
	            // 返回数值类型的值
	        	return String.valueOf((int)hssfCell.getNumericCellValue());
	        } else {
	            // 返回字符串类型的值
	            return String.valueOf(hssfCell.getStringCellValue());
	        }
				
		    }
		
	 
	   private static String getValue(HSSFCell hssfCell) {
	        if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
	            // 返回布尔类型的值
	            return String.valueOf(hssfCell.getBooleanCellValue());
	        } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
	            // 返回数值类型的值
	            return String.valueOf(hssfCell.getNumericCellValue());
	        } else {
	            // 返回字符串类型的值
	            return String.valueOf(hssfCell.getStringCellValue());
	        }
	        
	    }
	   
	   public static String nullToZero(String str){
		   return StringUtils.isNotBlank(str)?str:"0";
	   }
	   
	   
	   public static String isNull(String str){
		   return StringUtils.isNotBlank(str)?str:null;
	   }
}
