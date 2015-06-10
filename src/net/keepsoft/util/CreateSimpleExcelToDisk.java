package net.keepsoft.util;

import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class CreateSimpleExcelToDisk
{
	/**
	 * @功能：手工构建一个简单格式的Excel
	 */
	
	public void excelExp(String fileName,List<String> head,List<List<String>> list){
		
		HSSFWorkbook wb = new HSSFWorkbook();
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.createSheet("1");
		wb.setSheetName(0,"firstSheet",HSSFWorkbook.ENCODING_UTF_16);
		// 第四步，创建单元格，并设置值表头 设置表头居中
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		//生成标题
		createHeadCell(head,style,sheet);
		int rowNum = 1;
		for(List<String> data:list){
			HSSFRow row = sheet.createRow((short)rowNum);
			for(int i = 0;i<data.size();i++){
				createCell(i, row, data.get(i),style);
			}
			rowNum++;
		}
		try
		{
			FileOutputStream fout = new FileOutputStream(fileName);
			wb.write(fout);
			fout.close();
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}
	
	public void createCell(int num ,HSSFRow row,String value,HSSFCellStyle style){
		HSSFCell cell = row.createCell((short) num);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellType(HSSFCell.CELL_TYPE_STRING);
		cell.setCellValue(value);
		cell.setCellStyle(style);
	}
	
	public void createHeadCell(List<String> head,HSSFCellStyle style,HSSFSheet sheet){
		HSSFRow row = sheet.createRow((short)0);
		for(int i = 0;i<head.size();i++){
			HSSFCell cell = row.createCell((short)i);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			cell.setCellValue(head.get(i));	
			cell.setCellStyle(style);
		}
	}
	
	public void createDoubleCell(int num ,HSSFRow row,String value,HSSFCellStyle style){
		HSSFCell cell = row.createCell((short) num);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellType(HSSFCell.CELL_TYPE_STRING);
		cell.setCellValue(value);
		cell.setCellStyle(style);
	}
	
	
	
	
	public  void DBToExcel(List<String> head,ResultSet rs,String fileName) throws Exception
	{
		boolean ff = false;
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet();
		HSSFCellStyle style = workbook.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式
		HSSFRow row= sheet.createRow((short)0);;
		HSSFCell cell;
		ResultSetMetaData md=rs.getMetaData();
		int nColumn=md.getColumnCount();
		if(head == null ||head.size() ==0){
			//写入各个字段的名称
			for(int i=1;i<=nColumn;i++)
			{ 
				createCell(i-1, row, md.getColumnLabel(i),style);
			}
		}else{
			createHeadCell(head, style, sheet);
		}
		
		int iRow=1;
		//写入各条记录，每条记录对应Excel中的一行
		while(rs.next())
		{
			ff =true;
			row= sheet.createRow((short)iRow);;
			for(int j=1;j<=nColumn;j++)
			{ 
				createDoubleCell(j-1, row, rs.getObject(j)+"",style);	
			}
			iRow++;
		}
		if(ff){
			FileOutputStream fOut = new FileOutputStream(fileName);
			workbook.write(fOut);
			fOut.flush();
			fOut.close();
		}
//		JOptionPane.showMessageDialog(null,"导出数据成功！");
	}
	
	//给予一个list<object>对象 和一个class，导出excel
	public void classToExcel(Class c,List<Object> list,List<String> head,String fileName){
		HSSFWorkbook wb = new HSSFWorkbook();
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.createSheet("学生表一");
		wb.setSheetName(0,"firstSheet",HSSFWorkbook.ENCODING_UTF_16);
		// 第四步，创建单元格，并设置值表头 设置表头居中
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		Field[] fields = c.getDeclaredFields();
		if(head ==null ||head.size() ==0){ //如果给没有给定标题list,就以字段名称为标题
			HSSFRow row = sheet.createRow((short)0);
			for(int i = 0;i<fields.length;i++){
				HSSFCell cell = row.createCell((short)i);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				cell.setCellValue(fields[i].getName());	
				cell.setCellStyle(style);
			}
		
		}else{
			createHeadCell(head, style, sheet);
		}
		int rowNum = 1;
		int col = 0;
		for (Object obj : list) {
			HSSFRow row = sheet.createRow((short)rowNum);
			col = 0;
			for (Field field : fields) {
				try {
					field.setAccessible(true);
					createCell(col, row, field.get(obj)+"", style);
				} catch (IllegalArgumentException e) {
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					e.printStackTrace();
				}
				col++;
			}
			rowNum++;
		}
		
		try
		{
			FileOutputStream fout = new FileOutputStream(fileName);
			wb.write(fout);
			fout.close();
		}
			catch (Exception e)
			{
				e.printStackTrace();
			}
	}
	
	
}
