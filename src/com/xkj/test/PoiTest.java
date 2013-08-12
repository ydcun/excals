package com.xkj.test;

import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import sun.misc.Sort;

public class PoiTest {
	public PoiTest() {

	}

	private static String path_2003 = "d:" + File.separator + "aa.xls";
	private static String path_2007 = "d:" + File.separator + "aa.xlsx";

	public static void main(String[] args) {
		List list = new ArrayList();
		Ueser ueser = new Ueser();
		ueser.setRole("管理员");
		ueser.setName("DDDDD");
		ueser.setApplication("厦门移动");

		Ueser ueser1 = new Ueser();
		ueser1.setApplication("厦门电信");
		ueser1.setName("DDDDEEE");
		ueser1.setRole("普通职员");
		list.add(ueser);
		list.add(ueser1);
		String[] str = new String[] { "应用", "姓名", "角色" };
		try {
			create2007Excel(str, list, Ueser.class);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void create2007Excel() throws FileNotFoundException, IOException {
		// POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(new
		// File(path_2007)));
		XSSFWorkbook wb = new XSSFWorkbook();// 创建文档对象 workbook
		XSSFSheet sheet = wb.createSheet();// 创建表单sheet
		// sheet.setColumnHidden(columnIndex, hidden) 设置表单的行高
		// columnIndex第几行，从0开始
		// sheet.setColumnWidth(columnIndex, width) 设置表单的行宽
		XSSFRow row = sheet.createRow(0);// 创建一个行对象 参数表示第几行，从0开始
		// row.setRowNum(rowIndex) 设置行数
		// row.setHeightInPoints(height) 设置行高 以px为单位
		XSSFCellStyle style = wb.createCellStyle();// 创建样式对象
		XSSFFont font = wb.createFont();// 创建字体对象
		// font.setFontHeightInPoints(height) 设置字体大小 px为单位
		// font.setColor(color) 设置字体颜色
		// font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD) 设置字体粗细
		// font.setFontName("黑体"); 设置为黑体
		style.setFont(font);// 设置为字体记得要加入到样式对象中
		// style.setAlignment(XSSFCellStyle.ALIGN_CENTER_SELECTION);//水平居中对齐
		style.setBorderTop(HSSFCellStyle.BORDER_THICK);// 顶部边框粗线

		// style.setTopBorderColor(HSSFCellStyle.);// 设置为红色

		style.setBorderBottom(HSSFCellStyle.BORDER_DOUBLE);// 底部边框双线

		style.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);// 左边边框

		style.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);// 右边边框

	}

	public static void create2007Excel(Object[] head, List list, Class clas)
			throws Exception {
		// 创建文档对象 workbook
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet();// 创建表单sheet
		BeanInfo info = Introspector.getBeanInfo(Ueser.class, Object.class);
		PropertyDescriptor[] des = info.getPropertyDescriptors();
		XSSFRow rowHead = sheet.createRow(0);// 创建第一行对象
		CellStyle style = wb.createCellStyle();// 创建样式对象
		style.setBorderBottom((short) 1);// 顶部边框粗线
		style.setBorderTop((short) 1);// 顶部边框粗线
		style.setBorderLeft((short) 1);// 顶部边框粗线
		style.setBorderRight((short) 1);// 顶部边框粗线
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		// 设置表头
		for (int i = 0; i < head.length; i++) {
			XSSFCell cell = rowHead.createCell(i);
			cell.setCellValue(head[i].toString());
			cell.setCellStyle(style);
			// XSSFCellStyle cellStyle = cell.getCellStyle();
			// XSSFFont font = wb.createFont();// 创建字体对象
			// /font.setFontName("黑体"); // 设置为黑体
			// font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			// cellStyle.setFont(font);
		}
		for (int i = 1; i < list.size() + 1; i++) {
			XSSFRow rowBody = sheet.createRow(i);// 创建表体每一行对象
			for (int j = 0; j < des.length; j++) {

				Method method = des[j].getReadMethod();

				Object name = method.invoke(list.get(i - 1), null);
				rowBody.createCell(j).setCellValue(name.toString());

			}
		}
		FileOutputStream out = new FileOutputStream(new File(path_2007));
		wb.write(out);
		out.close();

	}

	/*
	 * public static void test(String path) throws Exception { POIFSFileSystem
	 * fs = new POIFSFileSystem(new FileInputStream(new File( path)));
	 * HSSFWorkbook wb = new HSSFWorkbook(fs);//
	 * 由POIFSFileSystem构建一个HSSFWorkbook文档对象 HSSFSheet sheet =
	 * wb.getSheetAt(0);// 表单 for (int i = 0; i < 10; i++) { HSSFRow row =
	 * sheet.createRow(i);// 行 // 单元格 // row.createCell(0).setCellValue(i); //
	 * row.createCell(1).setCellValue("dfdfdfd"); //
	 * row.createCell(2).setCellValue("rrrr"); HSSFCell cell =
	 * row.createCell((short) (0)); cell.setCellValue(i); HSSFCell cell2 =
	 * row.createCell((short) (1)); cell2.setCellValue("张三"); HSSFCell cell3 =
	 * row.createCell((short) (2)); cell3.setCellValue("好好好"); }
	 * FileOutputStream out = new FileOutputStream(path, true); wb.write(out);
	 * out.close(); }
	 */
	public static void test1(List list, Class clas) throws Exception {

		BeanInfo info = Introspector.getBeanInfo(Ueser.class, Object.class);
		PropertyDescriptor[] des = info.getPropertyDescriptors();
		for (Object obj : list) {
			for (PropertyDescriptor p : des) {
				Method method = p.getReadMethod();
				Object name = method.invoke(obj, null);
				System.out.println(name);

			}
		}
	}
}
