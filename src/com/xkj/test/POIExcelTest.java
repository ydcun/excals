package com.xkj.test;

import java.awt.Color;

import java.io.FileInputStream;

import java.io.FileNotFoundException;

import java.io.FileOutputStream;

import java.io.IOException;

import java.io.InputStream;

import java.util.Calendar;

import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;

import org.apache.poi.hssf.usermodel.HSSFFooter;

import org.apache.poi.hssf.usermodel.HSSFPatriarch;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;

import org.apache.poi.hssf.usermodel.HSSFShape;

import org.apache.poi.hssf.usermodel.HSSFSheet;

import org.apache.poi.hssf.usermodel.HSSFSimpleShape;

import org.apache.poi.hssf.usermodel.HSSFTextbox;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.hssf.util.HSSFColor;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.CellStyle;

import org.apache.poi.ss.usermodel.ClientAnchor;

import org.apache.poi.ss.usermodel.CreationHelper;

import org.apache.poi.ss.usermodel.DataFormat;

import org.apache.poi.ss.usermodel.DateUtil;

import org.apache.poi.ss.usermodel.Drawing;

import org.apache.poi.ss.usermodel.Font;

import org.apache.poi.ss.usermodel.IndexedColors;

import org.apache.poi.ss.usermodel.Picture;

import org.apache.poi.ss.usermodel.PrintSetup;

import org.apache.poi.ss.usermodel.RichTextString;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.ss.usermodel.WorkbookFactory;

import org.apache.poi.ss.util.*;

import org.apache.poi.ss.util.CellRangeAddress;

import org.apache.poi.ss.util.CellReference;

import org.apache.poi.util.IOUtils;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;



/**
 * 
 * @author WESTDREAM
 * 
 * @since 2010-8-7 下午10:34:03
 * 
 */

public class POIExcelTest {

	/**
	 * 
	 * @throws java.lang.Exception
	 * 
	 */

	public static final String XLS_WORKBOOK_LOCATION = "D:/workbook.xls";

	public static final String XLS_OR_XLSX_DIR = "D:/";

	public static final String XLSX_WORKBOOK_LOCATION = "D:/workbook.xlsx";

	public static final String IMAGE_LOCATION = "F:/Pictures/Picture/love2.jpg";

	public static void setUpBeforeClass() throws Exception {

	}

	public void testWriteExcel() {

		// ## 重复利用 的对象 ##//

		Workbook wb = null;

		FileOutputStream fileOut = null;

		CellStyle cellStyle = null;

		Cell cell = null;

		Font font = null;

		/**
		 * 
		 * EXCEL早期版本
		 * 
		 */

		try {

			// ## 创建早期EXCEL的Workbook ##//

			wb = new HSSFWorkbook();

			// ## 获取HSSF和XSSF的辅助类 ##//

			CreationHelper createHelper = wb.getCreationHelper();

			// ## 创建一个名为“New Sheet”的Sheet ##//

			Sheet sheet = wb.createSheet("New Sheet");

			/** 第一行 --- CELL创建，数据填充及日期格式 * */

			Row row1 = sheet.createRow(0);

			// Cell cell = row.createCell(0);

			// cell.setCellValue(1);

			// ## 在相应的位置填充数据 ##//

			row1.createCell(0).setCellValue(1);

			row1.createCell(1).setCellValue(1.2);

			row1.createCell(2).setCellValue(
					createHelper.createRichTextString("CreationHelper---字符串"));

			row1.createCell(3).setCellValue(true);

			// ## 填充日期类型的数据---未设置Cell Style ##//

			row1.createCell(4).setCellValue(new Date());

			// ## 填充日期类型的数据---已设置Cell Style ##//

			cellStyle = wb.createCellStyle();

			cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
					"yyyy年MM月dd日 hh:mm:ss"));

			// cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy
			// h:mm"));

			cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
					"yyyy-MM-dd hh:mm:ss"));

			cell = row1.createCell(5);

			cell.setCellValue(new Date());

			cell.setCellStyle(cellStyle);

			// ## 另一种创建日期的方法 ##//

			/*
			 * cell = row1.createCell(6);
			 * 
			 * cell.setCellValue(Calendar.getInstance());
			 * 
			 * cell.setCellStyle(cellStyle);
			 */

			/** 第二行 --- 数据类型 * */

			Row row2 = sheet.createRow(1);

			row2.createCell(0).setCellValue(1.1);

			row2.createCell(1).setCellValue(new Date());

			row2.createCell(2).setCellValue(Calendar.getInstance());

			row2.createCell(3).setCellValue("字符串");

			row2.createCell(4).setCellValue(true);

			// ## 错误的CELL数据格式 ##//

			row2.createCell(5).setCellType(HSSFCell.CELL_TYPE_ERROR);

			/** 第三行 --- CELL的各种对齐方式 * */

			Row row3 = sheet.createRow(2);

			row3.setHeightInPoints(30);

			// ## 水平居中,底端对齐 ##//

			createCell(wb, row3, (short) 0, XSSFCellStyle.ALIGN_CENTER,
					XSSFCellStyle.VERTICAL_BOTTOM);

			// ## 水平居中,垂直居中 ##//

			createCell(wb, row3, (short) 1,
					XSSFCellStyle.ALIGN_CENTER_SELECTION,
					XSSFCellStyle.VERTICAL_BOTTOM);

			// ## 填充 ,垂直居中 ##//

			createCell(wb, row3, (short) 2, XSSFCellStyle.ALIGN_FILL,
					XSSFCellStyle.VERTICAL_CENTER);

			// ## 左对齐,垂直居中 ##//

			createCell(wb, row3, (short) 3, XSSFCellStyle.ALIGN_GENERAL,
					XSSFCellStyle.VERTICAL_CENTER);

			// ## 左对齐,顶端对齐 ##//

			createCell(wb, row3, (short) 4, XSSFCellStyle.ALIGN_JUSTIFY,
					XSSFCellStyle.VERTICAL_JUSTIFY);

			// ## 左对齐,顶端对齐 ##//

			createCell(wb, row3, (short) 5, XSSFCellStyle.ALIGN_LEFT,
					XSSFCellStyle.VERTICAL_TOP);

			// ## 右对齐,顶端对齐 ##//

			createCell(wb, row3, (short) 6, XSSFCellStyle.ALIGN_RIGHT,
					XSSFCellStyle.VERTICAL_TOP);

			/** 第四行 --- CELL边框 * */

			Row row4 = sheet.createRow(3);

			cell = row4.createCell(1);

			cell.setCellValue(4);

			cellStyle = wb.createCellStyle();

			// ## 设置底部边框为THIN ##//

			cellStyle.setBorderBottom(CellStyle.BORDER_THIN);

			// ## 设置底部边框颜色为黑色 ##//

			cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());

			// ## 设置左边边框为THIN ##//

			cellStyle.setBorderLeft(CellStyle.BORDER_THIN);

			// ## 设置左边边框颜色为红色 ##//

			cellStyle.setLeftBorderColor(IndexedColors.RED.getIndex());

			// ## 设置右边边框为THIN ##//

			cellStyle.setBorderRight(CellStyle.BORDER_THIN);

			// ## 设置右边边框颜色为蓝色 ##//

			cellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());

			// ## 设置顶部边框为MEDIUM DASHED ##//

			cellStyle.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED);

			// ## 设置顶部边框颜色为黑色 ##//

			cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

			cell.setCellStyle(cellStyle);

			/** 第五行 --- 填充与颜色 * */

			Row row5 = sheet.createRow((short) 4);

			// ## Aqua背景 ##//

			cellStyle = wb.createCellStyle();

			cellStyle.setFillBackgroundColor(IndexedColors.AQUA.getIndex());

			// ## 设置填充模式为BIG SPOTS ##//

			cellStyle.setFillPattern(CellStyle.BIG_SPOTS);

			cell = row5.createCell((short) 1);

			cell.setCellValue("Aqua背景");

			cell.setCellStyle(cellStyle);

			// ## 橙色前景色（相对 于CELL背景） ##//

			cellStyle = wb.createCellStyle();

			cellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());

			// ## 设置填充模式为SOLID FOREGROUND ##//

			cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

			cell = row5.createCell((short) 2);

			cell.setCellValue("橙色前景色");

			cell.setCellStyle(cellStyle);

			/** 第六行 --- 合并单元格 * */

			Row row6 = sheet.createRow((short) 5);

			cell = row6.createCell((short) 4);

			cell.setCellValue("合并单元格测试");

			// ## Wrong:EXCEL 2007中打开workbook.xls文件看不到"合并单元格测试"，但单元格已经合并了 ##//

			/*
			 * sheet.addMergedRegion(new CellRangeAddress(
			 * 
			 * 3, //first row (0-based)
			 * 
			 * 5, //last row (0-based)
			 * 
			 * 4, //first column (0-based)
			 * 
			 * 6 //last column (0-based)
			 * 
			 * ));
			 */

			// ## 正确合并单元格 注意：与上不同的是first row=last row ##//
			sheet.addMergedRegion(new CellRangeAddress(

			5, // first row (0-based)

					5, // last row (0-based)

					4, // first column (0-based)

					6 // last column (0-based)

					));

			/** 第七行 --- 字体 * */

			Row row7 = sheet.createRow(6);

			// ## 创建字体 ##//

			// 注意：POI限制一个Workbook创建的Font对象最多为32767，所以不要为每个CELL创建一个字体，建议重用字体

			font = wb.createFont();

			// ## 设置字体大小为24 ##//

			font.setFontHeightInPoints((short) 24);

			// ## 设置字体样式为华文隶书 ##//

			font.setFontName("华文隶书");

			// ## 斜体 ##//

			font.setItalic(true);

			// ## 添加删除线 ##//

			font.setStrikeout(true);

			// ## 将字体添加到样式中 ##//

			cellStyle = wb.createCellStyle();

			cellStyle.setFont(font);

			cell = row7.createCell(1);

			cell.setCellValue("字体测试");

			cell.setCellStyle(cellStyle);

			/** 第八行 --- 自定义颜色 * */

			Row row8 = sheet.createRow(7);

			cell = row8.createCell(0);

			cell.setCellValue("自定义颜色测试");

			cellStyle = wb.createCellStyle();

			// ## 设置填充前景色为LIME ##//

			cellStyle.setFillForegroundColor(HSSFColor.LIME.index);

			// ## 设置填充模式为SOLID FOREGROUND ##//

			cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

			font = wb.createFont();

			// ## 设置字体颜色为红色 ##//

			font.setColor(HSSFColor.RED.index);

			cellStyle.setFont(font);

			cell.setCellStyle(cellStyle);

			/*
			 * cell.setCellValue("自定义颜色测试Palette");
			 * 
			 * //creating a custom palette for the workbook
			 * 
			 * HSSFPalette palette = ((HSSFWorkbook)wb).getCustomPalette();
			 * 
			 * //replacing the standard red with freebsd.org red
			 * 
			 * palette.setColorAtIndex(HSSFColor.RED.index,
			 * 
			 * (byte) 153, //RGB red (0-255)
			 * 
			 * (byte) 0, //RGB green
			 * 
			 * (byte) 0 //RGB blue
			 *  );
			 * 
			 * //replacing lime with freebsd.org gold
			 * 
			 * palette.setColorAtIndex(HSSFColor.LIME.index, (byte) 255, (byte)
			 * 204, (byte) 102);
			 */

			/** 第九行 --- 换行 * */

			Row row9 = sheet.createRow(8);

			cell = row9.createCell(2);

			cell.setCellValue("使用 /n及Word-wrap创建一个新行");

			cellStyle = wb.createCellStyle();

			// ## 设置WrapText为true ##//

			cellStyle.setWrapText(true);

			cell.setCellStyle(cellStyle);

			// ## 设置行的高度以适应新行 ---两行##//

			row9.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));

			// ## 调整列宽 ##//

			sheet.autoSizeColumn(2);

			/** 第十行 --- 数据格式 * */

			DataFormat format = wb.createDataFormat();

			Row row10 = sheet.createRow(9);

			cell = row10.createCell(0);

			cell.setCellValue(11111.25);

			cellStyle = wb.createCellStyle();

			// ## 一位小数 ##//

			cellStyle.setDataFormat(format.getFormat("0.0"));

			cell.setCellStyle(cellStyle);

			cell = row10.createCell(1);

			cell.setCellValue(11111.25);

			cellStyle = wb.createCellStyle();

			// ## 四位小数，千位逗号隔开 ##//

			// #,###.0000效果一样

			cellStyle.setDataFormat(format.getFormat("#,##0.0000"));

			cell.setCellStyle(cellStyle);

			// ## 将文件写到硬盘上 ##//

			fileOut = new FileOutputStream(XLS_WORKBOOK_LOCATION);

			wb.write(fileOut);

			fileOut.close();

		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

		/**
		 * 
		 * EXCEL 2007及以后
		 * 
		 */

		/*
		 * try {
		 * 
		 * wb = new XSSFWorkbook();
		 * 
		 * wb.createSheet("sheet1");
		 * 
		 * Cell cell = row.createCell( 0);
		 * 
		 * cell.setCellValue("custom XSSF colors");
		 * 
		 * CellStyle style1 = wb.createCellStyle();
		 * 
		 * style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(128,
		 * 0, 128)));
		 * 
		 * style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
		 * 
		 * fileOut = new FileOutputStream("d:/workbook.xlsx");
		 * 
		 * wb.write(fileOut);
		 * 
		 * fileOut.close();
		 *  } catch (FileNotFoundException e) {
		 * 
		 * e.printStackTrace();
		 *  } catch (IOException e) {
		 * 
		 * e.printStackTrace();
		 *  }
		 */

	}

	/**
	 * 
	 * 创建相应格式的CELL
	 * 
	 */

	public void createCell(Workbook wb, Row row, short column, short halign,
			short valign) {

		Cell cell = row.createCell(column);

		// ## 给CELL赋值 ##//

		cell.setCellValue("对齐排列");

		CellStyle cellStyle = wb.createCellStyle();

		// ## 设置水平对齐方式 ##//

		cellStyle.setAlignment(halign);

		// ## 设置垂直对齐方式 ##//

		cellStyle.setVerticalAlignment(valign);

		// ## 添加CELL样式 ##//

		cell.setCellStyle(cellStyle);

	}

	/**
	 * 
	 * 测试POI EXCEL迭代和或CELL中的值
	 * 
	 */

	public void testExcelIteratorAndCellContents() {

		try {

			// ## 创建HSSFWorkbook实例 ##//

			Workbook wb = new HSSFWorkbook(new FileInputStream(
					XLS_WORKBOOK_LOCATION));

			// ## 获得第一个SHEET ##//

			Sheet sheet = wb.getSheetAt(0); // or we could cast into
											// HSSFSheet,that doesn't matter

			/** 第一种迭代方法 * */

			/*
			 * 
			 * //## 迭代ROW ##//
			 * 
			 * for (Iterator rit = sheet.rowIterator(); rit.hasNext(); ) {
			 * 
			 * Row row = rit.next();
			 * 
			 * //## 迭代CELL ##//
			 * 
			 * for (Iterator cit = row.cellIterator(); cit.hasNext(); ) {
			 * 
			 * Cell cell = cit.next();
			 * 
			 * System.out.println(cell);
			 *  }
			 *  }
			 * 
			 */

			/** 第二种迭代方法 * */

			for (Row row : sheet) {

				for (Cell cell : row) {

					// ## 获取CellReference对象 ##/

					CellReference cellRef = new CellReference(row.getRowNum(),
							cell.getColumnIndex());

					System.out.print(cellRef.formatAsString());

					System.out.print(" - ");

					// ## 根据CELL值类型进行相应处理 ##/

					switch (cell.getCellType()) {

					case Cell.CELL_TYPE_STRING:

						System.out.println(cell.getRichStringCellValue()
								.getString());

						break;

					case Cell.CELL_TYPE_NUMERIC:

						// ## yyyy年mm月dd日 hh:mm:ss此种格式日期不能识别 ##//

						// ## mm/dd/yyyy h:mm,yyyy-MM-dd
						// hh:mm:ss可以识别,估计是POI对中文日期支持不怎么好的问题 ##//

						if (DateUtil.isCellDateFormatted(cell)) {

							System.out.println(cell.getDateCellValue());

						} else {

							System.out.println(cell.getNumericCellValue());

						}

						break;

					case Cell.CELL_TYPE_BOOLEAN:

						System.out.println(cell.getBooleanCellValue());

						break;

					case Cell.CELL_TYPE_FORMULA:

						System.out.println(cell.getCellFormula());

						break;

					case Cell.CELL_TYPE_ERROR:

						System.out.println(cell.getErrorCellValue());

						break;

					default:

						System.out.println();

					}

				}

			}

		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

	}

	/**
	 * 
	 * 修改文件测试
	 * 
	 */

	public void testReadingAndRewritingWorkbooks() {

		InputStream inp = null;

		try {

			inp = new FileInputStream(XLS_WORKBOOK_LOCATION);

			// inp = new FileInputStream("workbook.xlsx");

			// ## 获得要修改的Workbook ##/

			Workbook wb = WorkbookFactory.create(inp);

			// ## 获取要修改的Sheet ##//

			Sheet sheet = wb.getSheetAt(0);

			// ## 获取要修改的Row ##//

			Row row = sheet.getRow(1);

			// ## 获取要修改的Cell，如果没有相应位置的Cell那么就创建一个 ##//

			Cell cell = row.getCell(2);

			if (cell == null)

				cell = row.createCell(2);

			// ## 写入修改数据 ##//

			cell.setCellType(Cell.CELL_TYPE_STRING);

			cell.setCellValue("修改文件测试");

			// ## 将文件写到硬盘上 ##//

			FileOutputStream fileOut = new FileOutputStream(
					XLS_WORKBOOK_LOCATION);

			wb.write(fileOut);

			fileOut.close();

		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (InvalidFormatException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

	}

	/**
	 * 
	 * 暂时没看到有什么区别
	 * 
	 */

	public void testFitSheetToOnePage() {

		try {

			Workbook wb = new HSSFWorkbook();

			Sheet sheet = wb.createSheet("format sheet");

			PrintSetup ps = sheet.getPrintSetup();

			sheet.setAutobreaks(true);

			ps.setFitHeight((short) 1);

			ps.setFitWidth((short) 1);

			// Create various cells and rows for spreadsheet.

			FileOutputStream fileOut = new FileOutputStream(
					XLS_WORKBOOK_LOCATION);

			wb.write(fileOut);

			fileOut.close();

		} catch (Exception e) {

			e.printStackTrace();

		}

	}

	/**
	 * 
	 * 设置打印区域测试
	 * 
	 */

	 
	public void testSetPrintArea() {

		/**
		 * 
		 * 注意：我测试的时候用的是EXCEL 2007打开的，效果不明显，只能控制列且列好像也是不正确的。
		 * 
		 * 但是我用EXCEL 2007转换了一下，xls，xlsx的都正确了，目前还不知道是什么问题。
		 * 
		 */

		try {

			Workbook wb = new HSSFWorkbook();

			Sheet sheet = wb.createSheet("Print Area Sheet");

			Row row = sheet.createRow(0);

			row.createCell(0).setCellValue("第一个单元格");

			row.createCell(1).setCellValue("第二个单元格");

			row.createCell(2).setCellValue("第三个单元格");

			row = sheet.createRow(1);

			row.createCell(0).setCellValue("第四个单元格");

			row.createCell(1).setCellValue("第五个单元格");

			row = sheet.createRow(2);

			row.createCell(0).setCellValue("第六个单元格");

			row.createCell(1).setCellValue("第七个单元格");

			// ## 设置打印区域 A1--C2 ##//

			// wb.setPrintArea(0, "$A$1:$C$2");

			// ## 或者使用以下方法设置 ##//

			wb.setPrintArea(

			0, // Sheet页

					0, // 开始列

					2, // 结束列

					0, // 开始行

					1 // 结束行

					);

			FileOutputStream fileOut = new FileOutputStream(
					XLS_WORKBOOK_LOCATION);

			wb.write(fileOut);

			fileOut.close();

		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

	}

	/**
	 * 
	 * 设置页脚测试
	 * 
	 * 用“页面布局”可以看到效果
	 * 
	 * 下列代码只适用xls
	 * 
	 */

	 
	public void testSetPageNumbersOnFooter() {

		try {

			HSSFWorkbook wb = new HSSFWorkbook();

			HSSFSheet sheet = wb.createSheet("Footer Test");

			// ## 获得页脚 ##/

			HSSFFooter footer = sheet.getFooter();

			Row row;

			// ## 将 当前页/总页数 写在右边 ##/

			footer.setRight(HSSFFooter.page() + "/" + HSSFFooter.numPages());

			for (int i = 0; i < 100; i++) {

				row = sheet.createRow(i);

				for (int j = 0; j < 20; j++) {

					row.createCell(j).setCellValue("A" + i + j);

				}

			}

			FileOutputStream fileOut = new FileOutputStream(
					XLS_WORKBOOK_LOCATION);

			wb.write(fileOut);

			fileOut.close();

		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

	}

	/**
	 * 
	 * 测试一些POI提供的比较方便的函数
	 * 
	 * 文档中有些以HSSF为前缀的类的方法以过时(e.g: HSSFSheet, HSSFCell etc.)，
	 * 
	 * 测试的时候我去掉了HSSF前缀，当然也就是现在POI推荐的接口(Sheet,Row,Cell etc.)
	 * 
	 */

	 
	public void testConvenienceFunctions() {

		try {

			Workbook wb = new HSSFWorkbook();

			Sheet sheet1 = wb.createSheet("Convenience Functions");

			// ## 设置Sheet的显示比例 这里是3/4,也就是 75% ##//

			sheet1.setZoom(3, 4);

			// ## 合并单元格 ##//

			Row row = sheet1.createRow((short) 1);

			Row row2 = sheet1.createRow((short) 2);

			Cell cell = row.createCell((short) 1);

			cell.setCellValue("合并单元格测试");

			// ## 创建合并区域 ##//

			CellRangeAddress region = new CellRangeAddress(1, (short) 1, 4,
					(short) 4);

			sheet1.addMergedRegion(region);

			// ## 设置边框及边框颜色 ##//

			final short borderMediumDashed = CellStyle.BORDER_MEDIUM_DASHED;

			RegionUtil.setBorderBottom(borderMediumDashed,

			region, sheet1, wb);

			RegionUtil.setBorderTop(borderMediumDashed,

			region, sheet1, wb);

			RegionUtil.setBorderLeft(borderMediumDashed,

			region, sheet1, wb);

			RegionUtil.setBorderRight(borderMediumDashed,

			region, sheet1, wb);

			// ## 设置底部边框的颜色 ##//

			RegionUtil.setBottomBorderColor(HSSFColor.AQUA.index, region,
					sheet1, wb);

			// ## 设置顶部边框的颜色 ##//

			RegionUtil.setTopBorderColor(HSSFColor.AQUA.index, region, sheet1,
					wb);

			// ## 设置左边边框的颜色 ##//

			RegionUtil.setLeftBorderColor(HSSFColor.AQUA.index, region, sheet1,
					wb);

			// ## 设置右边边框的颜色 ##//

			RegionUtil.setRightBorderColor(HSSFColor.AQUA.index, region,
					sheet1, wb);

			// ## CellUtil的一些用法 ##/

			CellStyle style = wb.createCellStyle();

			style.setIndention((short) 10);

			CellUtil.createCell(row, 8, "CellUtil测试", style);

			Cell cell2 = CellUtil.createCell(row2, 8, "CellUtil测试");

			// ## 设置对齐方式为居中对齐 ##//

			CellUtil.setAlignment(cell2, wb, CellStyle.ALIGN_CENTER);

			// ## 将Workbook写到硬盘上 ##//

			FileOutputStream fileOut = new FileOutputStream(
					XLS_WORKBOOK_LOCATION);

			wb.write(fileOut);

			fileOut.close();

		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

	}

	/**
	 * 
	 * 测试冻结窗格和拆分
	 * 
	 */

	 
	public void testSplitAndFreezePanes() {

		try {

			Workbook wb = new HSSFWorkbook();

			Sheet sheet1 = wb.createSheet("冻结首行Sheet");

			Sheet sheet2 = wb.createSheet("冻结首列Sheet");

			Sheet sheet3 = wb.createSheet("冻结两行两列 Sheet");

			Sheet sheet4 = wb.createSheet("拆分Sheet");

			/** 冻结窗格 * */

			/*
			 * 
			 * createFreezePane( colSplit, rowSplit, topRow, leftmostColumn )
			 * 
			 * colSplit 冻结线水平位置
			 * 
			 * rowSplit 冻结线垂直位置
			 * 
			 * topRow Top row visible in bottom pane
			 * 
			 * leftmostColumn Left column visible in right pane.
			 * 
			 */

			// ## 冻结首行 ##//
			sheet1.createFreezePane(0, 1, 0, 1);

			// ## 冻结首列 ##//

			sheet2.createFreezePane(1, 0, 1, 0);

			// ## 冻结两行两列 ##//

			sheet3.createFreezePane(2, 2);

			// ## 拆分,左下的为面板为激活状态 ##//

			sheet4.createSplitPane(2000, 2000, 0, 0, Sheet.PANE_LOWER_LEFT);

			FileOutputStream fileOut = new FileOutputStream(
					XLS_WORKBOOK_LOCATION);

			wb.write(fileOut);

			fileOut.close();

		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

	}

	/**
	 * 
	 * 测试简单图形
	 * 
	 */

	 
	public void testDrawingShapes() {

		try {

			Workbook wb = new HSSFWorkbook();

			Sheet sheet = wb.createSheet("Drawing Shapes");

			// ## 得到一个HSSFPatriarch对象,有点像画笔但是注意区别 ##//

			HSSFPatriarch patriarch = (HSSFPatriarch) sheet
					.createDrawingPatriarch();

			/*
			 * 构造器：
			 * 
			 * HSSFClientAnchor(int dx1, int dy1, int dx2, int dy2, short col1,
			 * int row1, short col2, int row2)
			 * 
			 * 描述：
			 * 
			 * 创建HSSFClientAnchor类的实例，设置该anchor的顶-左和底-右坐标(相当于锚点，也就是图像出现的位置,大小等).
			 * 
			 * Creates a new client anchor and sets the top-left and
			 * bottom-right coordinates of the anchor.
			 * 
			 * 参数：
			 * 
			 * dx1 第一个单元格的x坐标
			 * 
			 * dy1 第一个单元格的y坐标
			 * 
			 * dx2 第二个单元格的x坐标
			 * 
			 * dy2 第二个单元格的y坐标
			 * 
			 * col1 第一个单元格所在列
			 * 
			 * row1 第一个单元格所在行
			 * 
			 * col2 第二个单元格所在列
			 * 
			 * row2 第二个单元格所在行
			 * 
			 */

			HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 512, 255,
					(short) 1, 0, (short) 1, 0);

			// ## 通过HSSFClientAnchor类的对象创建HSSFSimpleShape的实例 ##//

			HSSFSimpleShape shape = patriarch.createSimpleShape(anchor);

			// ## 画个椭圆 ##//

			shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);

			// ## 这几个是没问题的 ##//

			// shape.setLineStyleColor(10,10,10);

			// shape.setFillColor(90,10,200);

			// shape.setLineStyleColor(HSSFColor.BLUE.index); //设置不了,尚不知原因

			// ## 设置线条颜色为红色 ##//

			// shape.setLineStyleColor(Color.BLUE.getRGB()); //搞不清楚为什是反的BLUE:红色
			// RED:蓝色，是不是开发POI的有点色盲，JUST KIDDING!

			// ## 设置填充颜色为灰色 ##//

			shape.setFillColor(Color.GRAY.getRGB()); // 这个又可以

			// ## 设置线条宽度为3pt ##//

			shape.setLineWidth(HSSFShape.LINEWIDTH_ONE_PT * 3);

			// ## 设置线条的样式为点式 ##//

			shape.setLineStyle(HSSFShape.LINESTYLE_DOTSYS);

			// ## 创建文本框并填充文字 “创建文本框” ##//

			HSSFTextbox textbox = patriarch.createTextbox(

			new HSSFClientAnchor(0, 0, 0, 0, (short) 1, 1, (short) 2, 2));

			RichTextString text = new HSSFRichTextString("创建文本框");

			// ## 创建字体 ##//

			Font font = wb.createFont();

			// ## 斜体 ##//

			font.setItalic(true);

			// ## 设置字体颜色为蓝色 ##//

			// font.setColor((short)Color.BLUE.getBlue()); not work

			font.setColor(HSSFColor.BLUE.index);

			// ## 添加字体 ##//

			text.applyFont(font);

			textbox.setString(text);

			// ## 将文件写到硬盘上 ##//

			FileOutputStream fileOut = new FileOutputStream(
					XLS_WORKBOOK_LOCATION);

			wb.write(fileOut);

			fileOut.close();

		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

	}

	/**
	 * 
	 * 添加图片到工作薄测试
	 * 
	 * 已测试PNG，JPG，GIF
	 * 
	 */

	 
	public void testImages() {

		try {

			// ## 创建一个新的工作薄 ##//

			Workbook wb = new XSSFWorkbook(); // or new HSSFWorkbook();

			// ## 添加图片到该工作薄 ##//

			InputStream is = new FileInputStream(IMAGE_LOCATION);

			byte[] bytes = IOUtils.toByteArray(is);

			int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);

			is.close();

			CreationHelper helper = wb.getCreationHelper();

			// ## 创建一个名为“添加图片”的Sheet ##//

			Sheet sheet = wb.createSheet("添加图片");

			// ## 创建一个DrawingPatriarch实例 ##//

			Drawing drawing = sheet.createDrawingPatriarch();

			//## 设置图片的形状，位置等 ##// 

			ClientAnchor anchor = helper.createClientAnchor();

			//set top-left corner of the picture, 

			//subsequent call of Picture#resize() will operate relative to it 

			anchor.setCol1(3);

			anchor.setRow1(2);

			Picture pict = drawing.createPicture(anchor, pictureIdx);

			//## 自动设置图片的大小 注意：只支持PNG，JPG，GIF（BMP未测试）##// 

			pict.resize();

			//## 保存Workbook ##// 

			String file = "picture.xls";

			if (wb instanceof XSSFWorkbook)
				file += "x";

			FileOutputStream fileOut = new FileOutputStream(XLS_OR_XLSX_DIR
					+ file);

			wb.write(fileOut);

			fileOut.close();

		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

	}

}
