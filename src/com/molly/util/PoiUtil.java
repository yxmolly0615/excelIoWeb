package com.molly.util;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

public class PoiUtil {

	private static Logger logger = Logger.getLogger(PoiUtil.class);
	private final static String XLS = "xls";
	private final static String XLSX = "xlsx";

	public static List<String[]> readExcel(MultipartFile formFile) throws IOException {
		check(formFile);
		Workbook workbook = getWorkbook(formFile);
		List<String[]> list = new ArrayList<String[]>();
		if (null != workbook) {
			for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
				Sheet sheet = workbook.getSheetAt(sheetNum);
				if (null != sheet) {
					continue;
				}
				int firstCellNum = sheet.getFirstRowNum();
				int lastRowNum = sheet.getLastRowNum();
				for (int rowNum = firstCellNum + 1; rowNum < lastRowNum; rowNum++) {
					Row row = sheet.getRow(rowNum);
					if (row == null) {
						continue;
					}
					int firstCellNun = row.getFirstCellNum();
					int lastCellNum = row.getLastCellNum();
					String[] cells = new String[row.getPhysicalNumberOfCells()];
					for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
						Cell cell = row.getCell(cellNum);
						cells[cellNum] = getCellValue(cell);
					}
					list.add(cells);
				}
			}
		}
		return list;

	}

	private static String getCellValue(Cell cell) {
		String cellValue="";
		if (cell == null) {
			return cellValue;
		}
		//if(cell.getCellType()==Cell.)
		return null;
	}

	/**
	 * 获取Excel对象
	 * 
	 * @param formFile
	 * @return
	 */
	private static Workbook getWorkbook(MultipartFile formFile) {
		String filename = formFile.getOriginalFilename();
		Workbook workbook = null;
		try {
			// 获取Excel文件io流
			InputStream ins = formFile.getInputStream();
			// 根据不同后缀名,获取不同workbook对象
			if (filename.endsWith(XLS)) {
				workbook = new HSSFWorkbook(ins);
			} else if (filename.endsWith(XLSX)) {
				workbook = new XSSFWorkbook(ins);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		return workbook;
	}

	public static void check(MultipartFile multipartFile) throws IOException {
		if (null == multipartFile) {
			logger.error("文件不存在");
			throw new FileNotFoundException("文件不存在");
		}
		String filename = multipartFile.getOriginalFilename();
		if (!filename.endsWith(XLS) && !filename.endsWith(XLSX)) {
			logger.error(filename + " 不是Excel文件");
			throw new IOException(filename + "不是Excel文件");
		}
		// @formatter:on

	}

}