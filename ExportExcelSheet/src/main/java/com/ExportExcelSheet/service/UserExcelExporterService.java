package com.ExportExcelSheet.service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import javax.mail.MessagingException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ExportExcelSheet.entity.UserEntity;

public class UserExcelExporterService {

	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private List<UserEntity> listUsers;

	public UserExcelExporterService(List<UserEntity> listUsers) {
		super();
		this.listUsers = listUsers;
		workbook = new XSSFWorkbook();
	}

	private void writeHeaderRow() {
		sheet = workbook.createSheet("Users");
		Row row = sheet.createRow(0);

		CellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setFontHeight(12);
		style.setFont(font);
		
		createCell(row, 0, "User_Id", style);
		createCell(row, 1, "Name", style);
		createCell(row, 2, "Email", style);
	}

	private void createCell(Row row, int columnCount, Object value, CellStyle style) {
		sheet.autoSizeColumn(columnCount);
		Cell cell = row.createCell(columnCount);
		if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		}else {
			cell.setCellValue((String) value);
		}
		cell.setCellStyle(style);
	}
	
	private void writeDataRows() {
        int rowCount = 1;
 
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(12);
        style.setFont(font);
                 
        for (UserEntity user : listUsers) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;
             
            createCell(row, columnCount++, user.getUserId(), style);
            createCell(row, columnCount++, user.getName(), style);
            createCell(row, columnCount++, user.getEmail(), style);
             
        }
    }
	
	public byte[] export() throws IOException, MessagingException {
		writeHeaderRow();
		writeDataRows();
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		workbook.close();
		
		outputStream.close();
		byte[] excelFileAsBytes = outputStream.toByteArray();
		return excelFileAsBytes;
	}
	
}
