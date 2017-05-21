import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Runner {
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		List<ClanMember> plogs = ActivePlogs.list;
		plogs.forEach(plog -> plog.getInformation());
		plogs.removeIf(plog -> !plog.getActive());
		
		try {
			File file = new File("src/assets/Ploggystyle Log.xlsx");
			Workbook wb = new XSSFWorkbook(new FileInputStream(file));
			Row row;
			Cell dateCell;
			
			// HM Level
			Sheet lvlSheet = wb.getSheet("HM Level");
			row = lvlSheet.getRow(0);
			int lvlCol = row.getLastCellNum();
			dateCell = row.createCell(lvlCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellInt(plog, lvlSheet, lvlCol, Integer.parseInt(plog.getItemInformation().getMasteryLevel())));
			
			// Soul
			Sheet soulSheet = wb.getSheet("Soul");
			row = soulSheet.getRow(0);
			int soulCol = row.getLastCellNum();
			dateCell = row.createCell(soulCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellString(plog, soulSheet, soulCol, plog.getItemInformation().getSoul()));

			
			// Weapon
			Sheet weaponSheet = wb.getSheet("Weapon");
			row = weaponSheet.getRow(0);
			int weaponCol = row.getLastCellNum();
			dateCell = row.createCell(weaponCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellString(plog, weaponSheet, weaponCol, plog.getItemInformation().getWeapon()));


			// Ring
			Sheet ringSheet = wb.getSheet("Ring");
			row = ringSheet.getRow(0);
			int ringCol = row.getLastCellNum();
			dateCell = row.createCell(ringCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellString(plog, ringSheet, ringCol, plog.getItemInformation().getRing()));

			// Earring
			Sheet earringSheet = wb.getSheet("Earring");
			row = earringSheet.getRow(0);
			int earringCol = row.getLastCellNum();
			dateCell = row.createCell(earringCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellString(plog, earringSheet, earringCol, plog.getItemInformation().getEarring()));


			// Necklace
			Sheet necklaceSheet = wb.getSheet("Neck");
			row = necklaceSheet.getRow(0);
			int necklaceCol = row.getLastCellNum();
			dateCell = row.createCell(necklaceCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellString(plog, necklaceSheet, necklaceCol, plog.getItemInformation().getNecklace()));


			// Bracelet
			Sheet braceletSheet = wb.getSheet("Bracelet");
			row = braceletSheet.getRow(0);
			int braceletCol = row.getLastCellNum();
			dateCell = row.createCell(braceletCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellString(plog, braceletSheet, braceletCol, plog.getItemInformation().getBracelet()));


			// Belt
			Sheet beltSheet = wb.getSheet("Belt");
			row = beltSheet.getRow(0);
			int beltCol = row.getLastCellNum();
			dateCell = row.createCell(beltCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellString(plog, beltSheet, beltCol, plog.getItemInformation().getBelt()));


			// Pet
			Sheet petSheet = wb.getSheet("Pet");
			row = petSheet.getRow(0);
			int petCol = row.getLastCellNum();
			dateCell = row.createCell(petCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellString(plog, petSheet, petCol, plog.getItemInformation().getPet()));


			// Soul Badge
			Sheet sbSheet = wb.getSheet("Soul Badge");
			row = sbSheet.getRow(0);
			int sbCol = row.getLastCellNum();
			dateCell = row.createCell(sbCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellString(plog, sbSheet, sbCol, plog.getItemInformation().getSoulBadge()));


			// Mystic Badge
			Sheet mbSheet = wb.getSheet("Mystic Badge");
			row = mbSheet.getRow(0);
			int mbCol = row.getLastCellNum();
			dateCell = row.createCell(mbCol);
			formatForDate(wb, dateCell);
			plogs.forEach(plog -> writePlogToCellString(plog, mbSheet, mbCol, plog.getItemInformation().getMysticBadge()));
			
			XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
			
			FileOutputStream out = new FileOutputStream(file);
			wb.write(out);
			wb.close();
						
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private static Row findPlogRow(ClanMember plog, Sheet sheet) {
		for(int i = 1; i <= sheet.getLastRowNum(); i ++) {
			String name = sheet.getRow(i).getCell(0).getStringCellValue();
			if(name.equals(plog.getCharacterName())) {
				return sheet.getRow(i);
			}
			if(name == null || name.isEmpty()) {
				sheet.getRow(i).getCell(0).setCellValue(plog.getCharacterName());
				return sheet.getRow(i);
			}
		}
		Row rowToReturn = sheet.createRow(sheet.getLastRowNum() + 1);
		rowToReturn.createCell(0).setCellValue(plog.getCharacterName());
		return rowToReturn;
	}
	
	private static void formatForDate(Workbook wb, Cell cell) {
		CellStyle dateCellStyle = wb.createCellStyle();
		CreationHelper createHelper = wb.getCreationHelper();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
		cell.setCellValue(new Date());
		cell.setCellStyle(dateCellStyle);
	}
	
	private static void writePlogToCellInt(ClanMember plog, Sheet sheet, int col, int info) {
		Row row = findPlogRow(plog, sheet);
		Cell cellToWrite = row.createCell(col);
		cellToWrite.setCellValue(info);
	}
	
	private static void writePlogToCellString(ClanMember plog, Sheet sheet, int col, String info) {
		Row row = findPlogRow(plog, sheet);
		Cell cellToWrite = row.createCell(col);
		cellToWrite.setCellValue(info);
	}

}
