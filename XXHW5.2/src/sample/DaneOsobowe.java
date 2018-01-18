package sample;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javafx.event.ActionEvent;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;

public class DaneOsobowe implements HierarchicalController<MainController> {

	public TextField imie;
	public TextField nazwisko;
	public TextField pesel;
	public TextField indeks;
	public TableView<Student> tabelka;
	private MainController parentController;

	public void dodaj(ActionEvent actionEvent) {
		Student st = new Student();
		st.setName(imie.getText());
		st.setSurname(nazwisko.getText());
		st.setPesel(pesel.getText());
		st.setIdx(indeks.getText());
		tabelka.getItems().add(st);
	}

	public void setParentController(MainController parentController) {
		this.parentController = parentController;
		// tabelka.getItems().addAll(parentController.getDataContainer().getStudents());
		tabelka.setItems(parentController.getDataContainer().getStudents());
	}

	public void usunZmiany() {
		tabelka.getItems().clear();
		tabelka.getItems().addAll(parentController.getDataContainer().getStudents());
	}

	public MainController getParentController() {
		return parentController;
	}

	public void initialize() {
		for (TableColumn<Student, ?> studentTableColumn : tabelka.getColumns()) {
			if ("imie".equals(studentTableColumn.getId())) {
				studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("name"));
			} else if ("nazwisko".equals(studentTableColumn.getId())) {
				studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("surname"));
			} else if ("pesel".equals(studentTableColumn.getId())) {
				studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("pesel"));
			} else if ("indeks".equals(studentTableColumn.getId())) {
				studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("idx"));
			}
		}

	}

	public void zapisz(ActionEvent actionEvent) {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Studenci");

		// Create heading
		Row rowHeading = sheet.createRow(0);
		rowHeading.createCell(0).setCellValue("Imie");
		rowHeading.createCell(1).setCellValue("Nazwisko");
		rowHeading.createCell(2).setCellValue("Ocena");
		rowHeading.createCell(3).setCellValue("Uzasadnienie");
		rowHeading.createCell(4).setCellValue("Nr indeksu");
		rowHeading.createCell(5).setCellValue("PESEL");

		CellStyle styleRowHeading = wb.createCellStyle();
		Font font = wb.createFont();
		font.setBold(true);
		font.setFontName(HSSFFont.FONT_ARIAL);
		font.setFontHeightInPoints((short) 11);
		styleRowHeading.setFont(font);
		styleRowHeading.setVerticalAlignment(VerticalAlignment.CENTER);
		for (int i = 0; i < 6; i++) {
			rowHeading.getCell(i).setCellStyle(styleRowHeading);
		}

		// Autofit
		for (int i = 0; i < 6; i++) {
			sheet.autoSizeColumn(i);
		}

		CellStyle style1 = wb.createCellStyle();
		style1.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		CellStyle style2 = wb.createCellStyle();
		style2.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		CellStyle style3 = wb.createCellStyle();
		style3.setFillForegroundColor(IndexedColors.RED.getIndex());
		style3.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		int row = 1;
		for (Student student : tabelka.getItems()) {
			CellStyle currentStyle = style1;
			XSSFRow r = sheet.createRow(row);
			if (student.getGrade() != null) {
				r.createCell(2).setCellValue(student.getGrade());
				if (student.getGrade() >= 3.0) {
					currentStyle = style2;
				} else {
					currentStyle = style3;
				}
				r.getCell(2).setCellStyle(currentStyle);
			}

			r.createCell(0).setCellValue(student.getName());
			r.getCell(0).setCellStyle(currentStyle);
			r.createCell(1).setCellValue(student.getSurname());
			r.getCell(1).setCellStyle(currentStyle);
			r.createCell(3).setCellValue(student.getGradeDetailed());
			r.getCell(3).setCellStyle(currentStyle);
			r.createCell(4).setCellValue(student.getIdx());
			r.getCell(4).setCellStyle(currentStyle);
			r.createCell(5).setCellValue(student.getPesel());
			r.getCell(5).setCellStyle(currentStyle);
			r.setRowStyle(currentStyle);
			row++;
		}
		try (FileOutputStream fos = new FileOutputStream("data.xlsx")) {
			wb.write(fos);
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * Uwaga na serializację:
	 * https://sekurak.pl/java-vs-deserializacja-niezaufanych-danych-i-zdalne-wykonanie-kodu-czesc-i/
	 */
	public void wczytaj(ActionEvent actionEvent) {
		ArrayList<Student> studentsList = new ArrayList<>();
		try (FileInputStream ois = new FileInputStream("data.xlsx")) {
			XSSFWorkbook wb = new XSSFWorkbook(ois);
			XSSFSheet sheet = wb.getSheet("Studenci");
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				XSSFRow row = sheet.getRow(i);
				Student student = new Student();
				student.setName(row.getCell(0).getStringCellValue());
				student.setSurname(row.getCell(1).getStringCellValue());
				if (row.getCell(2) != null) {
					student.setGrade(row.getCell(2).getNumericCellValue());
				}
				student.setGradeDetailed(row.getCell(3).getStringCellValue());
				student.setIdx(row.getCell(4).getStringCellValue());
				student.setPesel(row.getCell(5).getStringCellValue());
				studentsList.add(student);
			}
			tabelka.getItems().clear();
			tabelka.getItems().addAll(studentsList);
			ois.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
