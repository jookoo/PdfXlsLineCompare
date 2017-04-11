import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.regex.PatternSyntaxException;

import javax.sound.midi.SysexMessage;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.tools.PDFBox;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Liest eine PDF/XLS Datei ein.
 * Bei der XLS Datei werden alle Blätter mit Jahreszahl (2016,2017 etc) eingelesen
 * @author Josh
 *
 */
public class ConsoleUsage {

	final List<String> dataFromXLS = new LinkedList<>();
	
	final Set<Line> lines = new LinkedHashSet<>();
	
	public static void main(String[] args) throws IOException {
		final ConsoleUsage cu = new ConsoleUsage();
		if (0 == args.length) {
			System.out.println("Keine Argumente gefunden");
			return;
		}
		for (String path: args) {
			final File f = new File(path);
			if (!f.exists() || !f.isFile()) {
				System.out.println("keine Datei, " + f.getAbsolutePath() + " ignoriert.");
			}
			
			if (cu.isPdf(f.getAbsolutePath())) {
				cu.readPdf(f.getAbsolutePath());
			} else if (cu.isXls(f.getAbsolutePath())) {
				cu.readXls(f.getAbsolutePath(), 0);
			} else {
				System.out.println("falsche Endung, Datei " + f.getAbsolutePath() + " ignoriert.");
			}
			
		}
		
		cu.compose();
		cu.printOutput();
	}
	
	private void compose() {
		for (Line fl: lines) {
			for (String str: dataFromXLS) {
				if (fl.number.equals(str)) {
					fl.isin = true;
				}
			}
		}
		
	}

	private void readPdf(final String path) throws IOException {
		PDDocument pddDocument=PDDocument.load(new File(path));
		PDFTextStripper textStripper=new PDFTextStripper();
		final String text = textStripper.getText(pddDocument);
		lines.addAll(new LineParser("(\\d{2}\\.\\d{2}\\.\\d{4})\\s\\d{4}\\s(\\d{10})\\s.+").parse(text));
		pddDocument.close();
	}
	
	private void readXls(final String path, final int cellNumber) {
		try {
            FileInputStream file = new FileInputStream(
                    new File(path));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
 
            for (int idx = 0;idx < workbook.getNumberOfSheets();idx++) {
	            XSSFSheet sheet = workbook.getSheetAt(idx);
	            final String name = sheet.getSheetName();
	            if (name.matches("\\d{4}")) {
		            Iterator<Row> rowIterator = sheet.iterator();
		            while (rowIterator.hasNext()) {
		                Row row = rowIterator.next();
		                final Cell c1 = row.getCell(cellNumber);
		                if (null != c1) {
			                c1.setCellType(Cell.CELL_TYPE_STRING);
			                dataFromXLS.add(c1.getStringCellValue());
		                }
		
		                	
		            }
	            }
            }
            file.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
	}
	
	private boolean isPdf(final String path) {
		return "pdf".equals(readExt(path));
	}
	
	private boolean isXls(final String path) {
		return "xlsx".equals(readExt(path));
	}
	
	private String readExt(String path) {
		String ext = null;
		if (null != path) {
			final int idx = path.lastIndexOf(".");
			if (-1 < idx) {
				ext = path.substring(idx + 1).toLowerCase();
			}
		}
		return ext;
	}

	public ConsoleUsage() {
	}
	
	public void printOutput() {
		int count = 0;
		for (Line fl: lines) {
			if (!fl.isin) {
				System.out.println("FEHLT " + fl.toString());
				count++;
			} else {
//				System.out.println(fl.toString());
			}
			
		}
		System.out.println(count);
	}
	
	private class LineParser {
		
		private final String pattern;
		
		public LineParser(final String pattern) {
			this.pattern = pattern;
		}
		
		public Set<Line> parse(final String text) {
			final Set<Line> x = new LinkedHashSet<>();
			final Pattern p = Pattern.compile(pattern);
			final Matcher m = p.matcher(text);
			while (m.find()) {
				final String dat = m.group(1);
				final String nr = m.group(2);
				x.add(new Line(nr.trim().substring(3, nr.length()), dat));
			}
			return x;
		}
	}
	
	private class Line {
		
		public boolean isin = false;

		private final String number;
		
		private final String dat;
		
		
		
		public Line(final String number, final String dat) {
			this.number = number;
			this.dat = dat;
		}
		
		@Override
		public String toString() {
			final StringBuffer sb = new StringBuffer();
			sb.append("[");
			sb.append(Line.class.getSimpleName());
			sb.append(" number = ").append(number);
			sb.append(" dat = ").append(dat);
			sb.append("]");
			return sb.toString();
		}
	}
}
