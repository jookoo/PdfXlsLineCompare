import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
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

	public static final DateTimeFormatter DTFXLS = DateTimeFormatter.ofPattern("dd-MMM-uuuu");
	
	public static final DateTimeFormatter DTFL= DateTimeFormatter.ofPattern("dd.MM.yyyy");
	
	public static final SimpleDateFormat SDF = new SimpleDateFormat("dd.MM.yyyy");
	
	final Set<Line> pdflines = new LinkedHashSet<>();
	
	final Set<Line> xmllines = new LinkedHashSet<>();
	
	final Set<Integer> relevantMonths = new LinkedHashSet<>();
	
	final Calendar tempcal = GregorianCalendar.getInstance();
	
	final boolean inverted = true;
	
	public static void main(String[] args) throws IOException, ParseException {
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
		System.out.println(cu.relevantMonths);
		cu.compose(cu.inverted);
		cu.printOutput();
		
	}
	
	private void compose(final boolean inverted) throws ParseException {
		if (inverted) {
			for (Line fl: xmllines) {
				for (Line line: pdflines) {
					if (fl.number.equals(line.number)) {
						fl.isin = true;
					}
				}
			}
			
		} else {
			for (Line fl: pdflines) {
				for (Line line: xmllines) {
					if (fl.number.equals(line.number)) {
						fl.isin = true;
					}
				}
			}
		}
		
	}

	private void readPdf(final String path) throws IOException, ParseException {
		PDDocument pddDocument=PDDocument.load(new File(path));
		PDFTextStripper textStripper=new PDFTextStripper();
		final String text = textStripper.getText(pddDocument);
		pdflines.addAll(new LineParser("(\\d{2}\\.\\d{2}\\.\\d{4})\\s\\d{4}\\s(\\d{10})\\s.+").parse(text));
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
		                final Cell c0 = row.getCell(cellNumber);
		                final Cell c2 = row.getCell(2);
		                if (null != c0 && null != c2 && Cell.CELL_TYPE_NUMERIC == c2.getCellType()) {
			                c0.setCellType(Cell.CELL_TYPE_STRING);
			                final Date d = c2.getDateCellValue();
			                final String str = SDF.format(d);
			                final Line l = new Line(c0.getStringCellValue(), str);
			                xmllines.add(l);
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
	
	public void printOutput() throws ParseException {
		final Calendar cal = GregorianCalendar.getInstance();
		int count = 0;
		if (inverted) {
			System.out.println("Fehlt in der PDF (Monatsabhängig):");
			for (Line fl: xmllines) {
				cal.setTime(SDF.parse(fl.dat));
				final int month = cal.get(Calendar.MONTH) + 1;
                final int year = cal.get(Calendar.YEAR);
                final int myear = month * 10000 + year;
				if (relevantMonths.contains(myear)) {
				if (!fl.isin) {
					System.out.println("FEHLT " + fl.toString());
					count++;
				} else {
//				System.out.println(fl.toString());
				}
				}
			}
		} else {
			System.out.println("Fehlt in der XLS:");
			for (Line fl: pdflines) {
				if (!fl.isin) {
					System.out.println("FEHLT " + fl.toString());
					count++;
				} else {
//				System.out.println(fl.toString());
				}
				
			}
		}
		System.out.println(count);
	}
	
	private class LineParser {
		
		private final String pattern;
		
		public LineParser(final String pattern) {
			this.pattern = pattern;
		}
		
		public Set<Line> parse(final String text) throws ParseException {
			final Set<Line> x = new LinkedHashSet<>();
			final Pattern p = Pattern.compile(pattern);
			final Matcher m = p.matcher(text);
			while (m.find()) {
				final String dat = m.group(1);
				final String nr = m.group(2);
				final Date d = SDF.parse(dat);
				tempcal.setTime(d);
                final int month = tempcal.get(Calendar.MONTH) + 1;
                final int year = tempcal.get(Calendar.YEAR);
                final int myear = month * 10000 + year;
                if (!relevantMonths.contains(myear)) {
                	relevantMonths.add(myear);
                }
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
