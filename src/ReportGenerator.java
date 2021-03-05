import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ReportGenerator{

	public static void main(String[] args) {
		
		File reportsFolder = new File("REPORTS_ENGLISH");
		if (!reportsFolder.exists() || !reportsFolder.isDirectory())
			reportsFolder.mkdir();
		
		File studentsFile = new File("students.txt");
		File englishReport = new File("report.docx");
		File page2Report = new File("Page02.docx");
		
		FileReader fr;
		try {
			fr = new FileReader(studentsFile);

			BufferedReader br = new BufferedReader(fr);  //creates a buffering character input stream  
			String line;  
			while((line = br.readLine()) != null)  
			{  
				String studentName = line.split(",")[0];
				String className = line.split(",")[1];
				
				// Check if J-XX exists inside REPORTS_ENGLISH, if not create it
				File jFile = new File(reportsFolder.getAbsolutePath() + "/" + className);
				if (!jFile.exists() || !jFile.isDirectory())
					jFile.mkdir();
				
				// Copy the English report file
				File destination = new File(jFile.getAbsolutePath() + "/" + studentName + ".docx");
				FileUtils.copyFile(englishReport, destination);
				
				// Open the file, change the name and the class name, and save it
				updateReportFile(destination, studentName, className);
				
				// check if page2 exists inside REPORTS_ENGLISH/J-XX, if not create it
				File page2 = new File(jFile.getAbsolutePath() + "/page02");
				if (!page2.exists() || !page2.isDirectory())
					page2.mkdir();
				
				// Copy the page 2 report
				destination = new File(page2.getAbsolutePath() + "/" + studentName + ".docx");
				FileUtils.copyFile(page2Report, destination);
				
				// Open the file, change the name and the class name and save it
				updateReportFile(destination, studentName, className);
			}
			
			fr.close();    //closes the stream and release the resources  

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
	
	public static void updateReportFile(File report, String studentName, String classname) throws FileNotFoundException, IOException {

		XWPFDocument doc = new XWPFDocument(new FileInputStream(report));

		for (XWPFTable tbl : doc.getTables()) {
			for (XWPFTableRow row : tbl.getRows()) {	
				for (XWPFTableCell cell : row.getTableCells()) {
					String text = cell.getText();
					if (text != null && text.contains("J-XX")) {
						cell.removeParagraph(0);
						cell.setText("         " + classname);
					} else for (XWPFParagraph p : cell.getParagraphs()) {
						for (XWPFRun r : p.getRuns()) {
							text = r.getText(0);
							if (text != null && text.contains("Student Name")) {
								text = text.replace("Student Name", studentName);
								r.setText(text, 0);
							}
						}
					}
				}
			}
		}
		
		doc.write(new FileOutputStream(report));
		doc.close();
	}
	
}
