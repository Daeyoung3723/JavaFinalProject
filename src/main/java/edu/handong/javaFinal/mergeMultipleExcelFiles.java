package edu.handong.javaFinal;
import edu.handong.javaFinal.reader.ZipReader;

import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class mergeMultipleExcelFiles{

	private String inputPath;
	private String outputPath;
	private boolean help;

	public void run(String[] args) {
		
		Options options = createOptions();
		
		if(parseOptions(options, args)){
			if (help){
				printHelp(options);
				return;
			}
			
		}
		
		ZipReader reader = new ZipReader();

		ArrayList<String> summary1 = reader.readFileInZip(inputPath, "0001", 1);
		ArrayList<String> table_picture1 = reader.readFileInZip(inputPath, "0001", 2);
		ArrayList<String> summary2 = reader.readFileInZip(inputPath, "0002", 1);
		ArrayList<String> table_picture2 = reader.readFileInZip(inputPath, "0002", 2);
		ArrayList<String> summary3 = reader.readFileInZip(inputPath, "0003", 1);
		ArrayList<String> table_picture3 = reader.readFileInZip(inputPath, "0003", 2);
		ArrayList<String> summary4 = reader.readFileInZip(inputPath, "0004", 1);
		ArrayList<String> table_picture4 = reader.readFileInZip(inputPath, "0004", 2);
		ArrayList<String> summary5 = reader.readFileInZip(inputPath, "0005", 1);
		ArrayList<String> table_picture5 = reader.readFileInZip(inputPath, "0005", 2);
		
		try {
			boolean isEssue = false;
			for(int i = 0; i < 5; i++) {
				if(summary1.get(i).equals(summary2.get(i))) {
					
				} else {
					isEssue = true;
				}
			}
			
			if(isEssue)
			 throw new myException("0002");
		}catch(myException e) {
			System.out.println(e.getMessage());
			try {
				BufferedWriter fw = new BufferedWriter(new FileWriter("error.csv", true));
				fw.write("0002.zip");
			}catch(Exception e1) {
				e1.printStackTrace();
			}
		}
		try {
			boolean isEssue = false;
			for(int i = 0; i < 5; i++) {
				if(summary1.get(i).equals(summary3.get(i))) {
					
				} else {
					isEssue = true;
				}
			}
			
			if(isEssue)
			 throw new myException("0003");
		}catch(myException e) {
			System.out.println(e.getMessage());
			try {
				BufferedWriter fw = new BufferedWriter(new FileWriter("error.csv", true));
				fw.write("0003.zip");
			}catch(Exception e1) {
				e1.printStackTrace();
			}
		}
		try {
			boolean isEssue = false;
			for(int i = 0; i < 5; i++) {
				if(summary1.get(i).equals(summary4.get(i))) {
					
				} else {
					isEssue = true;
				}
			}
			
			if(isEssue)
			 throw new myException("0004");
		}catch(myException e) {
			System.out.println(e.getMessage());
			try {
				BufferedWriter fw = new BufferedWriter(new FileWriter("error.csv", true));
				fw.write("0004.zip");
			}catch(Exception e1) {
				e1.printStackTrace();
			}
		}
		try {
			boolean isEssue = false;
			for(int i = 0; i < 5; i++) {
				if(summary1.get(i).equals(summary5.get(i))) {
					
				} else {
					isEssue = true;
				}
			}
			
			if(isEssue)
			 throw new myException("0005");
		}catch(myException e) {
			System.out.println(e.getMessage());
			try {
				BufferedWriter fw = new BufferedWriter(new FileWriter("error.csv", true));
				fw.write("0005.zip");
			}catch(Exception e1) {
				e1.printStackTrace();
			}
		}
		
		XSSFWorkbook xworkbook = new XSSFWorkbook();
        XSSFSheet sheet = xworkbook.createSheet("merge"); // 货 矫飘(Sheet) 积己
        int columnIndex = 0;
        int rowIndex = 0;
        XSSFRow row = sheet.createRow(rowIndex++);
        XSSFCell cell = row.createCell(columnIndex++);
        cell.setCellValue("id");
        
        
        for(String str : summary1) {
        	cell = row.createCell(columnIndex++);
        	cell.setCellValue(str);
        	if(columnIndex % 8 == 0) {
        		columnIndex = 0;
        		row = sheet.createRow(rowIndex++);
        		cell = row.createCell(columnIndex++);
        		cell.setCellValue("0001");
        	}
        }
        
        int stack = 0;
        cell.setCellValue("0002");
        for(String str : summary2) {
        	if(stack < 7) {
        		stack++;
        		continue;
        	}
        	cell = row.createCell(columnIndex++);
        	cell.setCellValue(str);
        	if(columnIndex % 8 == 0) {
        		columnIndex = 0;
        		row = sheet.createRow(rowIndex++);
        		cell = row.createCell(columnIndex++);
        		cell.setCellValue("0002");
        	}
        }
        
        stack = 0;
        cell.setCellValue("0003");
        for(String str : summary3) {
        	if(stack < 7) {
        		stack++;
        		continue;
        	}
        	cell = row.createCell(columnIndex++);
        	cell.setCellValue(str);
        	if(columnIndex % 8 == 0) {
        		columnIndex = 0;
        		row = sheet.createRow(rowIndex++);
        		cell = row.createCell(columnIndex++);
        		cell.setCellValue("0003");
        	}
        }
        
        stack = 0;
        cell.setCellValue("0004");
        for(String str : summary4) {
        	if(stack < 7) {
        		stack++;
        		continue;
        	}
        	cell = row.createCell(columnIndex++);
        	cell.setCellValue(str);
        	if(columnIndex % 8 == 0) {
        		columnIndex = 0;
        		row = sheet.createRow(rowIndex++);
        		cell = row.createCell(columnIndex++);
        		cell.setCellValue("0004");
        	}
        }
        
        stack = 0;
        cell.setCellValue("0005");
        for(String str : summary5) {
        	if(stack < 7) {
        		stack++;
        		continue;
        	}
        	cell = row.createCell(columnIndex++);
        	cell.setCellValue(str);
        	if(columnIndex % 8 == 0) {
        		columnIndex = 0;
        		row = sheet.createRow(rowIndex++);
        		cell = row.createCell(columnIndex++);
        		cell.setCellValue("0005");
        	}
        }
        
        String path = outputPath.split(".xlsx")[0];
        try {
            FileOutputStream fileoutputstream = new FileOutputStream(path + "1.xlsx");
            xworkbook.write(fileoutputstream);
            fileoutputstream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        xworkbook = new XSSFWorkbook();
        sheet = xworkbook.createSheet("merge2"); // 货 矫飘(Sheet) 积己
        columnIndex = 0;
        rowIndex = 0;
        row = sheet.createRow(rowIndex++);
        cell = row.createCell(columnIndex);
        cell.setCellValue(table_picture1.get(0));
        row = sheet.createRow(rowIndex++);
        cell = row.createCell(columnIndex++);
       
        cell.setCellValue("id");
        
        stack = 0;
        for(String str : table_picture1) {
        	if(stack < 4) {
        		stack++;
        		continue;
        	}
        	cell = row.createCell(columnIndex++);
        	cell.setCellValue(str);
        	if(columnIndex % 6 == 0) {
        		columnIndex = 0;
        		row = sheet.createRow(rowIndex++);
        		cell = row.createCell(columnIndex++);
        		cell.setCellValue("0001");
        	}
        }
        
        stack = 0;
        for(String str : table_picture2) {
        	if(stack < 10) {
        		stack++;
        		continue;
        	}
        	cell = row.createCell(columnIndex++);
        	cell.setCellValue(str);
        	if(columnIndex % 6 == 0) {
        		columnIndex = 0;
        		row = sheet.createRow(rowIndex++);
        		cell = row.createCell(columnIndex++);
        		cell.setCellValue("0002");
        	}
        }
        
        stack = 0;
        for(String str : table_picture3) {
        	if(stack < 10) {
        		stack++;
        		continue;
        	}
        	cell = row.createCell(columnIndex++);
        	cell.setCellValue(str);
        	if(columnIndex % 6 == 0) {
        		columnIndex = 0;
        		row = sheet.createRow(rowIndex++);
        		cell = row.createCell(columnIndex++);
        		cell.setCellValue("0003");
        	}
        }
        
        stack = 0;
        for(String str : table_picture4) {
        	if(stack < 10) {
        		stack++;
        		continue;
        	}
        	cell = row.createCell(columnIndex++);
        	cell.setCellValue(str);
        	if(columnIndex % 6 == 0) {
        		columnIndex = 0;
        		row = sheet.createRow(rowIndex++);
        		cell = row.createCell(columnIndex++);
        		cell.setCellValue("0004");
        	}
        }
        
        stack = 0;
        for(String str : table_picture5) {
        	if(stack < 10) {
        		stack++;
        		continue;
        	}
        	cell = row.createCell(columnIndex++);
        	cell.setCellValue(str);
        	if(columnIndex % 6 == 0) {
        		columnIndex = 0;
        		row = sheet.createRow(rowIndex++);
        		cell = row.createCell(columnIndex++);
        		cell.setCellValue("0005");
        	}
        }
        
        
        path = outputPath.split(".xlsx")[0];
        try {
            FileOutputStream fileoutputstream = new FileOutputStream(path + "2.xlsx");
            xworkbook.write(fileoutputstream);
            fileoutputstream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        

		
		
	}

	private boolean parseOptions(Options options, String[] args) {
		CommandLineParser parser = new DefaultParser();

		try {

			CommandLine cmd = parser.parse(options, args);

			inputPath = cmd.getOptionValue("i");
			outputPath = cmd.getOptionValue("o");
			
			

		} catch (Exception e) {
			printHelp(options);
			return false;
		}

		return true;
	}

	private Options createOptions() {
		Options options = new Options();

		// add options by using OptionBuilder
		options.addOption(Option.builder("i").longOpt("input")
				.desc("Set an input file path")
				.hasArg()
				.argName("Input path")
				.required()
				.build());

		options.addOption(Option.builder("o").longOpt("output")
				.desc("Set an output file path")
				.hasArg()     
				.argName("Output path")
				.required()
				.build());
		
		options.addOption(Option.builder("h").longOpt("help")
		        .desc("Show a Help page")
		        .argName("Help")
		        .build());

		return options;	
	}
	
	private void printHelp(Options options) {
		HelpFormatter formatter = new HelpFormatter();
		String header = "Multiple excel files merger";
		String footer ="";
		formatter.printHelp("The first argument of the printHelp method is \"excelMerger\".", header, options, footer, true);
		
	}

}
