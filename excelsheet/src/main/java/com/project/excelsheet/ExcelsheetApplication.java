package com.project.excelsheet;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.FileSystems;
import java.nio.file.StandardWatchEventKinds;
import java.nio.file.WatchEvent;
import java.nio.file.WatchKey;
import java.nio.file.WatchService;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.util.StringUtils;

public class ExcelsheetApplication {

	private static final String PATH_SLASH = "\\";

	private static String getCurrentDirectory() {
		return System.getProperty("user.dir");
	}

	private static String getParentDirectory() {
		String currentDirectory = getCurrentDirectory();
		File directory = new File(currentDirectory);
		return directory.getParent();
	}

	private static List<String> textFiles(String directory) {
		File dir = new File(directory);
		if (dir.exists()) {
			List<String> textFiles = new ArrayList<String>();
			for (File file : dir.listFiles()) {
				if (file.getName().endsWith((".csv"))) {
					textFiles.add(file.getName());
				}
			}
			return textFiles;
		}
		return null;
	}

	private static String getAttributeName(int i) {
		switch (i) {
			case 1:
				return "Date";
			case 2:
				return "LED No.";
			case 3:
				return "Product Type";
			case 4:
				return "Product Cat. Ref.";
			case 5:
				return "Co-makersName";
			case 6:
				return "Serial No.";
			case 7:
				return "Driver Make";
			case 8:
				return "LED Make";
			case 9:
				return "Switching Test";
			case 10:
				return "Life Test";
			case 11:
				return "Start Date";
			case 12:
				return "Start Month";
			case 13:
				return "Start Year";
			case 14:
				return "Start Hour";
			case 15:
				return "Start Min";
			case 16:
				return "Stop Date";
			case 17:
				return "Stop Month";
			case 18:
				return "Stop Year";
			case 19:
				return "Stop Hour";
			case 20:
				return "Stop Min";
			case 21:
				return "Set Cycles";
			case 22:
				return "Completed Cycles";
			case 23:
				return "Status";
			case 24:
				return "Tested By :";
			case 25:
				return "Approved By :";
			default:
				return "";
		}
	}

	private static List<LedStatusBean> readLedStatusFromCSV(String fileName) {
		System.out.println("Started reading csv file from = " + fileName);
		System.out.println();
		List<LedStatusBean> ledStatusBeans = new ArrayList<>();
		File pathToFile = new File(fileName);

		try (BufferedReader br = new BufferedReader(new FileReader(pathToFile))) {
			String line = br.readLine();
			String[] attributes = new String[25];
			int count = 0;
			while (line != null) {
				String[] split = line.split(",");
				if (split.length == 2) {
					String attributeName = getAttributeName(count + 1);
					if (attributeName.equals(split[0])) {
						attributes[count] = split[1];
						count++;
					} else {
						break;
					}
				} else {
					break;
				}
				line = br.readLine();
			}
			if (count < 25) {
				System.out.println("Invalid data in file = " + fileName);
				System.out.println();
			} else {
				int i = 0;
				while (i < 25) {
					if (attributes[i] == null || StringUtils.isEmpty(attributes[i])) {
						System.out.println("Invalid data in file = " + fileName);
						System.out.println();
						break;
					}
					i++;
				}
				if (i == 25) {
					System.out.println("Completed reading csv file = " + fileName);
					System.out.println();
					LedStatusBean ledStatusBean = createLedStatusBean(attributes);
					ledStatusBeans.add(ledStatusBean);
				}
			}
		} catch (IOException ioe) {
			System.out.println();
		}
		return ledStatusBeans;
	}

	private static LedStatusBean createLedStatusBean(String[] metadata) {
		LedStatusBean ledBean = new LedStatusBean(metadata);
		return ledBean;
	}

	private static List<String> readDataFromConfigFile(String configFilePath) {
		List<String> configData = new ArrayList<>();
		BufferedReader reader = null;
		try {
			String currentLine;
			StringBuilder sb = new StringBuilder();
			reader = new BufferedReader(new FileReader(configFilePath));

			while ((currentLine = reader.readLine()) != null) {
				currentLine.replace(System.getProperty("line.separator"), "");
				configData.add(currentLine);
				System.out.println(sb.toString());
			}
		} catch (IOException e) {
			System.out.println();
		} finally {
			try {
				if (reader != null) {
					reader.close();
				}
			} catch (IOException ex) {
				System.out.println();
			}
		}
		return configData;
	}

	private static void updateNewFile(String path, LedStatusBean led) {
		try {
			FileInputStream inputStream = new FileInputStream(new File(path));
			Workbook workbook = WorkbookFactory.create(inputStream);
			if (workbook.getNumberOfSheets() < 1) {
				System.out.println();
				System.out.println("Not able to create excel file in output directory. Possible reason can be content is not available in default template.");
				return;
			}
			Sheet sheet = workbook.getSheetAt(0);
			if (sheet.getLastRowNum() < 30 || sheet.getRow(2).getLastCellNum() < 7) {
				System.out.println();
				System.out.println("Not able to create excel file in output directory. Possible reason can be content is not appropriate in default template.");
				return;
			}

			sheet.getRow(2).getCell(3).setCellValue(led.getDate());
			sheet.getRow(2).getCell(7).setCellValue(led.getSerialNumber());
			sheet.getRow(3).getCell(3).setCellValue(led.getProductType());
			sheet.getRow(3).getCell(7).setCellValue(led.getDriverMake());
			sheet.getRow(4).getCell(3).setCellValue(led.getProductCategoryReference());
			sheet.getRow(4).getCell(7).setCellValue(led.getLedMake());
			sheet.getRow(5).getCell(3).setCellValue(led.getCoMakersName());
			sheet.getRow(5).getCell(7).setCellValue(led.getLedNumber());
			sheet.getRow(6).getCell(3).setCellValue(led.getLifeTest());
			sheet.getRow(6).getCell(7).setCellValue(led.getSwitchingTest());
			sheet.getRow(18).getCell(3).setCellValue(led.getStartDate());
			sheet.getRow(18).getCell(5).setCellValue(led.getEndDate());
			sheet.getRow(18).getCell(7);
			sheet.getRow(21).getCell(3).setCellValue(led.getStartHour() + ":" + led.getStartMinute());
			sheet.getRow(21).getCell(5).setCellValue(led.getEndHour() + ":" + led.getEndMinute());
			sheet.getRow(21).getCell(7);
			sheet.getRow(24).getCell(3).setCellValue(led.getSetCyles());
			sheet.getRow(24).getCell(5).setCellValue(led.getCompletedCycles());
			sheet.getRow(24).getCell(7).setCellValue(led.getStatus());
			sheet.getRow(28).getCell(3).setCellValue(led.getStatus());
			sheet.getRow(30).getCell(3).setCellValue(led.getTestedBy());
			sheet.getRow(30).getCell(7).setCellValue(led.getApprovedBy());

			inputStream.close();

			FileOutputStream outputStream = new FileOutputStream(path);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
			System.out.println("Sucessfully created excel file at = " + path);
			System.out.println();
		} catch (IOException | EncryptedDocumentException | InvalidFormatException ex) {
			System.out.println();
		}
	}

	private static boolean copyFileUsingFiles(String source, String dest) {
		File s = new File(source);
		File d = new File(dest);
		if (!s.exists()) {
			System.out.println();
			System.out.println("Default template of excel file does not exist at this location = " + d + ".");
			return false;
		}
		try {
			FileUtils.copyFile(s, d);
			return true;
		} catch (IOException e) {
			System.out.println();
			System.out.println("Error occured while creating file at = " + s);
			return false;
		}
	}

	private static void readExcelFile(String outputDirectoryPath, String exisitingFileName, LedStatusBean led) {
		String strNew = exisitingFileName.replace(".csv", "");
		String defaultTemplate = getParentDirectory() + PATH_SLASH + "essentials" + PATH_SLASH + "DefaultFile.xlsx";
		String newFilePath = outputDirectoryPath + PATH_SLASH + strNew + ".xlsx";
		boolean fileCopyStatus = copyFileUsingFiles(defaultTemplate, newFilePath);
		if (fileCopyStatus) {
			updateNewFile(newFilePath, led);
		}
	}

	private static void analyzeDirectory(String listeningDirectoryPath, String outputDirectoryPath) throws URISyntaxException {
		List<String> strings = textFiles(listeningDirectoryPath);
		int i = strings.size();
		while (i > 0) {
			String textFilePath = listeningDirectoryPath + PATH_SLASH + strings.get(i - 1);
			List<LedStatusBean> ledStatusBeans = readLedStatusFromCSV(textFilePath);

			for (LedStatusBean led : ledStatusBeans) {
				readExcelFile(outputDirectoryPath, strings.get(i - 1), led);
			}
			i--;
		}
	}

	private static void analyzeSpecificFile(String listeningDirectoryPath, String outputDirectoryPath, String filePath) throws URISyntaxException {
		String textFilePath = listeningDirectoryPath + PATH_SLASH + filePath;
		List<LedStatusBean> ledStatusBeans = readLedStatusFromCSV(textFilePath);

		for (LedStatusBean led : ledStatusBeans) {
			readExcelFile(outputDirectoryPath, filePath, led);
		}
	}

	private static boolean checkFileExistsOrNot(String path) {
		File file = new File(path);
		return file.exists();
	}

	public static void main(String[] args) throws IOException, InterruptedException, URISyntaxException, URISyntaxException {
		String DEFAULT_TEMPLATE_PATH, CONFIG_FILE_PATH, LISTENING_DIRECTORY_PATH = "", OUTPUT_DIRECTORY_PATH = "";
		final String CONFIG_FILE_NAME = PATH_SLASH + "config.txt";
		final String DEFAULT_TEMPLATE_NAME = PATH_SLASH + "essentials" + PATH_SLASH + "DefaultFile.xlsx";

		CONFIG_FILE_PATH = getParentDirectory() + CONFIG_FILE_NAME;

		if (!checkFileExistsOrNot(CONFIG_FILE_PATH)) {
			System.out.println();
			System.out.println("Configuration file is missing. Expected location of config file will be = " + CONFIG_FILE_PATH);
			return;
		}

		List<String> configData = readDataFromConfigFile(CONFIG_FILE_PATH);

		if (configData.size() == 2) {
			String splitPath1[] = configData.get(0).split("=");
			if (splitPath1.length == 2 && splitPath1[0].equals("source")) {
				LISTENING_DIRECTORY_PATH = splitPath1[1];
			} else if (splitPath1.length != 2) {
				System.out.println();
				System.out.println("Source path defined in config file at line#1 is not correct.");
				return;
			} else {
				System.out.println();
				System.out.println("Did not find source path in config file at line#1.");
				return;
			}

			String splitPath2[] = configData.get(1).split("=");
			if (splitPath2.length == 2 && splitPath2[0].equals("destination")) {
				OUTPUT_DIRECTORY_PATH = splitPath2[1];
			} else if (splitPath2.length != 2) {
				System.out.println();
				System.out.println("Destination path defined in config file at line#2 is not correct.");
				return;
			} else {
				System.out.println();
				System.out.println("Did not find destination path in config file at line#2.");
				return;
			}

			DEFAULT_TEMPLATE_PATH = getParentDirectory() + DEFAULT_TEMPLATE_NAME;

			if (!checkFileExistsOrNot(LISTENING_DIRECTORY_PATH)) {
				System.out.println();
				System.out.println("Source directory does not exist at this location = " + LISTENING_DIRECTORY_PATH);
				return;
			} else if (!checkFileExistsOrNot(DEFAULT_TEMPLATE_PATH)) {
				System.out.println();
				System.out.println("Default template of excel file does not exist at this location = " + DEFAULT_TEMPLATE_PATH);
				return;
			}

			System.out.println();
			System.out.println("Program has started finding .csv files in source directory...");
			System.out.println();
			analyzeDirectory(LISTENING_DIRECTORY_PATH, OUTPUT_DIRECTORY_PATH);

			File f = new File(LISTENING_DIRECTORY_PATH);
			WatchService watchService = FileSystems.getDefault().newWatchService();
			f.toPath().register(watchService, StandardWatchEventKinds.ENTRY_CREATE, StandardWatchEventKinds.ENTRY_MODIFY);

			WatchKey key;
			while ((key = watchService.take()) != null) {
				for (WatchEvent<?> event : key.pollEvents()) {
					String fileName = event.context().toString();
					if (fileName.endsWith(".csv")) {
						System.out.println("Event kind:" + event.kind() + ". File affected: " + event.context() + ".");
						System.out.println();
						analyzeSpecificFile(LISTENING_DIRECTORY_PATH, OUTPUT_DIRECTORY_PATH, fileName);
					}
				}
				key.reset();
			}
		} else {
			System.out.println();
			System.out.println("Misconfiguration found in configuration file.");
			System.out.println();
			System.out.println("Only 2 lines are expected in configuration file as follows : ");
			System.out.println();
			System.out.println("source= < source_path >");
			System.out.println();
			System.out.println("destination= < destination_path >");
			System.out.println();
		}
	}
}
