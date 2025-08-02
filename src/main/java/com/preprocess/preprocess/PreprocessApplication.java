package com.preprocess.preprocess;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import service.DatToExcelService;


@SpringBootApplication(scanBasePackages = {"com.preprocess.preprocess", "service"})
public class PreprocessApplication implements CommandLineRunner {

	public static void main(String[] args) {
		SpringApplication.run(PreprocessApplication.class, args);
	}

	@Autowired
	private DatToExcelService datToExcelService;

	@Override
	public void run(String... args) throws Exception {
		// Set your paths here, or read from args
		String inputPath = "C:\\Users\\merup\\OneDrive\\Desktop\\Python_Preprocess\\Input_path";
		String inputFileName = "xyz_20250722_00001.DAT";
		String outputPath = "C:\\Users\\merup\\OneDrive\\Desktop\\Python_Preprocess\\output_path";

		try {
			datToExcelService.convertDatToExcel(inputPath, inputFileName, outputPath);
			System.out.println("File converted successfully!");
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("Error during conversion: " + e.getMessage());
		}
	}
}
