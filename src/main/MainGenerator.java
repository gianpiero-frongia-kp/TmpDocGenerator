package main;

import engine.FileGenerator;

public class MainGenerator {

	/**
	 * Main per la generazione dei doc/pdf. 
	 * Prima di eseguirlo bisogna attivare il servizio OpenOffice
	 * 
	 * path:
	 * /application/OpenOffice.org.app/Contents/MacOS/
	 * 
	 * servizio da lanciare:
	 * soffice -headless -nofirststartwizard -accept="socket,host=localhost,port=8100;urp;StarOffice.Service"&
	 * 
	 * @param args
	 */
	public static void main(String[] args) {

		FileGenerator fileGenerator = new FileGenerator();

		// Crea File.odt
		System.out.println("-----------------------------------------");
		System.out.println("-----------------------------------------");
		System.out.println("creazione File.odt in corso..");
		fileGenerator.createOdt();

		// Crea File.pdf
		System.out.println("-----------------------------------------");
		System.out.println("-----------------------------------------");
		System.out.println("conversione odt to File.pdf in corso...");
		fileGenerator.convertFromOdtToPDF();

		System.out.println("finished");

	}

}
