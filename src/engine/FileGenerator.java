package engine;

import java.io.File;
import java.io.FileFilter;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.ConnectException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;

import org.apache.commons.io.FilenameUtils;
import org.jdom.JDOMException;
import org.jopendocument.dom.template.RhinoStreamTemplate;
import org.jopendocument.dom.template.TemplateException;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

public class FileGenerator {

	static String data;

	/**
	 * genera i file odt
	 * @param line
	 */
	private void generateFilesOdt(String pLine) {

		RhinoStreamTemplate template;

		try {

			String[] splittedLine = pLine.split(";");

			InputStream templateIS = getClass().getClassLoader().getResourceAsStream("template" + File.separator + "tmpLettera.odt");

			Map<String, String> dataModelMap = getDataModel(splittedLine);
			Iterator<String> keys = dataModelMap.keySet().iterator();

			template = new RhinoStreamTemplate(templateIS);

			while (keys.hasNext()) {
				String key = (String) keys.next();
				Object obj = dataModelMap.get(key);

				template.setField(key, obj);
			}

			File odtFile = new File("fileOdt" + File.separator + getData() + File.separator + dataModelMap.get("destinatarioRagSociale") + ".odt");
			template.saveAs(odtFile);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TemplateException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (JDOMException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			System.out.println("Documento File.odt creato senza errori");
		}

	}

	/**
	 * DataModel. Gestisce la mappa dei dati da passare al documento.
	 * 
	 * @param String[] line
	 * @return Map<String, String> result
	 */
	private Map<String, String> getDataModel(String[] pLine) {
		Map<String, String> result = new HashMap<String, String>();
		result.put("mittenteRagSociale", pLine[0].trim());
		result.put("mittenteCitta", pLine[1].trim());
		result.put("mittenteIndirizzo", pLine[2].trim());
		result.put("mittenteCAP", pLine[3].trim());
		result.put("mittenteProvincia", pLine[4].trim());

		result.put("destinatarioRagSociale", pLine[5].trim());
		result.put("destinatarioCitta", pLine[6].trim());
		result.put("destinatarioIndirizzo", pLine[7].trim());
		result.put("destinatarioCAP", pLine[8].trim());
		result.put("destinatarioProvincia", pLine[9].trim());

		String testo = "Questo  un testo inserito da codice...";
		result.put("testo", testo);

		String firma = "Giamps";
		result.put("firma", firma);

		result.put("data", getData());
		return result;
	}

	/**
	 * Conversione da odt to pdf
	 */
	public void convertFromOdtToPDF() {

		// Servizio OpenOffice
		OpenOfficeConnection openOfficeCOnnection = new SocketOpenOfficeConnection("127.0.0.1", 8100);
		try {
			openOfficeCOnnection.connect();

			File dir = new File("fileOdt/" + getData());
			FileFilter fileFilter = new FileFilter() {
				public boolean accept(File file) {
					String extention = FilenameUtils.getExtension(file.getName());
					return !file.isDirectory() && "odt".equalsIgnoreCase(extention);
				}
			};
			File[] files = dir.listFiles(fileFilter);

			for (File file : files) {
				System.out.println("");
				System.out.println("-----------------------------------------");
				System.out.println("file da convertire: " + FilenameUtils.getBaseName(file.getName()));
				// Get handle to Document convertor
				DocumentConverter docConverter = new OpenOfficeDocumentConverter(openOfficeCOnnection);
				// Convert the source word file to destination pdf file.
				docConverter.convert(new File("fileOdt/" + getData() + "/" + file.getName()),
						new File("filePdf/" + getData() + "/" + FilenameUtils.getBaseName(file.getName()) + ".pdf"));
				System.out.println("conversione "+ FilenameUtils.getBaseName(file.getName()) + ".pdf  --> OK");
			}

			openOfficeCOnnection.disconnect();
		} catch (ConnectException e) {

			e.printStackTrace();
		} finally {
			openOfficeCOnnection = null;
			//System.out.println("Conversione File.pdf fine");
			
		}
	}

	public void createOdt() {
		File file = new File("dati.csv");
		try {
			int counter = 0;
			Scanner scanner = new Scanner(file);
			while (scanner.hasNextLine()) {
				String line = scanner.nextLine();
				if (!line.equals("")) {
					System.out.println("-----------------------------------------");
					System.out.println(counter + " : " + line);
					generateFilesOdt(line);
					System.out.println("");
				}
				counter++;
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	
	/**Restituisce una stringa con la prima maiuscala ed il resto minuscolo.
	 * @param s
	 * @return String
	 */
	private static String capitalize(String pString) {
		if (pString.length() == 0)
			return pString;
		return pString.substring(0, 1).toUpperCase() + pString.substring(1).toLowerCase();
	}

	
	/**
	 * @return
	 */
	public static String getData() {
		if (data == null) {
			String giorno, mese, anno;
			Date oggi = new Date();
			SimpleDateFormat formatter = new SimpleDateFormat("dd");
			giorno = formatter.format(oggi);
			formatter = new SimpleDateFormat("MMMM");
			mese = formatter.format(oggi);
			formatter = new SimpleDateFormat("yyyy");
			anno = formatter.format(oggi);
			String dataOggi = giorno + " " + capitalize(mese) + " " + anno;
			data = dataOggi;
		}
		return data;
	}

}
