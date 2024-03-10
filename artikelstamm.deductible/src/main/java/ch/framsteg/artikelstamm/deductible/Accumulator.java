
package ch.framsteg.artikelstamm.deductible;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.xml.XMLConstants;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.JDOMException;
import org.jdom2.Namespace;
import org.jdom2.input.SAXBuilder;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;

public class Accumulator {

	final static private String ASTAMM_PATH_PARAM = "-as";
	final static private String PUB_ZIP_PARAM = "-zip";
	final static private String PLISTE_NAME_PARAM = "-pl";
	final static private String WORKSHEET_NUMBER_PARAM = "-w";
	final static private String GTIN_COL_NUMBER_PARAM = "-gcn";
	final static private String PERCENTAGE_COL_NUMBER_PARAM = "-pcn";
	final static private String PERCENTAGE_PARAM = "-p";
	final static private String VERBOSE_PARAM = "-v";
	final static private String PHAR_ERR_MSG = "Missing parameters. The harvester needs three parameters to operate. Aborting...";
	final static private String EMPTY_PARAM_ERR_MSG = "Parameter {0} has no value";

	final static private String NAMESPACE = "http://elexis.ch/Elexis_Artikelstamm_v5";
	final static private String XML_TAG_GTIN = "GTIN";
	final static private String XML_TAG_DESCR = "DSCR";
	final static private String XML_TAG_DEDUCTIBLE = "DEDUCTIBLE";
	final static private String XML_TAG_ITEMS = "ITEMS";

	private static final Logger logger = LogManager.getLogger(Accumulator.class);

	private static HashMap<String, String> values;

	public static void main(String[] args) {
		try {
			if (args.length > 13 && args.length <= 15) {
				logger.info(args.length);
				String asPath = getParamValue(args, ASTAMM_PATH_PARAM);
				String zipPath = getParamValue(args, PUB_ZIP_PARAM);
				String plName = getParamValue(args, PLISTE_NAME_PARAM);
				int worksheetNumber = Integer.parseInt(getParamValue(args, WORKSHEET_NUMBER_PARAM));
				int gtinColNumber = Integer.parseInt(getParamValue(args, GTIN_COL_NUMBER_PARAM));
				int percentageColNumber = Integer.parseInt(getParamValue(args, PERCENTAGE_COL_NUMBER_PARAM));
				int percentage = Integer.parseInt(getParamValue(args, PERCENTAGE_PARAM));
				HashMap<String, String> result = readXLSX(zipPath, plName, worksheetNumber, gtinColNumber,
						percentageColNumber, percentage, args.length == 15 ? true : false);
				modifyXML(asPath, result, percentage, args.length == 15 ? true : false);
			} else {
				logger.info(PHAR_ERR_MSG);
				System.exit(2);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static String getParamValue(String[] args, String param) {
		int pos = 0;
		String extractedParamValue = new String();
		for (String s : args) {
			if (param.equalsIgnoreCase(s)) {
				if (args.length < pos + 1) {
					logger.info(MessageFormat.format(EMPTY_PARAM_ERR_MSG, args[pos]));
				} else {
					extractedParamValue = (args[pos + 1]);
				}
			}
			pos++;
		}
		return extractedParamValue;
	}

	static HashMap<String, String> readXLSX(String path, String filename, int worksheetNumber, int gtinColNumber,
			int percentageColNumber, int percentage, boolean verbose) throws IOException {
		values = new HashMap<String, String>();

		List<File> files = new ArrayList<>();
		try {
			ZipInputStream zin = new ZipInputStream(new FileInputStream(new File(path)));
			ZipEntry entry = null;
			while ((entry = zin.getNextEntry()) != null) {
				if (verbose) {
					logger.info(entry.getName() + " (" + entry.getSize() + " bytes) detected.");
				}
				File file = new File(entry.getName());
				FileOutputStream os = new FileOutputStream(file);
				if (entry.getName().equalsIgnoreCase(filename)) {
					for (int c = zin.read(); c != -1; c = zin.read()) {
						os.write(c);
					}
				}
				os.close();
				files.add(file);
			}
			zin.close();
			for (File f : files) {
				if (f.getName().equalsIgnoreCase(filename)) {
					Workbook workbook = new XSSFWorkbook(f);
					Sheet sheet = workbook.getSheetAt(worksheetNumber);
					for (Row row : sheet) {
						Cell gtin = row.getCell(gtinColNumber);
						Cell percentYesNo = row.getCell(percentageColNumber);
						values.put(gtin.getStringCellValue(), percentYesNo.getStringCellValue());
					}
					workbook.close();
				}
			}
			if (verbose) {
				logger.info("Identified substances with a deductible of " + percentage + "%");
				int i = 1;
				for (String key : values.keySet()) {
					logger.info(i + ". " + key);
					i++;
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return values;
	}

	static void modifyXML(String path, HashMap<String, String> values, int percentage, boolean verbose)
			throws IOException, JDOMException {
		ArrayList<String> modified = new ArrayList<String>();
		ArrayList<String> skipped = new ArrayList<String>();

		Namespace ns = Namespace.getNamespace(NAMESPACE);

		SAXBuilder sax = new SAXBuilder();
		sax.setProperty(XMLConstants.ACCESS_EXTERNAL_DTD, "");
		sax.setProperty(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");

		Document doc = sax.build(new File(path));

		Element rootNode = doc.getRootElement();
		Element itemsElement = rootNode.getChild(XML_TAG_ITEMS, ns);

		List<Element> items = itemsElement.getChildren();

		if (items.size() > 0) {

			for (Element itemElement : items) {
				Element gtinElement = itemElement.getChild(XML_TAG_GTIN, ns);
				Element deductibleElement = itemElement.getChild(XML_TAG_DEDUCTIBLE, ns);
				Element descriptionElement = itemElement.getChild(XML_TAG_DESCR, ns);
				if (deductibleElement != null && values.containsKey(gtinElement.getText())) {
					logger.info(gtinElement.getText() + " (" + descriptionElement.getText() + ") "
							+ deductibleElement.getText() + " --> " + percentage + "%");
					deductibleElement.setText(Integer.valueOf(percentage).toString());
					modified.add(gtinElement.getText());
				}
			}
		}
		
		logger.info(values.size()+" substances modified");
		
		FileWriter writer = new FileWriter(path);
		XMLOutputter outputter = new XMLOutputter();

		Format format = Format.getRawFormat();
		format.setEncoding("UTF-8");

		outputter.setFormat(format);
		outputter.output(doc, writer);
		
		String xmlDocument = outputter.outputString(doc);
		if (verbose) {
			logger.info(xmlDocument);
		}
	}
}
