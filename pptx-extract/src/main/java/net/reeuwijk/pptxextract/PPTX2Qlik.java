package net.reeuwijk.pptxextract;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

import com.aylien.textapi.TextAPIClient;
import com.aylien.textapi.TextAPIException;
import com.aylien.textapi.parameters.EntitiesParams;
import com.aylien.textapi.parameters.EntitiesParams.Builder;
import com.aylien.textapi.responses.Entities;
import com.aylien.textapi.responses.Entity;

public class PPTX2Qlik {

	// Scale for image rendering
	private static final int SCALE = 1;

	private static final String AYLIEN_APP_ID = "b08da0e5";
	private static final String AYLIEN_APP_KEY = "9100f1e05844c1fec08df9874331a93a";

	private static Document doc;

	private static TextAPIClient client;

	public static void main(String[] args) throws Exception {
		// We need at least the name of the presentation
		if (args.length == 0) {
			usage();
			return;
		}

        // Process file or directory
		ArrayList<File> toProcess = new ArrayList<File>();
		File pptxFile = new File(args[0]);
		if(pptxFile.isFile()) {
			toProcess.add(pptxFile);
		} else if(pptxFile.isDirectory()) {
			System.out.println("Looking for powerpoint presentations in directory " + pptxFile);
			File[] directory = pptxFile.listFiles();
			for (File file : directory) {
				if(file.getName().endsWith(".pptx")) {
					toProcess.add(file);
				}
			}
		}
		
		for (File file : toProcess) {
			// Open the file
			System.out.println("Processing file " + file);
			File outFile = new File(file.getParent(), file.getName().replaceAll("\\.pptx", ".xml"));

			// Read the .pptx file
			XMLSlideShow ppt = new XMLSlideShow(OPCPackage.open(file));

			// Setup XML document Builder
			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
			doc = docBuilder.newDocument();
			Element rootElement = doc.createElement("powerpoint");
			rootElement.setAttribute("name", pptxFile.getName().replaceAll("\\.pptx", ""));
			doc.appendChild(rootElement);

			// Setup imagebuffer
			Dimension pgsize = ppt.getPageSize();
			int width = (int) (pgsize.width * SCALE);
			int height = (int) (pgsize.height * SCALE);

			// Setup Text analysis API
			//client = new TextAPIClient(AYLIEN_APP_ID, AYLIEN_APP_KEY);

			List<XSLFSlide> slides = ppt.getSlides();
			for (XSLFSlide slide : slides) {
				// Create slide element
				Element slideElement = doc.createElement("slide");
				String title = slide.getTitle();
				System.out.print(".");
				if (title != null)
					slideElement.setAttribute("title", title);
				slideElement.setAttribute("ID", Integer.toString(slide.getSlideNumber()));

				//Thread.sleep(1500);
				// Add notes
				StringBuffer slideText = new StringBuffer();
				XSLFNotes notes = slide.getNotes();
				if (notes != null) {
					Element notesElement = doc.createElement("notes");
					String paragraph = getParagraphs(notes.getTextParagraphs());
					Text textNode = doc.createTextNode(paragraph);
					notesElement.appendChild(textNode);
					slideElement.appendChild(notesElement);
					slideText.append(paragraph);
				}

				//Thread.sleep(1000);
				// Add slide text
				List<XSLFShape> shapes = slide.getShapes();
				for (XSLFShape xslfShape : shapes) {
					if (xslfShape instanceof XSLFTextShape) {
						XSLFTextShape text = (XSLFTextShape) xslfShape;
						Element textShapeElement = doc.createElement("text");
						String paragraph = getParagraph(text.getTextParagraphs());
						Text textNode = doc.createTextNode(paragraph);
						textShapeElement.appendChild(textNode);
						slideElement.appendChild(textShapeElement);
						slideText.append(paragraph);
						// } else {
						// System.out.println("Process me: " + xslfShape.getClass());
					}
				}

				//slideElement.appendChild(getAnalysisElement(slideText.toString()));
				rootElement.appendChild(slideElement);

				// Render the slide
				BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
				Graphics2D graphics = img.createGraphics();
				DrawFactory.getInstance(graphics).fixFonts(graphics);

				// default rendering options
				graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
				graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
				graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
				graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);

				graphics.scale(SCALE, SCALE);

				graphics.setPaint(Color.white);
				graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
				slide.draw(graphics);
				ByteArrayOutputStream out = new ByteArrayOutputStream();
				javax.imageio.ImageIO.write(img, "png", out);
				Element imageElement = doc.createElement("image");
				Text imageNode = doc.createTextNode(Base64.encodeBase64String(out.toByteArray()));
				imageElement.appendChild(imageNode);
				slideElement.appendChild(imageElement);
			}
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			DOMSource source = new DOMSource(doc);
			StreamResult result = new StreamResult(outFile);

			transformer.transform(source, result);
			System.out.println(".");
			System.out.println("Done");
			ppt.close();
			
		}
	}

	private static Element getAnalysisElement(String text) throws TextAPIException {
		Element analysisElement = doc.createElement("analysis");
		Builder builder = EntitiesParams.newBuilder();
		builder.setText(text);
		builder.setLanguage("en");
		Entities entities = client.entities(builder.build());
		for (Entity entity : entities.getEntities()) {
			String entityType = entity.getType();
			for (String sf : entity.getSurfaceForms()) {
				Element entityElement = doc.createElement(entityType);
				Text entTextNode = doc.createTextNode(sf);
				entityElement.appendChild(entTextNode);
				analysisElement.appendChild(entityElement);
			}
		}
		return analysisElement;
	}

	private static String getParagraph(List<XSLFTextParagraph> textParagraphs) {
		StringBuffer str = new StringBuffer();

		for (XSLFTextParagraph xslfTextParagraph : textParagraphs) {
			String text = xslfTextParagraph.getText();
			if (!text.isEmpty()) {
				str.append(text.replaceAll("\\s+", " "));
				str.append(" ");
			}
		}
		return str.toString();
	}

	private static String getParagraphs(List<List<XSLFTextParagraph>> listparagraphs) {
		StringBuffer str = new StringBuffer();
		for (List<XSLFTextParagraph> textParagraphs : listparagraphs) {
			str.append(getParagraph(textParagraphs));
		}
		return str.toString();

	}

	private static void usage() {
		System.err.println("Usgae: PPTX2Qlik <pptx file|directory>");
	}

}
