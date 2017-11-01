package net.reeuwijk.pptxextract;

import java.awt.Dimension;
import java.io.File;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.openxmlformats.schemas.presentationml.x2006.main.CTExtension;
import org.openxmlformats.schemas.presentationml.x2006.main.CTExtensionList;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPresentation;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlideIdList;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlideIdListEntry;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.google.cloud.language.v1.Entity;
import com.google.cloud.language.v1.LanguageServiceClient;

import de.tiq.solutions.data.conversion.qvx.QVXWriter;

public class PPTX2Qlik {
	
	private static LanguageServiceClient languageClient;

	public static void main(String[] args) throws Exception {
		// We need at least the name of the presentation
		if (args.length == 0) {
			usage();
			return;
		}

		// Getting Google API Client
		languageClient = LanguageServiceClient.create();

		// Process file or directory
		ArrayList<File> toProcess = new ArrayList<File>();
		File pptxFile = new File(args[0]);
		File presentationQVX = new File(pptxFile.getParentFile(),"Presentations.QVX");
		File keywordQVX = new File(pptxFile.getParentFile(),"Keywords.QVX");
		
		if (pptxFile.isFile()) {
			toProcess.add(pptxFile);
		} else if (pptxFile.isDirectory()) {
			System.out.println("Looking for powerpoint presentations in directory " + pptxFile);
			File[] directory = pptxFile.listFiles();
			for (File file : directory) {
				if (file.getName().endsWith(".pptx")) {
					toProcess.add(file);
				}
			}
		}

//		PrintWriter outFile = new PrintWriter(new File(args[1]));
//		outFile.println("Title1\tTitle2\tTitle3\tImage");
		QVXWriter presentationWriter = new QVXWriter(presentationQVX.getAbsolutePath(),"Presentations");
		String[] qvxHeader = new String[] {"Id","Title1","Title2","Title3","Image","Text","Notes"};
		presentationWriter.writeXMLHeader(qvxHeader);
		
		QVXWriter keywordWriter = new QVXWriter(keywordQVX.getAbsolutePath(), "Keywords");
		qvxHeader = new String[] {"Id","Name","Type","Salience"};
		keywordWriter.writeXMLHeader(qvxHeader);
		
		int slideIdCounter = 1;

		for (File file : toProcess) {
			// Open the file
			System.out.println("Processing file " + file);

			// Read the .pptx file
			XMLSlideShow ppt = new XMLSlideShow(OPCPackage.open(file));

			// Setup imagebuffer
			Dimension pgsize = ppt.getPageSize();

			ArrayList<Section> sectionArray = new ArrayList<Section>();
			// Get Section information from presentation
			CTPresentation ctPresentation = ppt.getCTPresentation();
			CTExtensionList extensionList = ctPresentation.getExtLst();
			CTSlideIdList sldIdList = ctPresentation.getSldIdLst();
			List<CTSlideIdListEntry> slideIdList = sldIdList.getSldIdList();
			List<CTExtension> list = extensionList.getExtList();
			for (CTExtension ctExtension : list) {
				Node extensionNode = ctExtension.getDomNode();				
				Node firstChild = extensionNode.getChildNodes().item(0);
				if("p14:sectionLst".equals(firstChild.getNodeName())) {
					System.out.println("Processing Sections");
					NodeList sections = firstChild.getChildNodes();
					for (int i = 0; i < sections.getLength(); i++) {
						NamedNodeMap attrs = sections.item(i).getAttributes();
						Node nameAttr = attrs.getNamedItem("name");
						System.out.println("["+i+"]="+nameAttr.getNodeValue());
						 Section s = new Section();
						s.setName(nameAttr.getNodeValue());
						NodeList slideIds = sections.item(i).getFirstChild().getChildNodes();
						for(int j =0; j < slideIds.getLength(); j++) {
							String slideId = slideIds.item(j).getAttributes().getNamedItem("id").getNodeValue();
							s.addSlideId(slideId);
						}
						sectionArray.add(s);
					}
				}
			}

			Hashtable<String, Slide> slideTable = new Hashtable<String,Slide>();
			// Loop through slides
			List<XSLFSlide> slides = ppt.getSlides();
			int slideCounter = 0;
			for (XSLFSlide slide : slides) {
				// Create slide element
				Slide slideObj = new Slide(slide);
				long id = slideIdList.get(slideCounter).getId();
				slideObj.setId(id);
				System.out.print(".");

				slideObj.performKeywordAnalysis(languageClient);
				slideObj.renderImage(slide,pgsize.width,pgsize.height);
				slideTable.put(slideObj.getId(),slideObj);
				slideCounter++;
			}

			for(int i = 0; i < sectionArray.size(); i++) {
				Section s = sectionArray.get(i);
				ArrayList<String> ids = s.getSlideIds();
				for (String slideId : ids) {
					Slide slide = slideTable.get(slideId);
					if(slide != null) {
						presentationWriter.writeData(Integer.toString(slideIdCounter));
						presentationWriter.writeData(file.getName().replaceFirst("\\.pptx$", ""));
						presentationWriter.writeData(s.getName());
						presentationWriter.writeData(slide.getTitle());
						presentationWriter.writeData(slide.getImage());
						presentationWriter.writeData(slide.getText());
						presentationWriter.writeData(slide.getNotes());
						List<Entity> entities = slide.getEntities();
						for (Entity entity : entities) {
							keywordWriter.writeData(Integer.toString(slideIdCounter));
							keywordWriter.writeData(entity.getName());
							keywordWriter.writeData(entity.getType().toString());
							keywordWriter.writeData(Float.toString(entity.getSalience()));
						}
					} else {
						System.out.println("Missing:"+slideId);
					}
					slideIdCounter++;
				}
			}
			
			System.out.println(".");
			System.out.println("Done");
			ppt.close();
		}
		if (presentationWriter != null) {
			try {
				presentationWriter.close();
			} catch (Exception e) {
			}
		}
		if (keywordWriter != null) {
			try {
				keywordWriter.close();
			} catch (Exception e) {
			}
		}
	}

	private static void usage() {
		System.err.println("Usgae: PPTX2Qlik <pptx file|directory> <outputfile>");
	}

}
