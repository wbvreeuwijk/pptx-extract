package net.reeuwijk.pptxextract;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

import com.google.cloud.language.v1.AnalyzeEntitiesRequest;
import com.google.cloud.language.v1.AnalyzeEntitiesResponse;
import com.google.cloud.language.v1.Document;
import com.google.cloud.language.v1.Document.Type;
import com.google.cloud.language.v1.EncodingType;
import com.google.cloud.language.v1.Entity;
import com.google.cloud.language.v1.LanguageServiceClient;

public class Slide {
	// Scale for image rendering
	public static final int SCALE = 2;

	private long id = 0l;
	private String title = null;
	private String text = null;
	private String notes = null;

	private List<Entity> entities;
	private String imgText = null;

	public Slide(XSLFSlide slide) {
		this.title = slide.getTitle();
		// Add notes
		XSLFNotes notes = slide.getNotes();
		if (notes != null) {
			this.notes = getParagraphs(notes.getTextParagraphs());
		}

		// Add slide text
		StringBuffer slideText = new StringBuffer();
		List<XSLFShape> shapes = slide.getShapes();
		for (XSLFShape xslfShape : shapes) {
			if (xslfShape instanceof XSLFTextShape) {
				XSLFTextShape text = (XSLFTextShape) xslfShape;
				if (this.title == null) {
					this.title = text.getText();
				}
				slideText.append(getParagraph(text.getTextParagraphs()));
			}
		}
		this.text = slideText.toString();

	}

	public String getId() {
		return Long.toString(id);
	}

	public String getTitle() {
		return title;
	}

	public String getText() {
		if (text == null) {
			return "Slide " + id;
		} else {
			return text;
		}
	}

	public String getNotes() {
		return notes;
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

	public void performKeywordAnalysis(LanguageServiceClient languageClient) {
		Document googleDoc = Document.newBuilder().setContent(this.text).setType(Type.PLAIN_TEXT).build();
		AnalyzeEntitiesRequest request = AnalyzeEntitiesRequest.newBuilder().setDocument(googleDoc)
				.setEncodingType(EncodingType.UTF16).build();

		AnalyzeEntitiesResponse response = languageClient.analyzeEntities(request);
		this.entities = response.getEntitiesList();
	}

	public void renderImage(XSLFSlide slide, int pgWidth, int pgHeight) {
		// Render the slide
		int width = (int) (pgWidth * Slide.SCALE);
		int height = (int) (pgHeight * Slide.SCALE);

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
		graphics.fill(new Rectangle2D.Float(0, 0, pgWidth, pgHeight));
		slide.draw(graphics);
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		try {
			javax.imageio.ImageIO.write(img, "png", out);
			imgText = Base64.encodeBase64String(out.toByteArray());
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public List<Entity> getEntities() {
		return entities;
	}
	
	public String getImage() {
		return this.imgText;
	}

	public void setId(long id2) {
		this.id = id2;
	}
}
