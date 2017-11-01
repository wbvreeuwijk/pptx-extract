package net.reeuwijk.pptxextract;

import java.util.ArrayList;

public class Section {
	private String name = null;
	private ArrayList<String> slideIds = new ArrayList<String>();
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public ArrayList<String> getSlideIds() {
		return slideIds;
	}
	public void addSlideId(String slide) {
		this.slideIds.add(slide);
	}
}
