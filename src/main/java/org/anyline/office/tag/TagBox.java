package org.anyline.office.tag;

import javax.xml.bind.Element;
import java.util.ArrayList;
import java.util.List;

public class TagBox {
    private String text;
    private List<Element> ts = new ArrayList<>();
    public String text(){
        return text;
    }
    public void text(String text){
        this.text = text;
    }
    public void ts(List<Element> ts){
        this.ts = ts;
    }
    public List<Element> ts(){
        return ts;
    }

}
