package org.anyline.office.tag;

import javax.xml.bind.Element;

public class TagElement {
    private String text;
    private Element element;

    public String text() {
        return text;
    }

    public void text(String text) {
        this.text = text;
    }

    public Element element() {
        return element;
    }

    public void element(Element element) {
        this.element = element;
    }
}
