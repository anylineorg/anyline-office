package org.anyline.office.tag;

import javax.xml.bind.Element;
import java.util.List;

public class TagBox {
    private TagElement head;
    private TagElement foot;

    private List<Element> tops;

    public TagElement head() {
        return head;
    }

    public void head(TagElement head) {
        this.head = head;
    }

    public TagElement foot() {
        return foot;
    }

    public void foot(TagElement foot) {
        this.foot = foot;
    }

    public List<Element> tops() {
        return tops;
    }

    public void tops(List<Element> tops) {
        this.tops = tops;
    }
}
