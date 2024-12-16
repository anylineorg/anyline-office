package org.anyline.office.tag;

public class Chart extends AbstractTag implements Tag {
    private String type;
    public void release(){
        super.release();
        type = null;
    }
}
