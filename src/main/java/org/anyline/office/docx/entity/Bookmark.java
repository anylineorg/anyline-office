package org.anyline.office.docx.entity;

public class Bookmark {
    private String name;
    public Bookmark(){}
    public Bookmark(String name){
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
