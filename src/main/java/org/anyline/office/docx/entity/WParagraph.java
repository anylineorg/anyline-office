/*
 * Copyright 2006-2023 www.anyline.org
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */




package org.anyline.office.docx.entity;

import org.anyline.handler.Uploader;
import org.anyline.office.docx.util.DocxUtil;
import org.anyline.office.docx.util.StyleUtil;
import org.anyline.util.BasicUtil;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;

public class WParagraph extends WElement {
    private List<WRun> wrs = new ArrayList<>();
    public WParagraph(WDocument doc, Element src){
        this.root = doc;
        this.src = src;
        load();
    }
    public void reload(){
        load();
    }
    private void load(){
        wrs.clear();
        List<Element> elements = src.elements("r");
        for(Element element:elements){
            WRun wr = new WRun(root, element);
            wrs.add(wr);
        }
    }
    public WParagraph setColor(String color){
        for(WRun wr:wrs){
            wr.setColor(color);
        }
        Element pr = DocxUtil.element(src, "pPr");
        DocxUtil.element(pr, "color","val", color.replace("#",""));
        return this;
    }
    public WParagraph setFont(String size, String eastAsia, String ascii, String hint){

        for(WRun wr:wrs){
            wr.setFont(size, eastAsia, ascii, hint);
        }
        int pt = StyleUtil.fontSize(size);
        Element pr = DocxUtil.element(src, "pPr");
        DocxUtil.element(pr, "sz","val", pt+"");
        DocxUtil.element(pr, "rFonts","eastAsia", eastAsia);
        DocxUtil.element(pr, "rFonts","ascii", ascii);
        DocxUtil.element(pr, "rFonts","hint", hint);

        return this;
    }
    public WParagraph setFontSize(String size){
        for(WRun wr:wrs){
            wr.setFontSize(size);
        }
        int pt = StyleUtil.fontSize(size);
        Element pr = DocxUtil.element(src, "pPr");
        DocxUtil.element(pr, "sz","val", pt+"");
        return this;
    }
    public WParagraph setFontFamily(String font){
        for(WRun wr:wrs){
            wr.setFontFamily(font);
        }
        Element pr = DocxUtil.element(src, "pPr");
        DocxUtil.element(pr, "rFonts","eastAsia", font);
        DocxUtil.element(pr, "rFonts","ascii", font);
        DocxUtil.element(pr, "rFonts","hAnsi", font);
        DocxUtil.element(pr, "rFonts","cs", font);
        DocxUtil.element(pr, "rFonts","hint", font);
        return this;
    }

    public WParagraph setAlign(String align){
        Element pr = DocxUtil.element(src, "pPr");
        DocxUtil.element(pr, "jc","val", align);
        return this;
    }

    public WParagraph setBackgroundColor(String color){
        Element pr = DocxUtil.element(src, "pPr");
        color = color.replace("#","");
        DocxUtil.element(pr, "highlight", "val", color);
        for(WRun wr:wrs){
            wr.setBackgroundColor(color);
        }
        return this;
    }

    public WParagraph setBold(boolean bold){
        Element pr = DocxUtil.element(src, "pPr");
        Element b = pr.element("b");
        if(bold){
            if(null == b){
                pr.addElement("w:b");
            }
        }else{
            if(null != b){
                pr.remove(b);
            }
        }
        for(WRun wr:wrs){
            wr.setBold(bold);
        }
        return this;
    }
    public WParagraph setUnderline(boolean underline){
        Element pr = DocxUtil.element(src, "pPr");
        Element u = pr.element("u");
        if(underline){
            if(null == u){
                DocxUtil.element(pr, "u", "val", "single");
            }
        }else{
            if(null != u){
                pr.remove(u);
            }
        }
        for(WRun wr:wrs){
            wr.setUnderline(underline);
        }
        return this;
    }
    public WParagraph setStrike(boolean strike){
        Element pr = DocxUtil.element(src, "pPr");
        Element s = pr.element("strike");
        if(strike){
            if(null == s){
                pr.addElement("w:strike");
            }
        }else{
            if(null != s){
                pr.remove(s);
            }
        }
        for(WRun wr:wrs){
            wr.setStrike(strike);
        }
        return this;
    }
    public WParagraph setItalic(boolean italic){
        Element pr = DocxUtil.element(src, "pPr");
        DocxUtil.element(pr, "i","val",italic+"");
        for(WRun wr:wrs){
            wr.setItalic(italic);
        }
        return this;
    }

    /**
     * 清除样式
     * @return wp
     */
    public WParagraph removeStyle(){
        Element pr = src.element("pPr");
        if(null != pr){
            src.remove(pr);
        }
        for(WRun wr:wrs){
            wr.removeStyle();
        }
        return this;
    }
    /**
     * 清除背景色
     * @return wp
     */
    public WParagraph removeBackgroundColor(){
        DocxUtil.removeElement(src,"shd");
        return this;
    }
    /**
     * 清除颜色
     * @return wp
     */
    public WParagraph removeColor(){
        DocxUtil.removeElement(src,"color");
        return this;
    }
    public WParagraph addWr(WRun wr){
        wrs.add(wr);
        return this;
    }
    public WParagraph replace(String target, String replacement){
        for(WRun wr:wrs){
            wr.replace(target, replacement);
        }
        return this;
    }

    public LinkedHashMap<String, String> styles(){
        LinkedHashMap<String, String> styles = new LinkedHashMap<>();
        Element pr = src.element("pPr");
        if(null != pr){
           Element rpr = pr.element("rPr");
           if(null != rpr){
               //<w:color w:val="0070C0"/>
               Element color = rpr.element("color");
               if(null != color){
                   String val = color(color.attributeValue("val"));
                   if(null != val){
                       styles.put("color", val);
                   }
               }
               //<w:highlight w:val="yellow"/>
               Element highlight = pr.element("highlight");
               if(null != highlight){
                   String val = color(highlight.attributeValue("val"));
                   if(null != val){
                       styles.put("background-color", val);
                   }
               }
               //<w:rFonts w:hint="eastAsia"/>
               Element font = pr.element("rFonts");
               if(null != font){
                   String hint = font.attributeValue("hint");
                   if(null != hint){
                       styles.put("font-family", hint);
                   }
               }
               //<w:sz w:val="24"/>
               Element size = pr.element("sz");
               if(null == size){
                   size = pr.element("szCs");
               }
               if(null != size){
                   int val = BasicUtil.parseInt(size.attributeValue("val"), 0);
                   if(val > 0){
                       styles.put("font-size", val+"pt");
                   }
               }
               //<w:jc w:val="center"/>
               Element jc = pr.element("jc");
               if(null != jc){
                   String val = jc.attributeValue("val");
                   if(null != val){
                       styles.put("text-align", val);
                   }
               }
           }
        }
        return styles;
    }

    public String html(Uploader uploader, int lvl){
        StringBuilder builder = new StringBuilder();
        StringBuilder body = new StringBuilder();
        Iterator<Element> items = src.elementIterator();
        while (items.hasNext()){
            Element item = items.next();
            String tag = item.getName();
           if(tag.equalsIgnoreCase("r")){
                body.append("\n");
                body.append(new WRun(getDoc(), item).html(uploader, lvl+1));
            } else if(tag.equalsIgnoreCase("tbl")){
                body.append("\n");
                body.append(new WTable(getDoc(), item).html(uploader, lvl+1));
            }
        }
        t(builder, lvl);
        builder.append("<div");
        styles(builder);
        builder.append(">");
        builder.append(body);
        builder.append("\n");
        t(builder, lvl);
        builder.append("</div>");
        return builder.toString();
    }
}
