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
import org.anyline.util.BasicUtil;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;

public class WTr extends WElement {
    private WTable parent;
    private List<WTc> wtcs = new ArrayList<>();
    private String widthUnit = "px";     // 默认长度单位 px pt cm/厘米
    public WTr(WDocument doc, WTable parent, Element src){
        this.root = doc;
        this.src = src;
        this.parent = parent;
        load();
    }

    public void reload(){
        load();
    }
    private WTr load(){
        wtcs.clear();
        List<Element> items = src.elements("tc");
        for(Element tc:items){
            WTc wtc = new WTc(root, this, tc);
            wtcs.add(wtc);
        }
        return this;
    }

    public WTable getParent(){
        return parent;
    }

    public WTr setHeight(String height){
        int dxa = DocxUtil.dxa(height);
        Element pr = DocxUtil.addElement(src, "trPr");
        DocxUtil.addElement(pr,"trHeight", "val", dxa+"" );
        return this;
    }
    public WTr setHeight(int height){
        return setHeight(height+widthUnit);
    }
    public WTr setHeight(double height){
        return setHeight(height+widthUnit);
    }
    public List<WTc> getWtcs(){
        if(wtcs.isEmpty()){
            List<Element> elements = src.elements("tc");
            for(Element element:elements){
                WTc tc = new WTc(root,this, element);
                wtcs.add(tc);
            }
        }
        return wtcs;
    }
    public WTc getTc(int index){
        return wtcs.get(index);
    }

    public String getWidthUnit() {
        return widthUnit;
    }

    public void setWidthUnit(String widthUnit) {
        this.widthUnit = widthUnit;
        for(WTc tc:wtcs){
            tc.setWidthUnit(widthUnit);
        }
    }

    /**
     * 根据书签或占位符获取列
     * @param bookmark 书签或占位符 包含{和}的按占位符搜索
     * @return wtr
     */
    public WTc tc(String bookmark){
        List<WTc> tcs = tcs(bookmark);
        if(!tcs.isEmpty()){
            return tcs.get(0);
        }
        return null;
    }
    public List<WTc> tcs(String bookmark){
        List<WTc> list = new ArrayList<>();
        if(null != bookmark) {
            for(WTc item:wtcs){
                String txt = item.getTexts();
                if(txt.contains(bookmark)){
                    list.add(item);
                }
            }
        }
        return list;
    }
    /**
     * 获取单元格,计算合并列
     * @param index 索引
     * @param prev 如果index位置被合并了,是否返 当前合并组中的第一个单元格
     * @return tc
     */
    public WTc getTcWithColspan(int index, boolean prev){
        int qty = -1;
        for(WTc tc:wtcs){
            qty += tc.getColspan();
            if(qty == index){
                return tc;
            }

            if(qty > index){
                if(prev){
                    return tc;
                }else {
                    break;
                }
            }
        }
        return null;
    }
    public List<WTc> getTcs(){
        return wtcs;
    }


    private WTr removeBorder(){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.removeBorder();
        }
        return this;
    }
    public WTr setBorder(){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setBorder();
        }
        return this;
    }

    /**
     * 设置边框
     * @param size 宽度根据width unit单位
     * @param color 颜色
     * @param style 样式
     * @return tr
     */
    public WTr setBorder(int size, String color, String style){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setBorder(size, color, style);
        }
        return this;
    }
    public WTr setTopBorder(int size, String color, String style){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setTopBorder(size, color, style);
        }
        return this;
    }
    public WTr removeTopBorder(){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.removeTopBorder();
        }
        return this;
    }
    public WTr setBottomBorder(int size, String color, String style){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setBottomBorder(size, color, style);
        }
        return this;
    }
    public WTr removeBottomBorder(){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.removeBottomBorder();
        }
        return this;
    }
    /**
     * 设置颜色
     * @param color color
     * @return tr
     */
    public WTr setColor(String color){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setColor(color);
        }
        return this;
    }

    /**
     * 设置字体
     * @param size 字号
     * @param eastAsia 中文字体
     * @param ascii 英文字体
     * @param hint 默认字体
     * @return tr
     */
    public WTr setFont(String size, String eastAsia, String ascii, String hint){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setFont(size, eastAsia, ascii, hint);
        }
        return this;
    }

    /**
     * 设置字号
     * @param size px|pt|cm
     * @return tr
     */
    public WTr setFontSize(String size){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setFontSize(size);
        }
        return this;
    }

    /**
     * 设置字体
     * @param font 字体
     * @return tr
     */
    public WTr setFontFamily(String font){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setFontFamily(font);
        }
        return this;
    }

    /**
     * 设置水平对齐方式
     * @param align start/left center end/right
     * @return tr
     */
    public WTr setAlign(String align){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setAlign(align);
        }
        return this;
    }

    /**
     * 设置垂直对齐方式
     * @param align top/center/bottom
     * @return Wtr
     */
    public WTr setVerticalAlign(String align){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setVerticalAlign(align);
        }
        return this;
    }

    /**
     * 设置整行背景色
     * @param color color
     * @return  Wtr
     */
    public WTr setBackgroundColor(String color){
        List<WTc> tcs = getWtcs();
        for(WTc tc:tcs){
            tc.setBackgroundColor(color);
        }
        return this;
    }

    public WTr removeStyle(){
        for(WTc tc:wtcs){
            tc.removeStyle();
        }
        return this;
    }
    public WTr removeBackgroundColor(){
        for(WTc tc:wtcs){
            tc.removeBackgroundColor();
        }
        return this;
    }
    public WTr removeColor(){
        for(WTc tc:wtcs){
            tc.removeColor();
        }
        return this;
    }
    public WTr replace(String src, String tar){
        for(WTc tc:wtcs){
            tc.replace(src, tar);
        }
        return this;
    }
    public WTr setBold(){
        for(WTc tc:wtcs){
            tc.setBold();
        }
        return this;
    }
    public WTr setBold(boolean bold){
        for(WTc tc:wtcs){
            tc.setBold(bold);
        }
        return this;
    }
    /**
     * 下划线
     * @param underline 是否
     * @return Wtc
     */
    public WTr setUnderline(boolean underline){
        for(WTc tc:wtcs){
            tc.setUnderline(underline);
        }
        return this;
    }
    public WTr setUnderline(){
        setUnderline(true);
        return this;
    }

    /**
     * 删除线
     * @param strike 是否
     * @return Wtc
     */
    public WTr setStrike(boolean strike){
        for(WTc tc:wtcs){
            tc.setStrike(strike);
        }
        return this;
    }
    public WTr setStrike(){
        setStrike(true);
        return this;
    }

    /**
     * 斜体
     * @param italic 是否
     * @return Wtc
     */
    public WTr setItalic(boolean italic){
        for(WTc tc:wtcs){
            tc.setItalic(italic);
        }
        return this;
    }

    public WTr setItalic(){
        return setItalic(true);
    }
    public WTr setPadding(String side, double padding){
        for(WTc tc:wtcs){
            tc.setPadding(side, padding);
        }
        return this;
    }
    public WTr setPadding(String side, String padding){
        for(WTc tc:wtcs){
            tc.setPadding(side, padding);
        }
        return this;
    }
    public WTr setPadding(String side, int padding){
        for(WTc tc:wtcs){
            tc.setPadding(side, padding);
        }
        return this;
    }

    public WTr setLeftPadding(double padding){
        for(WTc tc:wtcs){
            tc.setLeftPadding(padding);
        }
        return this;
    }
    public WTr setLeftPadding(String padding){
        for(WTc tc:wtcs){
            tc.setLeftPadding(padding);
        }
        return this;
    }
    public WTr setLeftPadding(int padding){
        for(WTc tc:wtcs){
            tc.setLeftPadding(padding);
        }
        return this;
    }

    public WTr setRightPadding(double padding){
        for(WTc tc:wtcs){
            tc.setRightPadding(padding);
        }
        return this;
    }
    public WTr setRightPadding(String padding){
        for(WTc tc:wtcs){
            tc.setRightPadding(padding);
        }
        return this;
    }
    public WTr setRightPadding(int padding){
        for(WTc tc:wtcs){
            tc.setRightPadding(padding);
        }
        return this;
    }

    public WTr setTopPadding(double padding){
        for(WTc tc:wtcs){
            tc.setTopPadding(padding);
        }
        return this;
    }
    public WTr setTopPadding(String padding){
        for(WTc tc:wtcs){
            tc.setTopPadding(padding);
        }
        return this;
    }
    public WTr setTopPadding(int padding){
        for(WTc tc:wtcs){
            tc.setTopPadding(padding);
        }
        return this;
    }


    public WTr setBottomPadding(double padding){
        for(WTc tc:wtcs){
            tc.setBottomPadding(padding);
        }
        return this;
    }
    public WTr setBottomPadding(String padding){
        for(WTc tc:wtcs){
            tc.setBottomPadding(padding);
        }
        return this;
    }
    public WTr setBottomPadding(int padding){
        for(WTc tc:wtcs){
            tc.setBottomPadding(padding);
        }
        return this;
    }


    public WTr setPadding(double padding){
        for(WTc tc:wtcs){
            tc.setPadding(padding);
        }
        return this;
    }
    public WTr setPadding(String padding){
        for(WTc tc:wtcs){
            tc.setPadding(padding);
        }
        return this;
    }
    public WTr setPadding(int padding){
        for(WTc tc:wtcs){
            tc.setPadding(padding);
        }
        return this;
    }


    /**
     * 复制一行
     * @param content 是否复制其中内容
     * @return wtr
     */
    public WTr clone(boolean content){
        WTr tr = new WTr(root, this.getParent(), this.getSrc().createCopy());
        if(!content){
            tr.removeContent();
        }
        return tr;
    }


    public LinkedHashMap<String, String> styles(){
        LinkedHashMap<String, String> styles = new LinkedHashMap<>();
        Element pr = src.element("trPr");
        if(null != pr){
            //<w:trHeight w:val="284"/>
            Element h = pr.element("trHeight");
            if(null != h){
                int height = BasicUtil.parseInt(h.attributeValue("val"), 0);
                if(height > 0){
                    styles.put("height", (int)DocxUtil.dxa2px(height)+"px");
                }
            }
        }
        return styles;
    }
    public String html(Uploader uploader, int lvl){
        StringBuilder builder = new StringBuilder();
        LinkedHashMap<String, String> styles = new LinkedHashMap<>();
        StringBuilder body = new StringBuilder();
        Iterator<Element> items = src.elementIterator();
        while (items.hasNext()){
            Element item = items.next();
            String tag = item.getName();if(tag.equalsIgnoreCase("tc")){
                body.append("\n");
                body.append(new WTc(getDoc(), this, item).html(uploader, lvl+1));
            }
        }
        t(builder, lvl);
        builder.append("<tr");
        styles(builder);
        builder.append(">");
        builder.append(body);
        builder.append("\n");
        t(builder, lvl);
        builder.append("</tr>");
        return builder.toString();
    }
}
