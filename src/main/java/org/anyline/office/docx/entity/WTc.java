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
import org.anyline.util.DomUtil;
import org.anyline.util.HtmlUtil;
import org.anyline.util.StyleParser;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;

import java.util.*;

public class WTc extends WElement {
    private WTr parent;
    private List<WParagraph> wps = new ArrayList<>();
    private String widthUnit = "px";     // 默认长度单位 px pt cm/厘米
    private int colspan = -1; //-1:未设置  1:不合并  0:被合并  >1:合并其他单元格
    private int rowspan = -1;
    private static Map<Element, WTc> map = new HashMap<>();

    public static WTc tc(Element src){
        if(null == src){
            return null;
        }
        return map.get(src);
    }
    public WTc(WDocument doc, WTr parent, Element src){
        this.root = doc;
        this.src = src;
        this.parent = parent;
        load();
    }
    public void reload(){
        load();
    }
    private WTc load(){
        map.put(src, this);
        wps.clear();
        List<Element> ps = src.elements("p");
        for(Element p:ps){
            WParagraph wp = new WParagraph(root, p);
            wps.add(wp);
        }
        return this;
    }

    /**
     * 当前单元格内所有书签名称
     * @return list
     */
    public List<String> getBookmarks(){
        List<String> list = new ArrayList<>();
        List<Element> marks = DomUtil.elements(src, "bookmarkStart");
        for(Element mark:marks){
            list.add(mark.attributeValue("name"));
        }
        return list;
    }

    /**
     * 当前单元格内第一个书签名称
     * @return String
     */
    public String getBookmark(){
        Element mark = DomUtil.element(src, "bookmarkStart");
        if(null != mark){
            return mark.attributeValue("name");
        }
        return null;
    }
    public WTc setBorder(String side, String style){
        return this;
    }

    /**
     * 宽度尺寸单位
     * @return String
     */
    public String getWidthUnit() {
        return widthUnit;
    }

    public void setWidthUnit(String widthUnit) {
        this.widthUnit = widthUnit;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }

    public void setRowspan(int rowspan) {
        this.rowspan = rowspan;
    }

    /**
     * 当前单元格合并列数量
     * @return colspan
     */
    public int getColspan(){
        if(colspan == -1) {
            colspan = parseColspan();
        }
        if(colspan == -1){
            return 1;
        }
        return colspan;
    }

    public int parseColspan(){
        Element tcPr = src.element("tcPr");
        if (null != tcPr) {
            Element gridSpan = tcPr.element("gridSpan");
            if (null != gridSpan) {
                colspan = BasicUtil.parseInt(gridSpan.attributeValue("val"), 1);
            }
        }
        return colspan;
    }

    /**
     * 当前单元格合并行数量,被合并返回-1
     * @return rowspan
     */
    public int getRowspan(){
        if(rowspan == -1){
            return 1;
        }
        return rowspan;
    }

    /**
     *
     * @return 2:合并其他行 1:不合并 0:被合并
     */
    public int parseRowspan(){
        Element tcPr = src.element("tcPr");
        if(null != tcPr){
            Element vMerge = tcPr.element("vMerge");
            if(null != vMerge){
                String val = vMerge.attributeValue("val");
                if(!"restart".equalsIgnoreCase(val)){
                    return 0;
                }
                return 2;
            }
        }
        return 1;
    }

    /**
     * 当前单元格 左侧单元格
     * @return wtc
     */
    public WTc left(){
        WTc left = null;
        List<WTc> tcs = parent.getTcs();
        int index = tcs.indexOf(this);
        if(index > 0){
            left = tcs.get(index-1);
        }
        return left;
    }
    /**
     * 当前单元格 右侧单元格
     * @return wtc
     */
    public WTc right(){
        WTc right = null;
        List<WTc> tcs = parent.getTcs();
        int index = tcs.indexOf(this);
        if(index < tcs.size()-1){
            right = tcs.get(index+1);
        }
        return right;
    }

    /**
     * 当前单元格 下方单元格
     * @return wtc
     */
    public WTc bottom(){
        WTc bottom = null;
        WTable table = parent.getParent();
        List<WTr> trs = table.getTrs();
        int y = trs.indexOf(parent);
        if(y < trs.size()-1){
            WTr tr = trs.get(y+1);
            int x = parent.getTcs().indexOf(this);
            bottom = tr.getTc(x);
        }
        return bottom;
    }
    /**
     * 当前单元格 上方单元格
     * @return wtc
     */
    public WTc top(){
        WTc top = null;
        WTable table = parent.getParent();
        List<WTr> trs = table.getTrs();
        int y = trs.indexOf(parent);
        if(y < trs.size()-1 && y>0){
            WTr tr = trs.get(y-1);
            int x = parent.getTcs().indexOf(this);
            top = tr.getTc(x);
        }
        return top;
    }
    /**
     * 删除左边框
     * @return wtc
     */
    public WTc removeLeftBorder(){
        removeBorder(src, "left");
        WTc left = left();
        if(null != left) {
            removeBorder(left.getSrc(), "right");
        }
        return this;
    }
    /**
     * 删除右边框
     * @return wtc
     */
    public WTc removeRightBorder(){
        removeBorder(src, "right");
        WTc right = right();
        if(null != right) {
            removeBorder(right.getSrc(), "left");
        }
        return this;
    }
    /**
     * 删除上边框
     * @return wtc
     */
    public WTc removeTopBorder(){
        removeBorder(src, "top");
        WTc top = top();
        if(null != top) {
            removeBorder(top.getSrc(), "bottom");
        }
        return this;
    }
    /**
     * 删除下边框
     * @return wtc
     */
    public WTc removeBottomBorder(){
        removeBorder(src, "bottom");
        WTc bottom = bottom();
        if(null != bottom) {
            removeBorder(bottom.getSrc(), "top");
        }
        return this;
    }

    /**
     * 删除左上至右下分隔线
     * @return wtc
     */
    public WTc removeTl2brBorder(){
        removeBorder(src, "tl2br");
        return this;
    }
    /**
     * 删除右上至左下分隔线
     * @return wtc
     */
    public WTc removeTr2blBorder(){
        removeBorder(src, "tr2bl");
        return this;
    }
    /**
     * 删除边框
     * @return wtc
     */
    private void removeBorder(Element tc, String side){
        Element tcPr = DocxUtil.addElement(tc, "tcPr");
        Element borders = DocxUtil.addElement(tcPr, "tcBorders");
        Element border = DocxUtil.addElement(borders, side);
        border.addAttribute("w:val","nil");
        DocxUtil.removeAttribute(border, "sz");
        DocxUtil.removeAttribute(border, "space");
        DocxUtil.removeAttribute(border, "color");
    }


    /**
     * 删除所有
     * @return wtc
     */
    public WTc removeBorder(){
        removeLeftBorder();
        removeRightBorder();
        removeTopBorder();
        removeBottomBorder();
        removeTl2brBorder();
        removeTr2blBorder();
        return this;
    }
    /**
     * 设置上下左右默认边框
     * @return wtc
     */
    public WTc setBorder(){
        setLeftBorder();
        setRightBorder();
        setTopBorder();
        setBottomBorder();
        return this;
    }

    /**
     * 设置上下左右边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTc setBorder(int size, String color, String style){
        setLeftBorder(size, color, style);
        setRightBorder(size, color, style);
        setTopBorder(size, color, style);
        setBottomBorder(size, color, style);
        setTl2brBorder(size, color, style);
        setTr2blBorder(size, color, style);
        return this;
    }
    /**
     * 设置左默认边框
     * @return wtc
     */
    public WTc setLeftBorder(){
        setBorder(src, "left", 4, "auto", "single");
        return this;
    }
    /**
     * 设置右默认边框
     * @return wtc
     */
    public WTc setRightBorder(){
        setBorder(src, "right", 4, "auto", "single");
        return this;
    }
    /**
     * 设置上默认边框
     * @return wtc
     */
    public WTc setTopBorder(){
        setBorder(src, "top", 4, "auto", "single");
        return this;
    }
    /**
     * 设置下默认边框
     * @return wtc
     */
    public WTc setBottomBorder(){
        setBorder(src, "bottom", 4, "auto", "single");
        return this;
    }
    /**
     * 设置左上至右下默认边框
     * @return wtc
     */
    public WTc setTl2brBorder(){
        setBorder(src, "tl2br", 4, "auto", "single");
        return this;
    }

    /**
     * 设置 左上 至 右下分隔线
     * @param top 右上内容
     * @param bottom 左下内容
     * @return wtc
     */
    public WTc setTl2brBorder(String top, String bottom){
        setBorder(src, "tl2br", 4, "auto", "single");
        String html = "<div style='text-align:right;'>"+top+"</div><div style='text-align:left;'>"+bottom+"</div>";
        setHtml(html);
        return this;
    }
    /**
     * 设置 左上 至 右下默认样式分隔线
     * @return wtc
     */
    public WTc setTr2blBorder(){
        setBorder(src, "tr2bl", 4, "auto", "single");
        return this;
    }
    /**
     * 设置 右上 至 左下分隔线
     * @param top 左上内容
     * @param bottom 右下内容
     * @return wtc
     */
    public WTc setTr2blBorder(String top, String bottom){
        setBorder(src, "tr2bl", 4, "auto", "single");
        String html = "<div style='text-align:left;'>"+top+"</div><div style='text-align:right;'>"+bottom+"</div>";
        setHtml(html);
        return this;
    }

    /**
     * 设置左边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTc setLeftBorder(int size, String color, String style){
        setBorder(src, "left", size, color, style);
        return this;
    }
    /**
     * 设置右边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTc setRightBorder(int size, String color, String style){
        setBorder(src, "right", size, color, style);
        return this;
    }
    /**
     * 设置上边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTc setTopBorder(int size, String color, String style){
        setBorder(src, "top", size, color, style);
        return this;
    }
    /**
     * 设置下边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTc setBottomBorder(int size, String color, String style){
        setBorder(src, "bottom", size, color, style);
        return this;
    }
    /**
     * 设置左上至右下边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTc setTl2brBorder(int size, String color, String style){
        setBorder(src, "tl2br", size, color, style);
        return this;
    }
    /**
     * 设置右上至左下边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTc setTr2blBorder(int size, String color, String style){
        setBorder(src, "tr2bl", size, color, style);
        return this;
    }
    private void setBorder(Element tc, String side, int size, String color, String style){
        Element tcPr = DocxUtil.addElement(tc, "tcPr");
        Element borders = DocxUtil.addElement(tcPr, "tcBorders");
        Element border = DocxUtil.addElement(borders, side);
        if(null == style) {
            style = "single";
        }
        DocxUtil.addAttribute(border, "val", style);
        DocxUtil.addAttribute(border, "sz", size+"");
        if(null != color) {
            DocxUtil.addAttribute(border, "color", color.replace("#", ""));
        }
        DocxUtil.addAttribute(border, "space", "0");
    }

    /**
     * 设置下边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setBottomPadding(String padding){
        return setPadding(src, "bottom", padding);
    }
    /**
     * 设置下边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setBottomPadding(int padding){
        return setPadding(src, "bottom", padding);
    }

    /**
     * 设置下边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setBottomPadding(double padding){
        return setPadding(src, "bottom", padding);
    }

    /**
     * 设置上边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setTopPadding(String padding){
        return setPadding(src, "top", padding);
    }
    /**
     * 设置上边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setTopPadding(int padding){
        return setPadding(src, "top", padding);
    }
    /**
     * 设置上边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setTopPadding(double padding){
        return setPadding(src, "top", padding);
    }

    /**
     * 设置右边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setRightPadding(String padding){
        return setPadding(src, "right", padding);
    }
    /**
     * 设置右边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setRightPadding(int padding){
        return setPadding(src, "right", padding);
    }
    /**
     * 设置右边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setRightPadding(double padding){
        return setPadding(src, "right", padding);
    }

    /**
     * 设置左边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setLeftPadding(String padding){
        return setPadding(src, "left", padding);
    }
    /**
     * 设置左边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setLeftPadding(int padding){
        return setPadding(src, "left", padding);
    }
    /**
     * 设置左边距
     * @param padding 边距
     * @return wtc
     */
    public WTc setLeftPadding(double padding){
        return setPadding(src, "left", padding);
    }


    public WTc setPadding(String side, String padding){
        return setPadding(src, side, padding);
    }
    public WTc setPadding(String side, int padding){
        return setPadding(src, side, padding);
    }
    public WTc setPadding(String side, double padding){
        return setPadding(src, side, padding);
    }


    public WTc setPadding(String padding){
        setPadding(src, "top", padding);
        setPadding(src, "bottom", padding);
        setPadding(src, "right", padding);
        setPadding(src, "left", padding);
        return this;
    }
    public WTc setPadding(int padding){
        return setPadding(padding+widthUnit);
    }
    public WTc setPadding(double padding){
        return setPadding(padding+widthUnit);
    }


    private WTc setPadding(Element tc, String side, int padding){
        return setPadding(tc, side, padding+widthUnit);
    }
    private WTc setPadding(Element tc, String side, double padding){
        return setPadding(tc, side, padding+widthUnit);
    }
    private WTc setPadding(Element tc, String side, String padding){
        Element pr = DocxUtil.addElement(tc, "tcPr");
        Element mar = DocxUtil.addElement(pr,"tcMar");
        DocxUtil.addElement(mar,side,"w",DocxUtil.dxa(padding)+"");
        DocxUtil.addElement(mar,side,"type","dxa");
        return this;
    }
    public WTc setColor(String color){
        for(WParagraph wp:wps){
            wp.setColor(color);
        }
        return this;
    }
    public WTc setFont(String size, String eastAsia, String ascii, String hint){
        for(WParagraph wp:wps){
            wp.setFont(size, eastAsia, ascii, hint);
        }
        return this;
    }
    public WTc setFontSize(String size){
        for(WParagraph wp:wps){
            wp.setFontSize(size);
        }
        return this;
    }
    public WTc setFontFamily(String font){
        for(WParagraph wp:wps){
            wp.setFontFamily(font);
        }
        return this;
    }

    public WTc setWidth(String width){
        Element pr = DocxUtil.addElement(src, "tcPr");
        DocxUtil.addElement(pr, "tcW","w", DocxUtil.dxa(width)+"");
        DocxUtil.addElement(pr, "tcW","type", DocxUtil.widthType(width));
        return this;
    }
    public WTc setWidth(int width){
        return setWidth(widthUnit+widthUnit);
    }
    public WTc setWidth(double width){
        return setWidth(widthUnit+widthUnit);
    }

    public WTc setAlign(String align){
        Element pr = DocxUtil.addElement(src, "tcPr");
        DocxUtil.addElement(pr, "jc","val", align);
        for(WParagraph wp:wps){
            wp.setAlign(align);
        }
        return this;
    }
    public WTc setVerticalAlign(String align){
        Element pr = DocxUtil.addElement(src, "tcPr");
        if(align.equals("middle")){
            align = "center";
        }
        DocxUtil.addElement(pr,"vAlign", "val", align );
        return this;
    }

    /**
     * 背景色
     * @param color 颜色
     * @return Wtc
     */
    public WTc setBackgroundColor(String color){
        Element pr = DocxUtil.addElement(src, "tcPr");
        DocxUtil.addElement(pr, "shd", "color","auto");
        DocxUtil.addElement(pr, "shd", "val","clear");
        DocxUtil.addElement(pr, "shd", "fill",color.replace("#",""));
        for(WParagraph wp:wps){
            // wp.setBackgroundColor(color);
        }
        return this;
    }

    /**
     * 清除样式
     * @return Wtc
     */
    public WTc removeStyle(){
        Element pr = src.element("tcPr");
        if(null != pr){
            src.remove(pr);
        }
        for(WParagraph wp:wps){
            wp.removeStyle();
        }
        return this;
    }
    /**
     * 清除背景色
     * @return Wtc
     */
    public WTc removeBackgroundColor(){
        DocxUtil.removeElement(src,"shd");
        return this;
    }

    /**
     * 清除颜色
     * @return wtc
     */
    public WTc removeColor(){
        DocxUtil.removeElement(src,"color");
        return this;
    }
    /**
     * 粗体
     * @param bold 是否
     * @return Wtc
     */
    public WTc setBold(boolean bold){
        for(WParagraph wp:wps){
            wp.setBold(bold);
        }
        return this;
    }
    public WTc setBold(){
        setBold(true);
        return this;
    }

    /**
     * 下划线
     * @param underline 是否
     * @return Wtc
     */
    public WTc setUnderline(boolean underline){
        for(WParagraph wp:wps){
            wp.setUnderline(underline);
        }
        return this;
    }
    public WTc setUnderline(){
        setUnderline(true);
        return this;
    }

    /**
     * 删除线
     * @param strike 是否
     * @return Wtc
     */
    public WTc setStrike(boolean strike){
        for(WParagraph wp:wps){
            wp.setStrike(strike);
        }
        return this;
    }
    public WTc setStrike(){
        setStrike(true);
        return this;
    }

    /**
     * 斜体
     * @param italic 是否
     * @return Wtc
     */
    public WTc setItalic(boolean italic){
        for(WParagraph wp:wps){
            wp.setItalic(italic);
        }
        return this;
    }

    public WTc setItalic(){
        return setItalic(true);
    }
    public List<WParagraph> getWps(){
        return wps;
    }
    public WTc setHtml(String html){

        DocxUtil.removeContent(src);
        try {
            if(root.IS_HTML_ESCAPE){
                html = HtmlUtil.name2code(html);
            }
            Document doc = DocumentHelper.parseText("<root>"+html+"</root>");
            Element root = doc.getRootElement();
            this.root.parseHtml(src, null, root, null, false);
        }catch (Exception e){
            e.printStackTrace();
        }
        return this;
    }
    public WTc setHtml(Element html){
        String tag = html.getName();
        DocxUtil.removeContent(src);
        List<Element> elements = html.elements();
        if(html.elements().size()>0){
            root.block(src, null, html, null);
        }else{
            setText(html.getText(), StyleParser.parse(html.attributeValue("style")));
        }
        return this;
    }
    public void remove(){
        src.getParent().remove(src);
        parent.getTcs().remove(this);
    }
    public WTc setText(String text){
        setText(text, null);
        return this;
    }
    public WTc setText(String text, Map<String, String> styles){
        DocxUtil.removeContent(src);
        Element p = DocxUtil.addElement(src, "p");
        Element r = DocxUtil.addElement(p, "r");
        Element t = DocxUtil.addElement(r, "t");
        if(root.IS_HTML_ESCAPE) {
            text = HtmlUtil.display(text);
        }
        DocxUtil.pr(r, styles);
        t.setText(text);
        return this;
    }
    public WTc addText(String text){
        Element p = DocxUtil.addElement(src, "p");
        Element r = DocxUtil.addElement(p, "r");
        Element t = r.addElement("w:t");
        if(root.IS_HTML_ESCAPE) {
            text = HtmlUtil.display(text);
        }
        t.setText(text);
        return this;
    }


    /**
     * 文本替换，不限层级查的所有t标签
     * @param target 查找target
     * @param replacement 替换成replacement
     * @return Wtc
     */
    public WTc replace(String target, String replacement){
        for(WParagraph wp:wps){
            wp.replace(target, replacement);
        }
        return this;
    }

    public LinkedHashMap<String, String> styles(){
        LinkedHashMap<String, String> styles = new LinkedHashMap<>();
        Element pr = src.element("tcPr");
        if(null != pr){
            //<w:tcW w:w="3436" w:type="dxa"/>
            Element w = pr.element("tcW");
            if(null != w){
                int width = BasicUtil.parseInt(w.attributeValue("w"), 0);
                if(width > 0){
                    styles.put("width", (int)DocxUtil.dxa2px(width)+"px");
                }
            }
            /*<w:tcBorders>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="C0504D" w:themeColor="accent2"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="C0504D" w:themeColor="accent2"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="C0504D" w:themeColor="accent2"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="C0504D" w:themeColor="accent2"/>
            </w:tcBorders>
            */
            //<w:shd w:val="pct5" w:color="92D050" w:fill="auto"/>
            Element shd = src.element("shd");
            if(null != shd){
                String color = color(shd.attributeValue("color"));
                if(null != color){
                    styles.put("background-color", color);
                }
            }
            //<w:vAlign w:val="center"/>
            Element valign = pr.element("vAlign");
            if(null != valign){
                String val = valign.attributeValue("val");
                if(null != val){
                    if(val.equalsIgnoreCase("center")){
                        val = "middle";
                    }
                    styles.put("vertical-align", val);
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
        int colspan = getColspan();
        int rowspan = getRowspan();
        while (items.hasNext()){
            Element item = items.next();
            String tag = item.getName();
            if(tag.equalsIgnoreCase("p")){
                body.append("\n");
                body.append(new WParagraph(getDoc(),  item).html(uploader, lvl+1));
            }else if(tag.equalsIgnoreCase("r")){
                body.append("\n");
                body.append(new WRun(getDoc(),  item).html(uploader, lvl+1));
            }else if(tag.equalsIgnoreCase("tbl")){
                body.append("\n");
                body.append(new WTable(getDoc(),  item).html(uploader, lvl+1));
            }else if(tag.equalsIgnoreCase("t")){
                body.append("\n");
                t(builder, lvl+1);
                body.append(item.getText());
            }
        }
        t(builder, lvl);
        builder.append("<td");
        styles(builder);
        if(colspan > 1){
            builder.append(" colspan='").append(colspan).append("'");
        }
        if(rowspan > 1){
            builder.append(" rowspan='").append(rowspan).append("'");
        }
        builder.append(">");
        builder.append(body);
        builder.append("\n");
        t(builder, lvl);
        builder.append("</td>");
        return builder.toString();
    }
    /**
     * 复制一列
     * @param content 是否复制其中内容
     * @return wtr
     */
    public WTc clone(boolean content){
        WTc tc = new WTc(root, parent, this.getSrc().createCopy());
        if(!content){
            tc.removeContent();
        }
        return tc;
    }
    public String toString(){
        return DocxUtil.text(src);
    }
}