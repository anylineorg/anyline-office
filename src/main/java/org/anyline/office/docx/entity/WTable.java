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

import org.anyline.entity.html.TableBuilder;
import org.anyline.handler.Uploader;
import org.anyline.office.docx.util.DocxUtil;
import org.anyline.util.*;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;

import java.util.*;

public class WTable extends WElement {
    private String widthUnit = "px";     // 默认长度单位 px pt cm/厘米
    private List<WTr> wtrs = new ArrayList<>();
    //是否自动同步(根据word源码重新构造 wtable wtr wtc)
    //在大批量操作时需要关掉自动同步,在操作完成后调用一次 reload()
    private boolean isAutoLoad = true;
    public WTable(WDocument doc){
        this.root = doc;
        load();
    }
    public WTable(WDocument doc, Element src){
        this.root = doc;
        this.src = src;
        load();
    }
    public void reload(){
        load();
    }
    private void load(){
        wtrs.clear();
        List<Element> elements = src.elements("tr");
        for(Element element:elements){
            WTr tr = new WTr(root, this, element);
            wtrs.add(tr);
        }
    }

    /**
     * 根据书签或点位符获取行
     * @param bookmark 书签或占位符 包含{和}的按占位符搜索
     * @return wtr
     */
    public WTr tr(String bookmark){
        List<WTr> trs = trs(bookmark);
        if(!trs.isEmpty()){
            return trs.get(0);
        }
        return null;
    }
    public List<WTr> trs(String bookmark){
        List<WTr> list = new ArrayList<>();
        if(null != bookmark) {
            if(bookmark.contains("{") && bookmark.contains("}")){
                for(WTr item:wtrs){
                    String txt = item.getTexts();
                    if(txt.contains(bookmark)){
                        list.add(item);
                    }
                }
            }else {
                Element src = parent(bookmark, "tr");
                WTr tr = new WTr(root, this, src);
                list.add(tr);
            }
        }
        return list;
    }

    /**
     * 根据书签或点位符获取列
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
            for(WTr item:wtrs){
                list.addAll(item.tcs(bookmark));
            }
        }
        return list;
    }

    public Element parent(String bookmark, String tag){
        return root.parent(bookmark, tag);
    }


    /**
     * 创建行 并复制模板样式
     * @param template 模板
     * @param src 根据src创建(html标签)
     * @return tr
     */
    private WTr tr(WTr template, Element src){
        WTr tr = new WTr(root, this, template.getSrc().createCopy());
        tr.removeContent();
        List<Element> tds = src.elements("td");
        for(int i=0; i<tds.size(); i++){
            WTc wtc = tr.getTc(i);
            Element td = tds.get(i);
            Map<String, String> styles = StyleParser.parse(td.attributeValue("style"));
            wtc.setHtml(td);
            // this.doc.block(tc, null, td, null);
            /*Element t = DomUtil.element(tc,"t");
            if(null == t){
                t = tc.element("p").addElement("w:r").addElement("w:t");
            }
            String text = td.getTextTrim();
            t.setText(td.getTextTrim());*/

        }
        return tr;
    }
    private WTr tr(Element src){
        Element tr = this.src.addElement("w:tr");
        WTr wtr = new WTr(this.root, this, tr);
        List<Element> tds = src.elements("td");
        for(int i=0; i<tds.size(); i++){
            Element tc = tr.addElement("w:tc");
            WTc wtc = new WTc(root, wtr, tc);
            Element td = tds.get(i);
            wtc.setHtml(td);
        }
        return wtr;
    }

    /**
     * 获取模板行
     * @param index 插入位置下标 负数表示倒数第index行 插入 null表示从最后追加与append效果一致
     * @return Wtr
     */
    public WTr template(Integer index){
        WTr template = null;
        int size = wtrs.size();
        if(size>0){
            if(null == index){
                template = wtrs.get(size-1);
            }else {
                index = index(index, size);
                template = wtrs.get(index);
            }

        }
        return template;
    }

    /**
     * 在最后位置插入一行
     * @param html html.tr源码
     */
    public void insert(String html){
        Integer index = null;
        insert(index, html);
    }

    /**
     * 在最后位置插入一行,半填充内容
     * 内容从data中获取
     * @param data DataRow/Map/Entity
     * @param cols data的属性
     */
    public void insert(Object data, String ... cols){
        Integer index = null;
        WTr template = template(index);
        insert(index, template, data, cols);
    }
    public void append(Object data, String ... cols){
        Integer index = null;
        WTr template = template(index);
        insert(index, template, data, cols);
    }

    /**
     * 在index位置插入一行,原来index位置的行被挤到下一行,并填充内容
     * 内容从data中获取
     * @param index 插入位置下标 负数表示倒数第index行 插入 null表示从最后追加与append效果一致
     * @param data DataRow/Map/Entity
     * @param cols data的属性
     */
    public void insert(Integer index, Object data, String ... cols){
        WTr template = template(index);
        insert(index, template, data, cols);
    }

    /**
     * 根据模版样式和数据 插入行
     * @param template 模版行
     * @param data 数据可以是一个实体也可以是一个集合
     * @param fields 指定从数据中提取的数据的属性或key
     */
    public void insert(WTr template, Object data, String ... fields){
        insert(null, template, data, fields);
    }
    public void append(WTr template, Object data, String ... fields){
        insert(null, template, data, fields);
    }

    public void insert(Integer index, WTr tr){
        List<Element> trs = src.elements("tr");
        index = index(index, wtrs.size());
        if(index > 0){
            wtrs.add(index, tr);
            trs.add(index, tr.getSrc());
        }else{
            wtrs.add(tr);
            trs.add(tr.getSrc());
        }
    }
    public void add(WTr tr){
        List<Element> trs = src.elements("tr");
        wtrs.add(tr);
        trs.add(tr.getSrc());
    }
    /**
     * 根据模版样式和数据 插入行,原来index位置的行被挤到下一行
     * @param index 插入位置下标 负数表示倒数第index行 插入 null表示从最后追加与append效果一致
     * @param template 模版行
     * @param data 数据可以是一个实体也可以是一个集合
     * @param fields 指定从数据中提取的数据的属性或key
     */
    public void insert(Integer index, WTr template, Object data, String ... fields){
        Collection datas = null;
        if(data instanceof Collection){
            datas = (Collection)data;
        }else{
            datas = new ArrayList();
            datas.add(data);
        }
        TableBuilder builder = TableBuilder.init().setFields(fields).setDatas(datas);
        String html = builder.build().build(false);
        insert(index, template, html);
    }

    /**
     * 插入行,原来index位置的行被挤到下一行,并填充内容
     * @param index 插入位置下标 负数表示倒数第index行 插入 null表示从最后追加与append效果一致
     * @param tds 每列的文本 数量多于表格列的 条目无效
     */
    public void insert(Integer index, List<String> tds){
        int size = NumberUtil.min(tds.size(), wtrs.get(0).getTcs().size());
        StringBuilder builder = new StringBuilder();
        builder.append("<tr>");
        for(int i=0; i<size; i++){
            builder.append("<td>");
            builder.append(tds.get(i));
            builder.append("</td>");
        }
        builder.append("</tr>");
        insert(index, builder.toString());
    }

    public void append(List<String> tds){
        insert(null, tds);
    }
    /**
     * 追加行
     * @param tds 每列的文本 数量多于表格列的 条目无效
     */
    public void insert(List<String> tds){
        Integer index = null;
        insert(index, tds);
    }
    /**
     * 在index位置插入行,原来index位置的行被挤到下一行,并填充内容
     * @param index 插入位置下标 负数表示倒数第index行 插入 null表示从最后追加与append效果一致
     * @param tds 每列的文本 数量多于表格列的 条目无效
     */
    public void insert(Integer index, String ... tds){
        insert(index, BeanUtil.array2list(tds));
    }
    /**
     * 追加行,并填充内容
     * @param tds 每列的文本 数量多于表格列的 条目无效
     */
    public void insert(String ... tds){
        Integer index = null;
        insert(index, tds);
    }
    public void append(String ... tds){
        Integer index = null;
        insert(index, tds);
    }
    /**
     * 在index位置插入行,原来index位置的行被挤到下一行,以template为模板
     * @param index 插入位置下标 负数表示倒数第index行 插入 null表示从最后追加与append效果一致
     * @param template 模板
     * @param qty 插入数量
     * @return Wtable
     */
    public WTable insert(Integer index, WTr template, int qty){
        List<Element> trs = src.elements("tr");
        int idx = index(index, trs.size());
        for(int i=0; i<qty; i++) {
            Element newTr = template.getSrc().createCopy();
            DocxUtil.removeContent(newTr);
            if(null == index) {
                trs.add(newTr);
            }else{
                trs.add(idx++, newTr);
            }
        }
        reload();
        return this;
    }
    public WTable append(WTr template, int qty){
        return insert(null, template, qty);
    }
    /**
     * 在index位置插入qty行,以原来index位置行为模板,原来index位置以下行的挤到下一行
     * @param index 插入位置下标 负数表示倒数第index行 插入 null表示从最后追加与append效果一致
     * @param qty 插入数量
     * @return Wtable
     */
    public WTable insert(Integer index, int qty){
        int size = wtrs.size();
        if(size > 0){
            WTr template = template(index);
            return insert(index, template, qty);
        }
        return this;
    }
    /**
     * 在最后位置插入qty行,以最后一行为模板
     * @param qty 插入数量
     * @return Wtable
     */
    public WTable insert(int qty){
        return insert(null, qty);
    }
    public WTable append(int qty){
        return insert(null, qty);
    }

    /**
     * 在index位置插入1行,以原来index位置行为模板,原来index位置以下行的挤到下一行
     * @param index 插入位置下标 负数表示倒数第index行 插入 null表示从最后追加与append效果一致
     * @param html html内容
     */
    public void insert(Integer index, String html){
        List<Element> trs = src.elements("tr");
        WTr template = template(index); //取原来在当前位置的一行作模板
        insert(index, template, html);
    }
    public void append(String html){
        Integer index = null;
        insert(index, html);
    }

    /**
     * 插入行 如果模板位于当前表中则从当前模板位置往后插入,否则插入到最后一行
     * @param template 模板
     * @param html html.tr源码
     */
    public void insert(WTr template, String html){
        Integer index = null;
        if(null != template) {
            List<Element> trs = src.elements("tr");
            index = trs.indexOf(template.getSrc());
        }
        insert(index, template,  html);
    }
    public void append(WTr template, String html){
        insert(template, html);
    }

    /**
     * 根据模版样式 插入行
     * @param index 插入位置下标 负数表示倒数 插入 null表示从最后追加与append效果一致
     * @param template 模版行
     * @param html html片段 片段中应该有多个tr,不需要上级标签table
     */
    public void insert(Integer index, WTr template, String html){
        List<Element> trs = src.elements("tr");
        int idx = index(index, trs.size());

        /*
        if(index == -1 && null != template){
            index = trs.indexOf(template.getSrc());
        }
        */
        try {
            if(root.IS_HTML_ESCAPE){
                html = HtmlUtil.name2code(html);
            }
            org.dom4j.Document doc = DocumentHelper.parseText("<root>"+html+"</root>");
            Element root = doc.getRootElement();
            List<Element> rows = root.elements("tr");
            for(Element row:rows){
                Element newTr = null;
                if(null != template) {
                    newTr = tr(template, row).getSrc();
                }else{
                    newTr = tr(row).getSrc();
                    trs.remove(newTr);
                }
                if(null == index){
                    trs.add(newTr);
                }else {
                    trs.add(idx++, newTr);
                }
            }
            if(isAutoLoad) {
                reload();
            }
        }catch (Exception e){
            e.printStackTrace();
        }

    }


    /**
     * 删除行
     * @param index  下标从0开始  负数表示倒数第index行
     */
    public void remove(int index){
        List<Element> trs = src.elements("tr");
        if(trs.size() == 0){
            return;
        }
        index = index(index, trs.size());
        trs.remove(index);
        if(isAutoLoad) {
            reload();
        }
    }
    public void remove(WTr tr){
        List<Element> trs = src.elements("tr");
        trs.remove(tr.getSrc());
        if(isAutoLoad) {
            reload();
        }
    }

    /**
     * 获取row行col列的文本
     * @param row 行
     * @param col 列
     * @return String
     */
    public String getText(int row, int col){
        String text = null;
        List<Element> trs = src.elements("tr");
        Element tr = trs.get(row);
        List<Element> tcs = tr.elements("tc");
        Element tc = tcs.get(col);
        text = DocxUtil.text(tc);
        return text;
    }

    /**
     * 设置row行col列的文本
     * @param row 行
     * @param col 列
     * @param text 内容 不支持html标签 如果需要html标签 调用setHtml()
     * @return wtable
     */
    public WTable setText(int row, int col, String text){
        return setText(row, col, text, null);
    }
    /**
     * 在row行col列原有基础上追加文本
     * @param row 行
     * @param col 列
     * @param text 内容 不支持html标签 如果需要html标签 调用setHtml()
     * @return wtable
     */
    public WTc addText(int row, int col, String text){
        return getTc(row, col).addText(text);
    }

    /**
     *
     * 设置row行col列的文本 并设置样式
     * @param row 行
     * @param col 列
     * @param text 内容 不支持html标签 如果需要html标签 调用setHtml()
     * @param styles css样式
     * @return wtable
     */
    public WTable setText(int row, int col, String text, Map<String, String> styles){
        WTc tc = getTc(row, col);
        if(null != tc){
            if(root.IS_HTML_ESCAPE) {
                text = HtmlUtil.display(text);
            }
            tc.setText(text, styles);
        }
        return this;
    }
    /**
     * 设置row行col列的文本 支持html标签
     * @param row 行
     * @param col 列
     * @param html 内容
     * @return wtable
     */
    public WTable setHtml(int row, int col, String html){
        WTc tc = getTc(row, col);
        if(null != tc) {
            tc.setHtml(html);
        }
        return this;
    }

    /**
     * 追加列, 每一行追加,追加的列将复制前一列的样式(背景色、字体等)
     * @param qty 追加数量
     * @return table table
     */
    public WTable addColumns(int qty){
        insertColumns(-1, qty);
        return this;
    }

    /**
     * 插入列
     * 追加的列将复制前一列的样式(背景色、字体等)
     * 如果col=0则将制后一列的样式(背景色、字体等)
     * @param col 插入位置 -1:表示追加以最后一行
     * @param qty 数量
     * @return table table
     */
    public WTable insertColumns(int col, int qty){
        List<Element> trs = src.elements("tr");
        for(Element tr:trs){
            List<Element> tcs = tr.elements("tc");
            int cols = tcs.size();
            if(cols > 0 && col < cols){
                Element template = null;
                if(col == 0){
                    template = tcs.get(0);
                }else if(col == -1){
                    template = tcs.get(cols-1);
                }else{
                    template = tcs.get(col-1);
                }
                int index = col;
                for (int i = 0; i < qty; i++) {
                    Element newTc = template.createCopy();
                    DocxUtil.removeContent(newTc);
                    if(col == -1){//追加到最后
                        tcs.add(newTc);
                    }else {
                        tcs.add(index++, newTc);
                    }
                }
            }else {
                for (int i = 0; i < qty; i++) {
                    tr.addElement("w:tc").addElement("w:p");
                }
            }
        }
        if(isAutoLoad) {
            reload();
        }
        return this;
    }
    /**
     * 追加行,追加的行将复制上一行的样式(背景色、字体等)
     * @param index 插入位置下标 负数表示倒数第index行 插入 null表示从最后追加与append效果一致
     * @param qty 追加数量
     * @return table table
     */
    public WTable insertRows(Integer index, int qty){
        if(wtrs.size()>0){
            insertRows(template(index), index, qty);
        }
        return this;
    }

    /**
     * 以template为模板 在index位置插入qty行,以原来index位置行为模板,原来index位置以下行的挤到下一行
     * @param template 模板
     * @param index 插入位置
     * @param qty 插入数量
     * @return wtable
     */
    public WTable insertRows(WTr template, Integer index, int qty){
        List<Element> trs = src.elements("tr");
        int idx = index(index, trs.size());
        if(trs.size()>0){
            for(int i=0; i<qty; i++) {
                Element newTr = template.getSrc().createCopy();
                DocxUtil.removeContent(newTr);
                if(null == index){
                    trs.add(newTr);
                }else {
                    trs.add(index++, newTr);
                }
            }
        }
        if(isAutoLoad) {
            reload();
        }
        return this;
    }

    /**
     * 追加qty行
     * @param qty 行数
     * @return tables
     */
    public WTable addRows(int qty){
        return insertRows(null, qty);
    }


    /**
     * 追加行,追加的行将复制上一行的样式(背景色、字体等)
     * @param index 位置  下标从0开始  负数表示倒数第index行
     * @param qty 追加数量
     * @return table table
     */
    public WTable addRows(int index, int qty){
        return insertRows(index, qty);
    }

    /**
     * 获取行数
     * @return int
     */
    public int getTrSize(){
        return src.elements("tr").size();
    }
    public WTable setWidth(String width){
        Element pr = DocxUtil.addElement(src, "tblPr");
        DocxUtil.addElement(pr, "tcW","w", DocxUtil.dxa(width)+"");
        DocxUtil.addElement(pr, "tcW","type", DocxUtil.widthType(width));
        return this;
    }

    /**
     * 设置表格宽度 默认px
     * @param width 宽度
     * @return wtable
     */
    public WTable setWidth(int width){
        return setWidth(width+widthUnit);
    }
    /**
     * 设置表格宽度 默认px
     * @param width 宽度
     * @return wtable
     */
    public WTable setWidth(double width){
        return setWidth(width+widthUnit);
    }

    /**
     * 合并行列
     * @param row 开始行
     * @param col 开始列
     * @param rowspan 合并行数量
     * @param colspan 合并列数量
     * @return wtable
     */
    public WTable merge(int row, int col, int rowspan, int colspan){
        reload();
        for(int r=row; r<row+rowspan; r++){
            for(int c=col; c<col+colspan; c++){
                WTc tc = getTc(r, c);
                Element pr = DocxUtil.addElement(tc.getSrc(), "tcPr");
                if(rowspan > 1){
                    if(r==row){
                        DocxUtil.addElement(pr, "vMerge", "val",   "restart");
                    }else{
                        DocxUtil.addElement(pr, "vMerge");
                    }
                }
                if(colspan>1){
                    if(c==col){
                        DocxUtil.addElement(pr, "gridSpan", "val",   colspan+"");
                    }else{
                        tc.remove();
                    }
                }
            }
        }
        reload();
        return this;
    }
    public List<WTr> getTrs(){
        return wtrs;
    }
    public WTr tr(int index){
        index = index(index, wtrs.size());
        return wtrs.get(index);
    }
    /**
     * 获取row行col列位置的单元格
     * @param row 行
     * @param col 列
     * @return wtc
     */
    public WTc getTc(int row, int col){
        WTr wtr = tr(row);
        if(null == wtr){
            return null;
        }
        return wtr.getTc(col);
    }

    public WTable removeBorder(){
        removeTopBorder();
        removeBottomBorder();
        removeLeftBorder();
        removeRightBorder();
        removeInsideHBorder();
        removeInsideVBorder();
        removeTl2brBorder();
        removeTr2blBorder();
        return this;
    }


    /**
     * 清除表格上边框
     * @return wtable
     */
    public WTable removeTopBorder(){
        removeBorder(src, "top");
        return this;
    }
    /**
     * 清除表格左边框
     * @return wtable
     */
    public WTable removeLeftBorder(){
        removeBorder(src, "left");
        return this;
    }
    /**
     * 清除表格右边框
     * @return wtable
     */
    public WTable removeRightBorder(){
        removeBorder(src, "right");
        return this;
    }
    /**
     * 清除表格下边框
     * @return wtable
     */
    public WTable removeBottomBorder(){
        removeBorder(src, "bottom");
        return this;
    }
    /**
     * 清除表格垂直边框
     * @return wtable
     */
    public WTable removeInsideVBorder(){
        removeBorder(src, "insideV");
        return this;
    }
    public WTable removeTl2brBorder(){
        removeBorder(src, "tl2br");
        return this;
    }
    public WTable removeTr2blBorder(){
        removeBorder(src, "tr2bl");
        return this;
    }

    /**
     * 清除表格水平边框
     * @return wtable
     */
    public WTable removeInsideHBorder(){
        removeBorder(src, "insideH");
        return this;
    }
    /**
     * 清除所有单元格边框
     * @return wtable
     */
    public WTable removeTcBorder(){
        for(WTr tr:wtrs){
            List<WTc> tcs = tr.getTcs();
            for(WTc tc:tcs){
                tc.removeBorder();
            }
        }
        return this;
    }

    /**
     * 清除所有单元格颜色
     * @return wtable
     */
    public WTable removeTcColor(){
        for(WTr tr:wtrs){
            List<WTc> tcs = tr.getTcs();
            for(WTc tc:tcs){
                tc.removeColor();
            }
        }
        return this;
    }

    /**
     * 清除所有单元格背景色
     * @return wtable
     */
    public WTable removeTcBackgroundColor(){
        for(WTr tr:wtrs){
            List<WTc> tcs = tr.getTcs();
            for(WTc tc:tcs){
                tc.removeBackgroundColor();
            }
        }
        return this;
    }


    private void removeBorder(Element tbl, String side){
        Element tcPr = DocxUtil.addElement(tbl, "tblPr");
        Element borders = DocxUtil.addElement(tcPr, "tblBorders");
        Element border = DocxUtil.addElement(borders, side);
        border.addAttribute("w:val","nil");
        DocxUtil.removeAttribute(border, "sz");
        DocxUtil.removeAttribute(border, "space");
        DocxUtil.removeAttribute(border, "color");
    }

    /**
     * 删除整行的上边框
     * @param row 行
     * @return Wtr
     */
    public WTr removeTopBorder(int row){
        WTr tr = tr(row);
        List<WTc> tcs = tr.getTcs();
        for(WTc tc:tcs){
            tc.removeTopBorder();
        }
        return tr;
    }

    /**
     * 删除整行的下边框
     * @param row 行
     * @return wtr
     */
    public WTr removeBottomBorder(int row){
        WTr tr = tr(row);
        tr.removeBottomBorder();
        return tr;
    }

    /**
     * 删除整列的左边框
     * @param col 列
     * @return Wtable
     */
    public WTable removeLeftBorder(int col){
        for(WTr tr: wtrs){
            WTc tc = tr.getTcWithColspan(col, true);
            if(null != tc){
                tc.removeLeftBorder();
            }
        }
        return this;
    }

    /**
     * 删除整列的右边框
     * @param col 列
     * @return Wtable
     */
    public WTable removeRightBorder(int col){
        for(WTr tr: wtrs){
            WTc tc = tr.getTcWithColspan(col, false);
            if(null != tc){
                tc.removeRightBorder();
            }
        }
        return this;
    }


    /**
     * 清除单元格左边框
     * @param row 行
     * @param col 列
     * @return Wtc
     */
    public WTc removeLeftBorder(int row, int col){
        return getTc(row, col).removeLeftBorder();
    }
    /**
     * 清除单元格右边框
     * @param row 行
     * @param col 列
     * @return Wtc
     */
    public WTc removeRightBorder(int row, int col){
        return getTc(row, col).removeRightBorder();
    }
    /**
     * 清除单元格上边框
     * @param row 行
     * @param col 列
     * @return Wtc
     */
    public WTc removeTopBorder(int row, int col){
        return getTc(row, col).removeTopBorder();
    }
    /**
     * 清除单元格下边框
     * @param row 行
     * @param col 列
     * @return Wtc
     */
    public WTc removeBottomBorder(int row, int col){
        return getTc(row, col).removeBottomBorder();
    }
    /**
     * 清除单元格左上到右下边框
     * @param row 行
     * @param col 列
     * @return wtable
     */
    public WTc removeTl2brBorder(int row, int col){
        return getTc(row, col).removeTl2brBorder();
    }
    /**
     * 清除单元格右上到左下边框
     * @param row 行
     * @param col 列
     * @return wtable
     */
    public WTc removeTr2blBorder(int row, int col){
        return getTc(row, col).removeBorder();
    }

    /**
     * 清除单元格所有边框
     * @param row 行
     * @param col 列
     * @return wtable
     */
    public WTc removeBorder(int row, int col){
        return getTc(row, col)
                .removeLeftBorder()
                .removeRightBorder()
                .removeTopBorder()
                .removeBottomBorder()
                .removeTl2brBorder()
                .removeTr2blBorder();
    }

    public WTr setBorder(int row){
        WTr tr = tr(row);
        tr.setBorder();
        return tr;
    }

    /**
     * 设置所有单元格默认边框
     * @return table table
     */
    public WTable setCellBorder(){
        for(WTr tr:wtrs){
            tr.setBorder();
        }
        return this;
    }
    /**
     * 设置单元格默认边框
     * @param row 行
     * @param col 列
     * @return  Wtc
     */
    public WTc setBorder(int row, int col){
        return getTc(row, col)
        .setLeftBorder()
        .setRightBorder()
        .setTopBorder()
        .setBottomBorder()
        .setTl2brBorder()
        .setTr2blBorder();
    }
    public WTc setBorder(int row, int col, int size, String color, String style){
        return getTc(row, col).setBorder(size, color, style);
    }
    public WTc setLeftBorder(int row, int col){
        return getTc(row, col).setLeftBorder();
    }
    public WTc setRightBorder(int row, int col){
        return getTc(row, col).setRightBorder();
    }
    public WTc setTopBorder(int row, int col){
        return getTc(row, col).setTopBorder();
    }
    public WTc setBottomBorder(int row, int col){
        return getTc(row, col).setBottomBorder();
    }
    public WTc setTl2brBorder(int row, int col){
        return getTc(row, col).setTl2brBorder();
    }
    public WTc setTl2brBorder(int row, int col, String top, String bottom){
        return getTc(row, col).setTl2brBorder(top, bottom);
    }
    public WTc setTr2blBorder(int row, int col){
        return getTc(row, col).setTr2blBorder();
    }

    public WTc setTr2blBorder(int row, int col, String top, String bottom){
        return getTc(row, col).setTr2blBorder(top, bottom);
    }

    public WTc setLeftBorder(int row, int col, int size, String color, String style){
        return getTc(row, col).setLeftBorder(size, color, style);
    }
    public WTc setRightBorder(int row, int col, int size, String color, String style){
        return getTc(row, col).setRightBorder(size, color, style);
    }
    public WTc setTopBorder(int row, int col, int size, String color, String style){
        return getTc(row, col).setTopBorder(size, color, style);
    }
    public WTc setBottomBorder(int row, int col, int size, String color, String style){
        return getTc(row, col).setBottomBorder(size, color, style);
    }
    public WTc setTl2brBorder(int row, int col, int size, String color, String style){
        return getTc(row, col).setTl2brBorder(size, color, style);
    }
    public WTc setTr2blBorder(int row, int col, int size, String color, String style){
        return getTc(row, col).setTr2blBorder(size, color, style);
    }


    /**
     * 设置所有行指定列的左边框
     * @param cols 列
     * @param size 边框宽度
     * @param color 颜色
     * @param style 样式
     * @return table table
     */
    public WTable setLeftBorder(int cols, int size, String color, String style){
        for(WTr tr:wtrs){
            tr.getTc(cols).setLeftBorder(size, color, style);
        }
        return this;
    }
    /**
     * 设置所有行指定列的右边框
     * @param cols 列
     * @param size 边框宽度
     * @param color 颜色
     * @param style 样式
     * @return table table
     */
    public WTable setRightBorder(int cols, int size, String color, String style){
        for(WTr tr:wtrs){
            tr.getTc(cols).setRightBorder(size, color, style);
        }
        return this;
    }

    /**
     * 设置整行所有单元格上边框
     * @param rows 行
     * @param size 边框宽度
     * @param color 颜色
     * @param style 样式
     * @return tr
     */
    public WTr setTopBorder(int rows, int size, String color, String style){
        return tr(rows).setTopBorder(size, color, style);
    }
    /**
     * 设置整行所有单元格下边框
     * @param rows 行
     * @param size 边框宽度
     * @param color 颜色
     * @param style 样式
     * @return tr
     */
    public WTr setBottomBorder(int rows, int size, String color, String style){
        return tr(rows).setBottomBorder(size, color, style);
    }


    /**
     * 设置表格左边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTable setLeftBorder(int size, String color, String style){
        setBorder("left", size, color, style);
        return this;
    }
    /**
     * 设置表格右边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTable setRightBorder(int size, String color, String style){
        setBorder("right", size, color, style);
        return this;
    }
    /**
     * 设置表格上边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTable setTopBorder(int size, String color, String style){
        setBorder("top", size, color, style);
        return this;
    }
    /**
     * 设置表格下边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTable setBottomBorder(int size, String color, String style){
        setBorder("bottom", size, color, style);
        return this;
    }
    /**
     * 设置表格边框
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     * @return wtc
     */
    public WTable setBorder(int size, String color, String style){
        setBorder("left", size, color, style);
        setBorder("top", size, color, style);
        setBorder("right", size, color, style);
        setBorder("bottom", size, color, style);
        return this;
    }
    /**
     * 设置表格边框
     * @param side left/top/right/bottom
     * @param size 边框宽度(1px)
     * @param color 颜色
     * @param style 样式(默认single)
     */
    private void setBorder(String side, int size, String color, String style){
        Element tcPr = DocxUtil.addElement(src, "tblPr");
        Element borders = DocxUtil.addElement(tcPr, "tblBorders");
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


    public WTc setColor(int row, int col, String color){
        return getTc(row, col).setColor(color);
    }

    /**
     * 设置整行颜色
     * @param rows 行
     * @param color 颜色
     * @return wtr
     */
    public WTr setColor(int rows, String color){
        WTr tr = tr(rows);
        tr.setColor(color);
        return tr;
    }
    /**
     * 设置单元格 字体
     * @param row 行
     * @param col 列
     * @param size 字号
     * @param eastAsia 中文字体
     * @param ascii 西文字体
     * @param hint 默认字体
     * @return wtc
     */
    public WTc setFont(int row, int col, String size, String eastAsia, String ascii, String hint){
        return getTc(row, col).setFont(size, eastAsia, ascii, hint);
    }

    /**
     * 设置整行 字体
     * @param row 行
     * @param size 字号
     * @param eastAsia 中文字体
     * @param ascii 西文字体
     * @param hint 默认字体
     * @return wtr
     */
    public WTr setFont(int row, String size, String eastAsia, String ascii, String hint){
        WTr tr = tr(row);
        tr.setFont(size, eastAsia, ascii, hint);
        return tr;
    }

    /**
     * 设置单元格字号
     * @param row 行
     * @param col 列
     * @param size 字号
     * @return wtc
     */
    public WTc setFontSize(int row, int col, String size){
        return getTc(row, col).setFontSize(size);
    }
    /**
     * 设置整行字号
     * @param rows 行
     * @param size 字号
     * @return wtr
     */
    public WTr setFontSize(int rows, String size){
        WTr tr = tr(rows);
        tr.setFontSize(size);
        return tr;
    }

    /**
     * 设置单元格字体
     * @param row 行
     * @param col 列
     * @param font 字体
     * @return wtc
     */
    public WTc setFontFamily(int row, int col, String font){
        return getTc(row, col).setFontFamily(font);
    }

    /**
     * 设置整行字体
     * @param rows 行
     * @param font 字体
     * @return wtr
     */
    public WTr setFontFamily(int rows, String font){
        WTr tr = tr(rows);
        tr.setFontFamily(font);
        return tr;
    }
    public WTc setWidth(int row, int col, String width){
        return getTc(row, col).setWidth(width);
    }
    public WTc setWidth(int row, int col, int width){
        return getTc(row, col).setWidth(width);
    }

    public WTc setWidth(int row, int col, double width){
        return getTc(row, col).setWidth(width);
    }

    /**
     * 设置整列宽度
     * @param cols 列
     * @param width 宽度
     * @return table table
     */
    public WTable setWidth(int cols, String width){
        for(WTr tr:wtrs){
            tr.getTc(cols).setWidth(width);
        }
        return this;
    }
    public WTable setWidth(int cols, int width){
        for(WTr tr:wtrs){
            tr.getTc(cols).setWidth(width);
        }
        return this;
    }
    public WTable setWidth(int cols, double width){
        for(WTr tr:wtrs){
            tr.getTc(cols).setWidth(width);
        }
        return this;
    }
    public WTr setHeight(int rows, String height){
        WTr tr = tr(rows);
        tr.setHeight(height);
        return tr;
    }

    public WTr setHeight(int rows, int height){
        return setHeight(rows, height+widthUnit);
    }
    public WTr setHeight(int rows, double height){
        return setHeight(rows, height+widthUnit);
    }

    /**
     * 设置单元格内容水平对齐方式
     * @param row 行
     * @param col 列
     * @param align 对齐方式
     * @return wtc
     */
    public WTc setAlign(int row, int col, String align){
        return getTc(row, col).setAlign(align);
    }
    /**
     * 设置整行单元格内容水平对齐方式
     * @param rows 行
     * @param align 对齐方式
     * @return wtcr
     */
    public WTr setAlign(int rows, String align){
        WTr tr = tr(rows);
        tr.setAlign(align);
        return tr;
    }

    /**
     * 设置整个表格单元格内容水平对齐方式
     * @param align 对齐方式
     * @return wtable
     */
    public WTable setAlign(String align){
        for(WTr tr:wtrs) {
            tr.setAlign(align);
        }
        return this;
    }
    /**
     * 设置单元格内容垂直对齐方式
     * @param row 行
     * @param col 列
     * @param align 对齐方式
     * @return wtc
     */
    public WTc setVerticalAlign(int row, int col, String align){
        return getTc(row, col).setVerticalAlign(align);
    }

    /**
     * 设置整行单元格内容垂直对齐方式
     * @param rows 行
     * @param align 对齐方式
     * @return wtr
     */
    public WTr setVerticalAlign(int rows, String align){
        WTr tr = tr(rows);
        tr.setVerticalAlign(align);
        return tr;
    }

    /**
     * 设置整个表格单元格内容垂直对齐方式
     * @param align 对齐方式
     * @return wtable
     */
    public WTable setVerticalAlign(String align){
        for(WTr tr:wtrs) {
            tr.setVerticalAlign(align);
        }
        return this;
    }
    /**
     * 设置单元格下边距
     * @param row 行
     * @param col 列
     * @param padding 边距 可以指定单位,如:10px
     * @return wtc
     */
    public WTc setBottomPadding(int row, int col, String padding){
        return getTc(row, col).setBottomPadding(padding);
    }
    /**
     * 设置单元格下边距
     * @param row 行
     * @param col 列
     * @param padding 边距 默认单位dxa
     * @return wtc
     */
    public WTc setBottomPadding(int row, int col, int padding){
        return getTc(row, col).setBottomPadding(padding);
    }
    public WTc setBottomPadding(int row, int col, double padding){
        return getTc(row, col).setBottomPadding(padding);
    }


    /**
     * 设置整行单元格下边距
     * @param rows 行
     * @param padding 边距 可以指定单位,如:10px
     * @return wtr
     */
    public WTr setBottomPadding(int rows, String padding){
        WTr tr = tr(rows);
        tr.setBottomPadding(padding);
        return tr;
    }
    public WTr setBottomPadding(int rows, int padding){
        WTr tr = tr(rows);
        tr.setBottomPadding(padding);
        return tr;
    }
    public WTr setBottomPadding(int rows, double padding){
        WTr tr = tr(rows);
        tr.setBottomPadding(padding);
        return tr;
    }
    /**
     * 设置整个表格中所有单元格下边距
     * @param padding 边距 可以指定单位,如:10px
     * @return wtable
     */
    public WTable setBottomPadding(String padding){
        for(WTr tr:wtrs){
            tr.setBottomPadding(padding);
        }
        return this;
    }
    public WTable setBottomPadding(int padding){
        for(WTr tr:wtrs){
            tr.setBottomPadding(padding);
        }
        return this;
    }
    public WTable setBottomPadding(double padding){
        for(WTr tr:wtrs){
            tr.setBottomPadding(padding);
        }
        return this;
    }

    public WTc setTopPadding(int row, int col, String padding){
        return getTc(row, col).setTopPadding(padding);
    }
    public WTc setTopPadding(int row, int col, int padding){
        return getTc(row, col).setTopPadding(padding);
    }
    public WTc setTopPadding(int row, int col, double padding){
        return getTc(row, col).setTopPadding(padding);
    }

    public WTr setTopPadding(int rows, String padding){
        WTr tr = tr(rows);
        tr.setTopPadding(padding);
        return tr;
    }
    public WTr setTopPadding(int rows, int padding){
        WTr tr = tr(rows);
        tr.setTopPadding(padding);
        return tr;
    }
    public WTr setTopPadding(int rows, double padding){
        WTr tr = tr(rows);
        tr.setTopPadding(padding);
        return tr;
    }
    public WTable setTopPadding(String padding){
        for(WTr tr:wtrs){
            tr.setTopPadding(padding);
        }
        return this;
    }
    public WTable setTopPadding(int padding){
        for(WTr tr:wtrs){
            tr.setTopPadding(padding);
        }
        return this;
    }
    public WTable setTopPadding(double padding){
        for(WTr tr:wtrs){
            tr.setTopPadding(padding);
        }
        return this;
    }
    public WTc setRightPadding(int row, int col, String padding){
        return getTc(row, col).setRightPadding(padding);
    }
    public WTc setRightPadding(int row, int col, int padding){
        return getTc(row, col).setRightPadding(padding);
    }
    public WTc setRightPadding(int row, int col, double padding){
        return getTc(row, col).setRightPadding(padding);
    }

    public WTr setRightPadding(int rows, String padding){
        WTr tr = tr(rows);
        tr.setRightPadding(padding);
        return tr;
    }
    public WTr setRightPadding(int rows, int padding){
        WTr tr = tr(rows);
        tr.setRightPadding(padding);
        return tr;
    }
    public WTr setRightPadding(int rows, double padding){
        WTr tr = tr(rows);
        tr.setRightPadding(padding);
        return tr;
    }
    public WTable setRightPadding(String padding){
        for(WTr tr:wtrs){
            tr.setRightPadding(padding);
        }
        return this;
    }
    public WTable setRightPadding(int padding){
        for(WTr tr:wtrs){
            tr.setRightPadding(padding);
        }
        return this;
    }
    public WTable setRightPadding(double padding){
        for(WTr tr:wtrs){
            tr.setRightPadding(padding);
        }
        return this;
    }


    public WTc setLeftPadding(int row, int col, String padding){
        return getTc(row, col).setLeftPadding(padding);
    }
    public WTc setLeftPadding(int row, int col, int padding){
        return getTc(row, col).setLeftPadding(padding);
    }
    public WTc setLeftPadding(int row, int col, double padding){
        return getTc(row, col).setLeftPadding(padding);
    }

    public WTr setLeftPadding(int rows, String padding){
        WTr tr = tr(rows);
        tr.setLeftPadding(padding);
        return tr;
    }
    public WTr setLeftPadding(int rows, int padding){
        WTr tr = tr(rows);
        tr.setLeftPadding(padding);
        return tr;
    }
    public WTr setLeftPadding(int rows, double padding){
        WTr tr = tr(rows);
        tr.setLeftPadding(padding);
        return tr;
    }

    public WTable setLeftPadding(String padding){
        for(WTr tr:wtrs){
            tr.setLeftPadding(padding);
        }
        return this;
    }
    public WTable setLeftPadding(int padding){
        for(WTr tr:wtrs){
            tr.setLeftPadding(padding);
        }
        return this;
    }
    public WTable setLeftPadding(double padding){
        for(WTr tr:wtrs){
            tr.setLeftPadding(padding);
        }
        return this;
    }



    public WTc setPadding(int row, int col, String side, String padding){
        return getTc(row, col).setPadding(side, padding);
    }
    public WTc setPadding(int row, int col, String side, int padding){
        return getTc(row, col).setPadding(side, padding);
    }
    public WTc setPadding(int row, int col, String side, double padding){
        return getTc(row, col).setPadding(side, padding);
    }
    public WTr setPadding(int rows, String side, String padding){
        WTr tr = tr(rows);
        tr.setPadding(side, padding);
        return tr;
    }
    public WTr setPadding(int rows, String side, int padding){
        WTr tr = tr(rows);
        tr.setPadding(side, padding);
        return tr;
    }
    public WTr setPadding(int rows, String side, double padding){
        WTr tr = tr(rows);
        tr.setPadding(side, padding);
        return tr;
    }
    public WTable setPadding(String side, String padding){
        for(WTr tr:wtrs){
            tr.setPadding(side, padding);
        }
        return this;
    }

    public WTable setPadding(String side, int padding){
        for(WTr tr:wtrs){
            tr.setPadding(side, padding);
        }
        return this;
    }

    public WTable setPadding(String side, double padding){
        for(WTr tr:wtrs){
            tr.setPadding(side, padding);
        }
        return this;
    }



    public WTc setPadding(int row, int col, String padding){
        return getTc(row, col).setPadding(padding);
    }
    public WTc setPadding(int row, int col, int padding){
        return getTc(row, col).setPadding(padding);
    }
    public WTc setPadding(int row, int col, double padding){
        return getTc(row, col).setPadding(padding);
    }
    public WTr setPadding(int rows, String padding){
        WTr tr = tr(rows);
        tr.setPadding(padding);
        return tr;
    }
    public WTr setPadding(int rows, int padding){
        WTr tr = tr(rows);
        tr.setPadding(padding);
        return tr;
    }
    public WTr setPadding(int rows, double padding){
        WTr tr = tr(rows);
        tr.setPadding(padding);
        return tr;
    }
    public WTable setPadding(String padding){
        for(WTr tr:wtrs){
            tr.setPadding(padding);
        }
        return this;
    }

    public WTable setPadding(int padding){
        for(WTr tr:wtrs){
            tr.setPadding(padding);
        }
        return this;
    }

    public WTable setPadding(double padding){
        for(WTr tr:wtrs){
            tr.setPadding(padding);
        }
        return this;
    }


    /**
     * 设置单元格背景色
     * @param row 行
     * @param col 列
     * @param color 颜色
     * @return Wtc
     */
    public WTc setBackgroundColor(int row, int col, String color){
        return getTc(row, col).setBackgroundColor(color);
    }

    /**
     * 设置整行单元格背景色
     * @param row 行
     * @param color 颜色
     * @return Wtr
     */
    public WTr setBackgroundColor(int row, String color){
        WTr tr = tr(row);
        tr.setBackgroundColor(color);
        return tr;
    }

    public WTable setBackgroundColor(String color){
        for(WTr tr:wtrs){
            tr.setBackgroundColor(color);
        }
        return this;
    }

    /**
     * 清除单元格样式
     * @param row 行
     * @param col 列
     * @return Wtc
     */
    public WTc removeStyle(int row, int col){
        return getTc(row, col).removeStyle();
    }
    /**
     * 清除整行单元格样式
     * @param row 行
     * @return Wtr
     */
    public WTr removeStyle(int row){
        WTr tr = tr(row);
        tr.removeContent();
        return tr;
    }
    public WTable removeStyle(){
        for(WTr tr:wtrs){
            tr.removeStyle();
        }
        return this;
    }
    /**
     * 清除单元格背景色
     * @param row 行
     * @param col 列
     * @return Wtc
     */
    public WTc removeBackgroundColor(int row, int col){
        return getTc(row, col).removeBackgroundColor();
    }

    /**
     * 清除整行单元格背景色
     * @param row 行
     * @return Wtr
     */
    public WTr removeBackgroundColor(int row){
        WTr tr = tr(row);
        tr.removeBackgroundColor();
        return tr;
    }
    public WTable removeBackgroundColor(){
        for(WTr tr:wtrs){
            tr.removeBackgroundColor();
        }
        return this;
    }

    /**
     * 清除单元格颜色
     * @param row 行
     * @param col 列
     * @return Wtc
     */
    public WTc removeColor(int row, int col){
        return getTc(row, col).removeColor();
    }
    /**
     * 清除整行单元格颜色
     * @param row 行
     * @return Wtr
     */
    public WTr removeColor(int row){
        WTr tr = tr(row);
        tr.removeColor();
        return tr;
    }
    public WTable removeColor(){
        for(WTr tr:wtrs){
            tr.removeColor();
        }
        return this;
    }
    /**
     * 粗体
     * @param row 行
     * @param col 列
     * @param bold 是否
     * @return Wtc
     */
    public WTc setBold(int row, int col, boolean bold){
        return getTc(row, col).setBold(bold);
    }
    public WTc setBold(int row, int col){
        return setBold(row, col, true);
    }
    public WTr setBold(int rows){
        return setBold(rows, true);
    }
    public WTr setBold(int rows, boolean bold){
        WTr tr = tr(rows);
        tr.setBold(bold);
        return tr;
    }
    public WTable setBold(boolean bold){
        for(WTr tr:wtrs){
            tr.setBold(bold);
        }
        return this;
    }
    public WTable setBold(){
        return setBold(true);
    }

    /**
     * 下划线
     * @param row 行
     * @param col 列
     * @param underline 是否
     * @return Wtc
     */
    public WTc setUnderline(int row, int col, boolean underline){
        return getTc(row, col).setUnderline(underline);
    }
    public WTc setUnderline(int row, int col){
        return setUnderline(row, col, true);
    }

    /**
     * 删除线
     * @param row 行
     * @param col 列
     * @param strike 是否
     * @return Wtc
     */
    public WTc setStrike(int row, int col, boolean strike){
        return getTc(row, col).setStrike(strike);
    }
    public WTc setStrike(int row, int col){
        return setStrike(row, col, true);
    }
    public WTr setStrike(int rows, boolean strike){
        WTr tr = tr(rows);
        tr.setStrike(strike);
        return tr;
    }
    public WTable setStrike(boolean strike){
        for(WTr tr:wtrs){
            tr.setStrike(strike);
        }
        return this;
    }
    public WTable setStrike(){
        return setStrike(true);
    }

    /**
     * 斜体
     * @param row 行
     * @param col 列
     * @param italic 是否
     * @return Wtc
     */
    public WTc setItalic(int row, int col, boolean italic){
        return getTc(row, col).setItalic(italic);
    }

    public WTc setItalic(int row, int col){
        return setItalic(row, col, true);
    }

    /**
     * 设置整行斜体
     * @param rows 行
     * @param italic 是否斜体
     * @return wtr
     */
    public WTr setItalic(int rows, boolean italic){
        WTr tr = tr(rows);
        tr.setItalic(italic);
        return tr;
    }
    public WTable setItalic(boolean italic){
        for(WTr tr:wtrs){
            tr.setItalic(italic);
        }
        return this;
    }
    public WTable setItalic(){
        return setItalic(true);
    }

    /**
     * 替换单元格内容
     * @param row 行
     * @param col 行
     * @param src src
     * @param tar tar
     * @return wtc
     */
    public WTc replace(int row, int col, String src, String tar){
        return getTc(row, col).replace(src, tar);
    }

    /**
     * 替换整行单元格内容<br/>
     * 如果有HTML转义符，需要通过WDocument对象的IS_HTML_ESCAPE属性设置是否解析文本转义符
     * @param rows 行(下标从0开始)
     * @param src src
     * @param tar tar
     * @return wtr
     */
    public WTr replace(int rows, String src, String tar){
        WTr tr = tr(rows);
        tr.replace(src, tar);
        return tr;
    }
    public WTable replace(String src, String tar){
        for(WTr tr:wtrs){
            tr.replace(src, tar);
        }
        return this;
    }


    public String getWidthUnit() {
        return widthUnit;
    }

    public void setWidthUnit(String widthUnit) {
        this.widthUnit = widthUnit;
        for(WTr tr:wtrs){
            tr.setWidthUnit(widthUnit);
        }
    }

    public boolean isAutoLoad() {
        return isAutoLoad;
    }

    public void setAutoLoad(boolean autoLoad) {
        isAutoLoad = autoLoad;
    }

    private final List<List<Spans>> spans_set = new ArrayList<>();
    public Spans spans(int row, int col){
        if(row < spans_set.size()){
            List<Spans> spans = spans_set.get(row);
            if(col < spans.size()){
                return spans.get(col);
            }
        }
        return null;
    }
    public int rowspan(int row, int col){
        Spans spans = spans(row, col);
        if(null != spans){
            return spans.getRowspan();
        }
        return -1;
    }
    public int colspan(int row, int col){
        Spans spans = spans(row, col);
        if(null != spans){
            return spans.getColspan();
        }
        return -1;
    }
    /**
     * 计算行列合并值
     */
    public void spans(){
        //1不合并 0被合并 >1合并其他单元格
        spans_set.clear();
        int rows = wtrs.size();
        for(int row=0; row<rows; row++){
            WTr tr = wtrs.get(row);
            List<Spans> spans_row = new ArrayList<>();
            spans_set.add(spans_row);
            List<WTc> tcs = tr.getTcs();
            int cols = tcs.size();
            for(int col=0; col<cols; col++){
                WTc tc = tcs.get(col);
                //合并列
                int colspan = tc.parseColspan();
                if(colspan == -1){
                    colspan = 1;
                }
                Spans spans = new Spans();
                spans_row.add(spans);
                spans.setColspan(colspan);
                if(colspan > 1){
                    //右则标签被删除，补齐
                    for(int i=1; i<colspan; i++){
                        Spans right = new Spans();
                        right.setColspan(0);
                        right.setRowspan(1);
                        spans_row.add(right);
                    }
                }
                //2:合并其他行 1:不合并 0:被合并
                int rowspan = tc.parseRowspan();
                if(rowspan == 0){
                    //被合并,上面单元格 +rowspan
                    for(int i=row-1; i>=0; i--){
                        Spans up = spans_set.get(i).get(col);
                        if(up.getRowspan() != 0){
                            up.addRowspan(1);
                            break;
                        }
                    }
                    spans.setRowspan(0);
                }else{
                    spans.setRowspan(1);
                }

            }
        }
    }
    public LinkedHashMap<String, String> styles(){
        LinkedHashMap<String, String> styles = new LinkedHashMap<>();
        return styles;
    }
    public String html(Uploader uploader, int lvl){
        spans();//计算colspan rowspan
        StringBuilder builder = new StringBuilder();
        StringBuilder body = new StringBuilder();
        int rows = 0;
        for(WTr tr:wtrs){
            body.append("\n");
            t(body, lvl+1);
            body.append("<tr");
            tr.styles(body);
            body.append(">\n");
            List<WTc> tcs = tr.getTcs();
            int cols = 0;
            for(WTc tc:tcs){
                Spans spans = spans(rows, cols);
                int rowspan = spans.getRowspan();
                int colspan = spans.getColspan();
                tc.setRowspan(rowspan);
                tc.setColspan(colspan);
                if(colspan > 0 && rowspan > 0) {
                    body.append(tc.html(uploader, lvl + 2));
                }
                if(colspan > 1){
                    cols += colspan-1;
                }
                cols ++;
            }
            body.append("\n");
            t(body, lvl+1);
            body.append("</tr>");
            rows ++;
        }
        t(builder, lvl);
        builder.append("<table");
        styles(builder);
        builder.append(">");
        builder.append(body);
        builder.append("\n");
        t(builder, lvl);
        builder.append("</table>\n");
        return builder.toString();
    }


    public WTable clone(boolean content){
        Element src = this.src.createCopy();
        WTable wtable = new WTable(root, src);
        if(!content){
            wtable.removeContent();
        }
        return wtable;
    }
}
