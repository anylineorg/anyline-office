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

package org.anyline.office.xlsx.entity;

import org.dom4j.Document;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

public class XSheet {
    private XWorkBook book;
    private Document doc;
    private Element root;
    private String name;
    private List<XRow> rows = new ArrayList<>();
    public XSheet(){}
    public XSheet(XWorkBook book, Document doc, String name){
        this.name = name;
        this.book = book;
        this.doc = doc;
        load();
    }
    public void load(){
        if(null == doc){
            return;
        }
        root = doc.getRootElement();
        Element data = root.element("sheetData");
        if(null == data){
            return;
        }
        List<Element> rows = data.elements("row");
        int index = 0;
        for(Element row:rows){
            XRow xr = new XRow(book, this, row, index++);
            this.rows.add(xr);
        }
    }

    public XWorkBook book(){
        return book;
    }
    public XSheet book(XWorkBook book){
        this.book = book;
        return this;
    }
    public Document doc(){
        return doc;
    }
    public XSheet doc(Document doc){
        this.doc = doc;
        return this;
    }
    public String name(){
        return name;
    }
    public XSheet doc(String name){
        this.name = name;
        return this;
    }
    public List<XRow> rows(){
        return rows;
    }
    /**
     * 解析标签
     * 注意有跨行的情况
     */
    public void parseTag(){
        //行解析跨行
        //再解析行内
        for(XRow row:rows){
            row.parseTag();
        }
    }
    public List<String> placeholders(String regex){
        List<String> placeholders = new ArrayList<>();
        for(XRow row:rows){
            placeholders.addAll(row.placeholders(regex));
        }
        return placeholders;
    }
    public List<String> placeholders(){
        return placeholders("\\$\\{.*?\\}");
    }
    /**
     * 插入行 注意插入行后 index 之后所有行与单元格需要重新计算r属性 如果插入量大 应该在插入完成后一次生调整
     * @param index 插入位置 下标从0开始 如果index<0 index=rows.size+index -1:表示最后一行
     * @param values 行内数据
     * @param template 模板行 如果null则以最后index上一行作模板(如果index是0则以index行作模板)
     * @return XRow
     */
    public XRow insert(int index, XRow template, List<Object> values){
        int size = values.size();
        if(size == 0){
            return null;
        }
        XRow row = XRow.build(index, book, this, template, values);
        return row;
    }

    /**
     * 插入行 以上一行为模板
     * @param index 插入位置 下标从0开始 如果index<0 index=rows.size+index -1:表示最后一行
     * @param values 行内数据
     * @return XRow
     */
    public XRow insert(int index, List<Object> values){
        return insert(index, null, values);
    }
    public XSheet insert(int index, XRow row){
        rows.add(index, row);
        Element datas = doc().getRootElement().element("sheetData");
        datas.elements().add(index, row.src);

        for(int i=index; i<rows.size(); i++){
            XRow item = rows.get(i);
            item.r(item.r()+1);
        }
        return this;
    }
    /**
     * 追加行
     * @param template 模板行 如果null则以最后一行作模板
     * @param values 行内数据
     * @return XRow
     */
    public XRow append(XRow template, List<Object> values){
        return insert(-1, template, values);
    }
    public XSheet append(XRow row){
        rows.add(row);
        Element datas = doc().getRootElement().element("sheetData");
        datas.elements().add(row.src);
        return this;
    }

    /**
     * 追加行，以最后一行作模板
     * @param values 行内数据
     * @return XRow
     */
    public XRow append(List<Object> values){
        XRow template = rows.get(rows.size()-1);
        return append(template, values);
    }
    public XSheet remove(XRow row){
        int index = rows.indexOf(row);
        if(index != -1){
            for(int i=index; i<rows.size(); i++){
                XRow item = rows.get(i);
                item.r(item.r()-1);
            }
        }
        rows.remove(row);
        Element datas = doc().getRootElement().element("sheetData");
        datas.elements().remove(row.src);
        return this;
    }
    public void replace(boolean parse, LinkedHashMap<String, String> replaces){
        for(XRow row:rows){
            row.replace(parse, replaces);
        }
    }

    /**
     * 如果删除过行 可以重新排序 避免r属性重复或不连续
     */
    public void sort(){
        int r = 1;
        for(XRow row:rows){
            row.r(r++);
        }
    }

}
