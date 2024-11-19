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

import org.anyline.util.BasicUtil;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

public class XRow extends XElement{
    private int index;  //下标 从0开始
    private int r;      //行号从1开始
    private String spans;
    private List<XCol> cols = new ArrayList<>();
    public XRow(XWorkBook book, XSheet sheet, Element src, int index){
        this.book = book;
        this.sheet = sheet;
        this.src = src;
        this.index = index;
        load();
    }
    public void load(){
        if(null == src){
            return;
        }
        //<row r="6" spans="1:6" ht="55.2" customHeight="1">
        this.r = BasicUtil.parseInt(src.attributeValue("r"), index+1);
        this.spans = src.attributeValue("spans");
        List<Element> cols = src.elements("c");
        int index = 0;
        for(Element col:cols){
            this.cols.add(new XCol(book, sheet, this, col, index));
        }
    }

    public List<String> placeholders(String regex){
        List<String> placeholders = new ArrayList<>();
        for(XCol col:cols){
            placeholders.addAll(col.placeholders(regex));
        }
        return placeholders;
    }
    public List<String> placeholders(){
        return placeholders("\\$\\{.*?\\}");
    }
    /**
     * 创建并插入行(index<0时 index = rows.size+index)
     * @param index 插入位置 下标从0开始 如果index<0 index=rows.size+index -1:表示最后一行
     * @param book XWorkBook
     * @param sheet XSheet
     * @param template 模板行 如果null则以最后一行作模板
     * @param values 行内数据
     * @return 插入的新行
     */
    public static XRow build(int index, XWorkBook book, XSheet sheet, XRow template, List<Object> values){
        if(null == values || values.isEmpty()){
            return null;
        }
        Element datas = sheet.doc().getRootElement().element("sheetData");
        Element row = datas.addElement("row");
        List<XRow> all = sheet.rows();
        boolean append = index == -1;
        if(!append) {
            index = BasicUtil.index(index, all.size());
        }
        if(null == template) {
            if(index == 0){
                template = all.get(index);
            }else {
                template = all.get(index - 1);
            }
        }

        //int x = index + 1;
        //if(append){
            //中间有可能有被删除的行 r比序号大的多
        //    x = all.size()+1;
       // }
        int x = template.r+1;
        int cols = values.size();
        String spans = "1:"+cols;
        row.addAttribute("r",x+"");
        row.addAttribute("spans", spans);
        XRow xr = new XRow(book, sheet, row, x);
        xr.spans = spans;
        xr.r = x;
        int y = 0;
        for(Object value:values){
            XCol tc = template.col(y);
            XCol col = XCol.build(book, sheet, xr, tc, value, x, y++);
            if(null != col){
                xr.add(col);
            }
            cols ++;
        }
        if(append){
            //最后一行 追加
            all.add(xr);
        } else {
            all.add(index, xr);
            datas.elements().remove(row);
            datas.elements().add(index,  row);
            //如果是插入的 插入行后 index 之后所有行与单元格需要重新计算r属性
            //如果插入量大 应该在插入完成后一次生调整

            for(int i=index; i<all.size(); i++){
                XRow item = all.get(i);
                item.r(item.r()+1);
            }
        }
        return xr;
    }
    public int r(){
        return r;
    }
    public XRow r(int r){
        this.r = r;
        src.attribute("r").setValue(r+"");
        for(XCol col:cols){
            col.x(r);
            col.r(col.y()+r);
        }
        return this;
    }
    public String spans(){
        return spans;
    }
    public XRow add(XCol col){
        cols.add(col);
        return this;
    }
    public XCol col(int index) {
        if(index < cols.size()){
            return cols.get(index);
        }
        return null;
    }
    public int index(){
        return index;
    }
    public XRow index(int index){
        this.index = index;
        return this;
    }
    /**
     * 解析标签
     * 注意有跨单元格的情况
     */
    public void parseTag(){
        //行解析跨单元格
        //再解析单元格内
        for(XCol col:cols){
            col.parseTag();
        }
    }
    public void replace(boolean parse, LinkedHashMap<String, String> replaces){
        for(XCol col:cols){
            col.replace(parse, replaces);
        }
    }

    /**
     * 复制一行
     * @param  r 同时修改行标识
     * @param content 是否复制其中内容
     * @return wtr
     */
    public XRow clone(int r, boolean content){
        XRow clone = new XRow(this.book, sheet, src.createCopy(), r-1);
        clone.r(r);
        return clone;
    }
}
