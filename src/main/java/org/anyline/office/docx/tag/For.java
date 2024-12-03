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

package org.anyline.office.docx.tag;

import org.anyline.office.docx.entity.Context;
import org.anyline.office.docx.entity.WTable;
import org.anyline.office.docx.entity.WTc;
import org.anyline.office.docx.entity.WTr;
import org.anyline.office.docx.util.DocxUtil;
import org.anyline.util.BasicUtil;
import org.anyline.util.DomUtil;
import org.anyline.util.regular.RegularUtil;
import org.dom4j.Element;

import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class For extends AbstractTag implements Tag {
    private Object items;
    private String var;
    private String status;
    private Integer begin;
    private Integer end;
    public void release(){
        super.release();
        items = null;
        var = null;
        status = null;
        begin = null;
        end = null;
    }

    /**
     * 解析标签
     * 务必注意:与普通标签不同的是，有可能需要控制的是外层tc,tr并且可能是连续的多个
     * 因为tc,tr的外层在word中接触不到所以当前标签只能写在td中
     * 通过scope属性指定 td或tc, tr,默认body即for标签体
     * @param text 原文
     * @return String
     * @throws Exception 异常
     */
    public String parse(String text) throws Exception {
        /*<aol:for
        data或items="${smaples}"
        item="samp"
        begin="0"
        end = "21"
        vol="3"
        direction="horizontal"
        scope="body"
        compensate="/,-"
        >${samp.CODE}</al:for>
        */
        StringBuilder html = new StringBuilder();
        String items_key = fetchAttributeValue(text, "items", "is");
        if(null == items_key){
            items_key = fetchAttributeValue(text, "data", "d");
        }
        if(null != items_key) {
            items = context.data(items_key);
        }
        String scope = fetchAttributeValue(text, "scope", "sp");

        String body = RegularUtil.fetchTagBody(text, "aol:for");

        Element tc = null;
        Element tr = null;
        Element table = null;
        if("tc".equalsIgnoreCase(scope) || "td".equalsIgnoreCase(scope)){
            tc = DocxUtil.getParent(wts.get(0), "tc");
            tr = tc.getParent();
            table = tr.getParent();
        }else if("tr".equalsIgnoreCase(scope)){
            tr = DocxUtil.getParent(wts.get(0), "tr");
            table = tr.getParent();
        }
        WTc wtc = WTc.tc(tc);
        WTr wtr = WTr.tr(tr);
        WTable wtable = WTable.table(table);
        boolean reload_table = false; //重新加载table(之前没有加载过会导致wtc获取不到)
        if(null != tc && null == wtc){
            reload_table = true;
        }
        if(null != tr && null == wtr){
            reload_table = true;
        }
        if(reload_table){
            doc.tables();
            wtc = WTc.tc(tc);
            wtr = WTr.tr(tr);
            wtable = WTable.table(table);
        }
        var = fetchAttributeValue(text, "var");
        status = fetchAttributeValue(text, "status", "s");
        begin = BasicUtil.parseInt(fetchAttributeValue(text, "begin", "b"), 0);
        end = BasicUtil.parseInt(fetchAttributeValue(text, "end", "e"), null);
        int row_index = -1; //模板行下标
        int col_index = -1; //模板列下标
        if(null != wtc){
            col_index = wtr.getTcs().indexOf(wtc);
        }else if(null != wtr){
            row_index = wtable.getTrs().indexOf(wtr);
            wtable.getSrc().remove(wtr.getSrc());
        }
        if(null != items) {//遍历集合
            if (items instanceof Collection) {
                Collection list = (Collection) items;
                int index = 0;
                Map<String, Object> map = new HashMap<>();
                for (Object item : list) {
                    if (null != begin && index < begin) {
                        index++;
                        continue;
                    }
                    if (null != end && index > end) {
                        break;
                    }
                    map.put("index", index);
                    context.variable(var, item);
                    context.variable(status, map);
                    if(null != wtc){
                        //遍历td
                        //在tr中添加td
                        tc(col_index++, index>0, wtr, wtc, body, context);
                    } else if(null != wtr){
                        //遍历tr
                        tr(row_index++, wtable, wtr, context);
                    } else if(null != body) {
                        //遍历文本
                        text(html, body);
                    }
                    index++;
                }
            }
        }else{//按计数遍历
            if(null != end){
                Map<String, Object> map = new HashMap<>();
                int index = 0;
                for(int i=begin; i<=end; i++){
                    map.put("index", index);
                    context.variable(var, i);
                    context.variable(status, map);
                    if(null != tc){
                        //遍历td
                        tc(col_index++, index>0, wtr, wtc, body, context);
                    } else if(null != tr){
                        //遍历tr
                        tr(row_index++,  wtable, wtr, context);
                    } else if(null != body) {
                        //遍历文本
                        text(html, body);
                    }
                    index++;
                }
            }
        }
        return html.toString();
    }
    private void text(StringBuilder html, String body) throws Exception{
        String parse = DocxUtil.parseTag(doc, wts, body, context);
        parse = context.placeholder(parse);
        html.append(parse);
    }

    /**
     *
     * @param index
     * @param insert 是否插入新行,第一个不插入
     * @param wtr
     * @param template
     * @param body
     * @throws Exception
     */
    private void tc(int index, boolean insert, WTr wtr, WTc template, String body, Context context) throws Exception{
        //遍历td
        //在tr中添加td
        String parse = DocxUtil.parseTag(doc, wts, body, context);
        parse = context.placeholder(parse);
        if(insert) {
            wtr.insert(index, parse);
        }else{
            template.setText(parse);
        }
    }
    private void tr(int index, WTable wtable, WTr template, Context context) throws Exception{
        WTr tr = template.clone(true);
        wtable.insert(index, tr);
        List<WTc> wtcs = tr.getTcs();
        for(WTc wtc:wtcs){
            Element csrc = wtc.getSrc();
            List<Element> cts = DomUtil.elements(csrc, "t");
            String txt = DocxUtil.text(cts);
            //删除for本身内容
            String regex = "<aol:for.*/>";
            txt = txt.replaceAll(regex, "");
            String parse = DocxUtil.parseTag(doc, cts, txt, context);
            parse = context.placeholder(parse);
            wtc.setText(parse);
        }
    }
}
