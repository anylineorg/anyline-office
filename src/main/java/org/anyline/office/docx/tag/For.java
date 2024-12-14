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

import org.anyline.adapter.KeyAdapter;
import org.anyline.entity.DataSet;
import org.anyline.office.docx.entity.Context;
import org.anyline.office.docx.entity.WTable;
import org.anyline.office.docx.entity.WTc;
import org.anyline.office.docx.entity.WTr;
import org.anyline.office.docx.util.DocxUtil;
import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;
import org.anyline.util.DomUtil;
import org.anyline.util.regular.RegularUtil;
import org.dom4j.Element;

import java.util.*;

public class For extends AbstractTag implements Tag {
    private Object items;
    private String var;
    private String status;
    private Integer begin;
    private Integer end;
    public void release(){
        super.release(); //不要清空context下个循环还要用
        /*if(null != var) {
            context.variables().remove(var);
        }
        if(null != status){
            context.variables().remove(status);
            }
        */
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
        //提取最外层标签属性 避免取到下一层属性
        String head = RegularUtil.fetchTagHead(text);
        items = fetchAttributeData("items", "is", "data", "d");
        String scope = fetchAttributeString(head, "scope", "sp");

        String body = RegularUtil.fetchTagBody(text, "aol:for");
        int type = 0; //0:txt 1:tc 2:tr
        int tr_index = -1; //模板行下标
        int tc_index = -1; //模板列下标
        List<Element> tcs = new ArrayList<>();
        List<WTc> wtcs = new ArrayList<>();
        List<Element> trs = new ArrayList<>();
        List<WTr> wtrs = new ArrayList<>();
        Element table = null;
        WTable wtable = null;
        if("tc".equalsIgnoreCase(scope) || "td".equalsIgnoreCase(scope)){
            type = 1;
            for(Element wt: wts){
                Element tc = DocxUtil.getParent(wt, "tc");
                if(!tcs.contains(tc)) {
                    tcs.add(tc);
                    if(tc_index == -1){
                        tc_index = DomUtil.elements(tc.getParent(), "tc").indexOf(tc);
                    }
                }
            }
        }else if("tr".equalsIgnoreCase(scope)){
            type = 2;
            for(Element wt: wts){
                Element tr = DocxUtil.getParent(wt, "tr");
                if(!trs.contains(tr)){
                    trs.add(tr);
                    if(tr_index == -1){
                        tr_index = DomUtil.elements(tr.getParent(), "tr").indexOf(tr);
                    }
                }
            }
        }else if("table".equalsIgnoreCase(scope)){
            type = 3;
            for(Element wt: wts){
                table = DocxUtil.getParent(wt, "tbl");
                if(null != table){
                    break;
                }
            }
        }
        boolean reload_table = false; //重新加载table(之前没有加载过会导致wtc获取不到)
        for(Element tc: tcs){
            WTc wtc = WTc.tc(tc);
            if(null == wtc && !reload_table){
                doc.tables();
                reload_table = true;
                wtc = WTc.tc(tc);
            }
            if(null != wtc) {
                wtcs.add(wtc);
            }
        }
        for(Element tr: trs){
            WTr wtr = WTr.tr(tr);
            if(null == wtr && !reload_table){
                doc.tables();
                reload_table = true;
                wtr = WTr.tr(tr);
            }
            if(null != wtr) {
                wtrs.add(wtr);
            }
        }
        if(null != table){
            wtable = WTable.table(table);
            if(null == wtable && !reload_table){
                doc.tables();
                reload_table = true;
                wtable = WTable.table(table);
            }
        }
        var = fetchAttributeString(head, "var");
        status = fetchAttributeString(head, "status", "s");
        begin = BasicUtil.parseInt(fetchAttributeString(head, "begin", "b"), 0);
        end = BasicUtil.parseInt(fetchAttributeString(head, "end", "e"), null);
        WTable prevTable = wtable;
        if(BasicUtil.isNotEmpty(items)) {//遍历集合
            if(items instanceof String){
                String str = (String) items;
                if(str.startsWith("[")){
                    if(str.startsWith("[{")){
                        items = DataSet.parseJson(KeyAdapter.KEY_CASE.SRC, str);
                    }
                }else{
                    items = BeanUtil.array2list(str.split(","));
                }
            }
            if (items instanceof Collection) {
                Collection list = (Collection) items;
                if(!list.isEmpty()){
                    int index = 0;
                    Context item_context = context.clone();
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
                        item_context.variable(var, item);
                        item_context.variable(status, map);
                        if(type == 1){
                            //遍历td
                            //在tr中添加td
                            tc(tc_index+index*wtcs.size(), wtcs, item_context);
                        } else if(type == 2){
                            //遍历tr
                            tr(tr_index+index*wtrs.size(), wtrs, item_context);
                        } else if(type == 3){
                            prevTable = table(prevTable, wtable, item_context);
                        } else if(null != body) {
                            //遍历文本
                            text(html, body, item_context);
                        }
                        index++;
                    }
                    //删除模板列、行
                    for(WTc tc: wtcs){
                        tc.remove();
                    }
                    for(WTr tr: wtrs){
                        tr.remove();
                    }
                }
            }
        }else{//按计数遍历
            if(null != end){
                Map<String, Object> map = new HashMap<>();
                int index = 0;
                Context item_context = context.clone();
                for(int i=begin; i<=end; i++){
                    map.put("index", index);
                    item_context.variable(var, i);
                    item_context.variable(status, map);
                    if(type == 1){
                        //遍历td
                        tc(tc_index++, wtcs, item_context);
                    } else if(type == 2){
                        //遍历tr
                        tr(tr_index++, wtrs, item_context);
                    } else if(type == 3){
                        prevTable = table(prevTable, wtable, item_context);
                    } else if(null != body) {
                        //遍历文本
                        text(html, body, item_context);
                    }
                    index++;
                }
            }
        }
        return html.toString();
    }
    //清除模板中的<aol:for <aol:a
    private void clear(WTr tr){

    }
    private void text(StringBuilder html, String body, Context context) throws Exception{
        String parse = DocxUtil.parseTag(doc, wts, body, context);
        parse = context.placeholder(parse);
        html.append(parse);
    }

    /**
     *
     * @param index 开始下标
     * @param templates
     * @throws Exception
     */
    private void tc(int index, List<WTc> templates, Context context) throws Exception{
        //遍历td
        //在tr中添加td
        WTr wtr = WTr.tr(templates.get(0).getSrc().getParent());
        int size = templates.size();
        int c = 0;
        for(WTc template:templates){
            String body = DocxUtil.text(template.getSrc());
            //TODO 注意<aol:a 格式
            //TODO 只清空当前层for
            if(c == 0 || c==size-1) {
                if (body.startsWith("<")) {
                    body = body.substring(body.indexOf(">") + 1);
                }
                body = body.replace("</aol:for>", "");
            }
            String parse = DocxUtil.parseTag(doc, wts, body, context);
            parse = context.placeholder(parse);
            wtr.insert(index + c, template,  parse);
            c++;
        }
    }
    private void tr(int index, List<WTr> templates, Context context) throws Exception{
        int size = templates.size();
        WTable wtable = WTable.table(templates.get(0).getSrc().getParent());
        int r = 0;
        for(WTr template:templates){
            WTr tr = template.clone(true);
            wtable.insert(index+r, tr);
            List<WTc> wtcs = tr.getTcs();
            for(WTc wtc:wtcs){
                Element csrc = wtc.getSrc();
                List<Element> cts = DomUtil.elements(csrc, "t");
                String txt = DocxUtil.text(cts);
                //TODO 注意<aol:a 格式
                //TODO 只清空当前层for
                //删除for本身内容
                if(r == 0 || r==size-1) {
                    String regex = "<aol:for.*/>";
                    txt = txt.replaceAll(regex, "");
                    regex = "<aol:for.*>";
                    txt = txt.replaceAll(regex, "");
                    txt = txt.replace("</aol:for>", "");
                }
                String parse = DocxUtil.parseTag(doc, cts, txt, context);
                parse = context.placeholder(parse);
                try {
                    wtc.setText("");//清空原内容
                    //Element box = DomUtil.element(wtc.getSrc(), "t").getParent();
                    doc.parseHtml(wtc.getSrc(), null, parse);
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
            r ++;
        }
    }
    private WTable table(WTable prev, WTable template, Context context) throws Exception{
        WTable insert = template.clone(true);
        DocxUtil.after(insert.getSrc(), prev.getSrc());
        return insert;
    }
}
