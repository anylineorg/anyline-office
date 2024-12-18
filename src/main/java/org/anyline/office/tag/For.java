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

package org.anyline.office.tag;

import org.anyline.adapter.KeyAdapter;
import org.anyline.entity.DataSet;
import org.anyline.office.docx.entity.WTable;
import org.anyline.office.docx.entity.WTc;
import org.anyline.office.docx.entity.WTr;
import org.anyline.office.docx.util.DocxUtil;
import org.anyline.office.util.Context;
import org.anyline.office.util.TagUtil;
import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;
import org.anyline.util.DomUtil;
import org.dom4j.Element;

import java.util.*;

public class For extends AbstractTag implements Tag {
    private Object items;
    private String var;
    private String status;
    private Integer begin;
    private Integer end;
    private Element prev;
    private Element head_top_copy = null;
    private int head_index = -1;
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
     * @throws Exception 异常
     */
    public void run() throws Exception {
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
        items = data();
        String scope = fetchAttributeString(head, "scope", "sp");

        int type = 0; //0:txt 1:tc 2:tr
        int tr_index = -1; //模板行下标
        int tc_index = -1; //模板列下标
        List<Element> tcs = new ArrayList<>();
        List<WTc> wtcs = new ArrayList<>();
        List<Element> trs = new ArrayList<>();
        List<WTr> wtrs = new ArrayList<>();

        //清空第一个t<for>和最后一个t(</for>) 继续下一层tag
        //先不要清空 for需要根据
       /* if(ts.size() > 1){
            DocxUtil.remove(ts.get(ts.size()-1));
            ts.remove(ts.size()-1);
        }
        DocxUtil.remove(ts.get(0));
        ts.remove(0);*/

        if("tc".equalsIgnoreCase(scope) || "td".equalsIgnoreCase(scope)){
            type = 1;
            for(Element wt: ts){
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
            for(Element wt: ts){
                Element tr = DocxUtil.getParent(wt, "tr");
                if(!trs.contains(tr)){
                    trs.add(tr);
                    if(tr_index == -1){
                        tr_index = DomUtil.elements(tr.getParent(), "tr").indexOf(tr);
                    }
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
        var = fetchAttributeString(head, "var");
        status = fetchAttributeString(head, "status", "s");
        begin = BasicUtil.parseInt(fetchAttributeString(head, "begin", "b"), 0);
        end = BasicUtil.parseInt(fetchAttributeString(head, "end", "e"), null);
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
                    int row = 0;
                    for (Object item : list) {
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
                        } else{
                            body(item_context, row);
                        }
                        row ++;
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
                    } else{
                        body(item_context, index);
                    }
                    index++;
                }
            }
        }
        TagUtil.clear(tops);
    }
    private void text(StringBuilder html, String body, Context context) throws Exception{
        String parse = TagUtil.parse(doc, ts, body, context);
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
            String parse = TagUtil.parse(doc, ts, body, context);
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
                String parse = TagUtil.parse(doc, cts, txt, context);
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
    private void body(Context context, int group) {
        int rows = tops.size();
        //复制tops
        log.warn("---------copy for body----------");
        //只复制head之后前foot之前内容，不包含head foot本身
      /*
        if(size == 1){
            List<Element> news = copyLine(tops.get(0), ts.get(ts.size()-1));
            appends.addAll(news);
        }else {
        }*/
        List<Element> news = new ArrayList<>();
        for (int i=0; i<rows; i++) {
            Element top = tops.get(i);
            log.warn("copy for item:{}", DocxUtil.text(top));
            if(i == 0){
                //保复制head之后内容
                if(null == head_top_copy){
                    head_top_copy = top.createCopy();
                    Element head = ts.get(0);
                    List<Element> contents = DocxUtil.contents(top);
                    head_index = contents.indexOf(head);
                }
                news.addAll(copyFirst());
            }else if(i == rows -1){
                news.addAll(copyLast(top));
            }else{
                news.addAll(copyInner(top));
                /*
                copy = top.createCopy();
                index++;
                items.add(index, copy);
                appends.add(copy);*/
            }
            log.warn("copy for item result:{}", DocxUtil.text(news));
        }
        try {
            TagUtil.parse(doc, news, context);
            doc.replace(news, context);
            log.warn("copy for body result:{}.{}", group, DocxUtil.text(news));
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    private List<Element> copyFirst(){
        List<Element> list = new ArrayList<>();
        //复制head及之后内容, 插入到foot之后
        List<Element> contents = DocxUtil.contents(head_top_copy);
         if(head_index == contents.size()-1){
            //head是最后一个 当前行没有其他内容 需要便利
            return list;
        }
         log.warn("copy first:{}", DocxUtil.text(head_top_copy));

        if(null == prev && head_index > 0){
            prev = DocxUtil.contents(tops.get(0)).get(head_index -1);
        }

        int size = contents.size();
        for(int i=head_index+1; i<size; i++){
            Element item = contents.get(i);
            Element copy = item.createCopy();
            log.warn("copy first item:{}", DocxUtil.text(copy));
            list.add(copy);
            if(null != prev){
                //如果有prev就插入到prev之后 如果prev是p 则插入
                if(prev.getName().equalsIgnoreCase("p")){
                    Element r = prev.addElement("w:r");
                    r.add(copy);
                }else {
                    DocxUtil.after(copy, prev);
                }
            }else{
                //如果没有prev就插入到head之前
                DocxUtil.before(copy, ts.get(0));
            }
            prev = copy;
        }
        prev = DocxUtil.getParent(prev, "p");
        return list;
    }
    private List<Element> copyLast(Element last){
        List<Element> list = new ArrayList<>();
        Element copy = last.createCopy();
        log.warn("copy last item:{}", DocxUtil.text(copy));
        //找到结束结束标签
        List<Element> contents = DocxUtil.contents(last);
        int idx = contents.indexOf(ts.get(ts.size()-1));
        contents = DocxUtil.contents(copy);
        int size = contents.size();
        for(int i=idx; i<size; i++){
            Element item = contents.get(i);
            DocxUtil.remove(item);
        }
        list.add(copy);
        DocxUtil.after(copy, prev);
        prev = copy;
        return list;
    }
    private List<Element> copyInner(Element inner){
        List<Element> list = new ArrayList<>();
        /*List<Element> contents = DocxUtil.contents(inner);
        for(Element item:contents){
            Element copy = item.createCopy();
            log.warn("copy inner item:{}", DocxUtil.text(copy));
            list.add(copy);
            DocxUtil.after(copy, prev);
            prev = copy;
        }*/
        Element copy = inner.createCopy();
        if(null != prev) {
            DocxUtil.after(copy, prev);
        }else{
            //如果没有prev就插入到head之前
            //第一个inner
            //first中没有需要复制的内容
            DocxUtil.before(copy, ts.get(0));
        }
        list.add(copy);
        prev = copy;
        return list;
    }

    /*
    private List<Element> copyLine(Element top, Element prev) {
        List<Element> list = new ArrayList<>();
        Element parent = prev.getParent();
        List<Element> elements = DocxUtil.contents(parent);
        boolean start = false;
        int index = elements.indexOf(prev);
        for(Element element:elements){
            if(element == ts.get(ts.size()-1)){
                break;
            }
            if(start){
                Element copy = element.createCopy();
                list.add(copy);
            }
            if(element == ts.get(0)){
                start = true;
            }
        }
        if(index == elements.size()-1){
            elements.addAll(list);
        }else {
            for (Element item : list) {
                elements.add(index++, item);
            }
        }
        return list;
    }*/
}
