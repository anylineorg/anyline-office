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
import org.anyline.util.NumberUtil;
import org.anyline.util.regular.RegularUtil;
import org.dom4j.Element;

import java.util.*;

public class For extends AbstractTag implements Tag {
    private Object items;
    private String var;
    private String status;
    private Integer begin;
    private Integer end;
    private Integer step;
    private Integer fill; //填充空白行列(数据不足fill条时)
    private Integer vol;//遍历列时 每行最多几组
    private Boolean neat;
    private boolean remove; //没有数据时 是否删除模板
    public void prepare(){
        super.prepare();
        items = data(false);
        var = fetchAttributeString("var");
        status = fetchAttributeString("status", "s");
        begin = BasicUtil.parseInt(fetchAttributeString("begin", "start", "b"), 0);
        step = BasicUtil.parseInt(fetchAttributeString("step"), 1);
        end = BasicUtil.parseInt(fetchAttributeString("end", "e"), null);
        fill = BasicUtil.parseInt(fetchAttributeString("fill"), null);
        vol = BasicUtil.parseInt(fetchAttributeString("vol"), null);
        neat = BasicUtil.parseBoolean(fetchAttributeString("neat"), false);
        remove = BasicUtil.parseBoolean(fetchAttributeString("remove"), true);
    }
    public void release(){
        super.release();
        items = null;
        var = null;
        status = null;
        begin = null;
        end = null;
        step = null;
        fill = null;
        vol = null;
        neat = null;
    }

    /**
     * 生成纯文本
     * @return string
     */
    public String parse(){
        StringBuilder builder = new StringBuilder();
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
            if(null != items){
                if(!(items instanceof Collection)){
                    log.error("数据源仅支持集合形式:{}", items);
                    return "";
                }
            }
            if (items instanceof Collection) {
                Collection cols = (Collection) items;
                remove = remove && cols.isEmpty();
                if(!cols.isEmpty()){
                    Context item_context = context.clone();
                    Map<String, Object> map = new HashMap<>();
                    List<Object> list = new ArrayList<>(cols);
                    int size = list.size();
                    int count = 0;
                    if(null == end || end > size-1 || end < 0){
                        end = size-1;
                    }
                    if(null != fill){
                        end = NumberUtil.max(end, fill);
                    }
                    if(begin < 0){
                        begin = 0;
                    }
                    if(begin > size-1){
                        begin = size -1;
                    }
                    for (int i = begin; i <= end; i+= step) {
                        map.clear();
                        count ++;
                        Object item = null;
                        if(i < size) {
                            item = list.get(i);
                        }
                        if(i<size-1){
                            map.put("next", list.get(i+1));
                        }
                        if(i>0 && i<size){
                            map.put("prev", list.get(i-1));
                        }
                        map.put("index", i);
                        map.put("count", count);
                        item_context.variable(var, item);
                        item_context.variable(status, map);
                        builder.append(parse(item_context));
                    }
                }
            }
        }else{//按计数遍历
            if(null != end){
                Map<String, Object> map = new HashMap<>();
                int count = 0;
                Context item_context = context.clone();
                for(int i=begin; i<=end; i+=step){
                    map.clear();
                    count++;
                    map.put("index", i);
                    map.put("count", count);
                    if(i<end){
                        map.put("next", i+1);
                    }
                    if(i>0){
                        map.put("prev", i-1);
                    }
                    item_context.variable(var, i);
                    item_context.variable(status, map);
                    builder.append(parse(item_context));
                }
            }
        }
        if(null == parent || !(parent instanceof For)){
            //最外层执行完后清空定位
            last(null);
        }
        return builder.toString();
    }
    /**
     * 解析标签
     * 务必注意:与普通标签不同的是，有可能需要控制的是外层tc,tr并且可能是连续的多个
     * 因为tc,tr的外层在word中接触不到所以当前标签只能写在td中
     * 通过scope属性指定 td或tc, tr,默认body即for标签体
     * @throws Exception 异常
     */
    public void run() throws Exception {
        //提取最外层标签属性 避免取到下一层属性
        String scope = fetchAttributeString("scope", "sp");

        int type = 0; //0:body 1:tc 2:tr 3:table
        int tr_index = -1; //模板行下标
        int tc_index = -1; //模板列下标
        List<Element> tcs = new ArrayList<>();
        List<WTc> wtcs = new ArrayList<>();
        List<Element> trs = new ArrayList<>();
        List<WTr> wtrs = new ArrayList<>();
        WTable wtable =null;

        //清空第一个t<for>和最后一个t(</for>) 继续下一层tag
        //先不要清空 for需要根据

        if("tc".equalsIgnoreCase(scope) || "td".equalsIgnoreCase(scope)){
            type = 1;
            for(Element wt: contents){
                Element tc = DocxUtil.getParent(wt, "tc");
                if(!tcs.contains(tc)) {
                    tcs.add(tc);
                    if(tc_index == -1){
                        tc_index = DomUtil.elements(true, tc.getParent(), "tc").indexOf(tc);
                    }
                }
            }
        }else if("tr".equalsIgnoreCase(scope)){
            type = 2;
            for(Element wt: contents){
                Element tr = DocxUtil.getParent(wt, "tr");
                if(!trs.contains(tr)){
                    trs.add(tr);
                    if(tr_index == -1){
                        tr_index = DomUtil.elements(true, tr.getParent(), "tr").indexOf(tr);
                    }
                }
            }
        }else if("table".equalsIgnoreCase(scope)){
            type = 3;
            Element table = DocxUtil.getParent(contents.get(0), "tbl");
            wtable = new WTable(doc, table);
        }
        boolean reload_table = false; //重新加载table(之前没有加载过会导致wtc获取不到)
        for(Element tc: tcs){
            WTc wtc = WTc.tc(tc);
            if(null == wtc && !reload_table){
                doc.tables(true);
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
                doc.tables(true);
                reload_table = true;
                wtr = WTr.tr(tr);
            }
            if(null != wtr) {
                wtrs.add(wtr);
            }
        }

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
            if(null != items){
                if(!(items instanceof Collection)){
                    log.error("数据源仅支持集合形式:{}", items);
                    return;
                }
            }
            if (items instanceof Collection) {
                Collection cols = (Collection) items;
                if(!cols.isEmpty()){
                    Context item_context = context.clone();
                    Map<String, Object> map = new HashMap<>();
                    List<Object> list = new ArrayList<>(cols);
                    int size = list.size();
                    int count = 0;
                    if(null == end || end > size-1 || end < 0){
                        end = size-1;
                    }
                    if(null != fill){
                        end = NumberUtil.max(end, fill);
                    }
                    if(begin < 0){
                        begin = 0;
                    }
                    if(begin > size-1){
                        begin = size -1;
                    }
                    for (int i = begin; i <= end; i+= step) {
                        map.clear();
                        count ++;
                        Object item = null;
                        if(i < size) {
                            item = list.get(i);
                        }
                        if(i<size-1){
                            map.put("next", list.get(i+1));
                        }
                        if(i>0 && i<size){
                            map.put("prev", list.get(i-1));
                        }
                        map.put("index", i);
                        map.put("count", count);
                        item_context.variable(var, item);
                        item_context.variable(status, map);
                        if(type == 1){
                            //遍历td
                            //在tr中添加td
                            tc(tc_index+(count-1)*wtcs.size(), wtcs, item_context);
                        } else if(type == 2){
                            //遍历tr
                            tr(tr_index+(count-1)*wtrs.size(), wtrs, item_context);
                        } else if(type == 3){
                            table(wtable, item_context);
                        }else{
                            body(item_context);
                        }
                    }
                }
            }
        }else{//按计数遍历
            if(null != end){
                Map<String, Object> map = new HashMap<>();
                int count = 0;
                Context item_context = context.clone();
                for(int i=begin; i<=end; i+=step){
                    map.clear();
                    count++;
                    map.put("index", i);
                    map.put("count", count);
                    if(i<end){
                        map.put("next", i+1);
                    }
                    if(i>0){
                        map.put("prev", i-1);
                    }
                    item_context.variable(var, i);
                    item_context.variable(status, map);
                    if(type == 1){
                        //遍历td
                        tc(tc_index+(count-1)*wtcs.size(), wtcs, item_context);
                    } else if(type == 2) {
                        //遍历tr
                        tr(tr_index+(count-1)*wtrs.size(), wtrs, item_context);
                    } else if(type == 3) {
                        table(wtable, item_context);
                    } else {
                        body(item_context);
                    }
                }
            }
        }
         if(remove) {
             box.remove();
             //删除模板列、行、表
             for (WTc tc : wtcs) {
                 tc.remove();
             }
             for (WTr tr : wtrs) {
                 tr.remove();
             }
             if (null != wtable) {
                 wtable.remove();
             }
         }
         if(null == parent || !(parent instanceof For)){
             //最外层执行完后清空定位
            last(null);
         }
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
        int c = 0;
        List<Element> news = new ArrayList<>();
        for(WTc template:templates){
            WTc copy = template.clone(true);
            wtr.insert(index+c, copy);
            news.add(copy.getSrc());
            c++;
        }
        TagUtil.run(doc, this, news, context);
        doc.replace(news, context);
    }
    private void tr(int index, List<WTr> templates, Context context) throws Exception{
        WTable wtable = WTable.table(templates.get(0).getSrc().getParent());
        int r = 0;
        List<Element> news = new ArrayList<>();
        for(WTr template:templates){
            WTr copy = template.clone(true);
            wtable.insert(index+r, copy);
            news.add(copy.getSrc());
            r ++;
        }
        TagUtil.run(doc, this, news, context);
        doc.replace(news, context);
    }
    private void table(WTable template, Context context){
        if(null == last()){
            last(box.tops().get(0));
        }
        WTable copy = template.clone(true);
        DocxUtil.after(copy.getSrc(), last());
        TagUtil.run(doc, this, copy, context);
        doc.replace(copy, context);
        last(copy.getSrc());
    }
    private String parse(Context context) {
        StringBuilder builder = new StringBuilder();
        try {
            String body = RegularUtil.fetchTagBody(text, doc.namespace()+":for");
            body = TagUtil.parse(doc, body, context);
            body = context.placeholder(body);
            builder.append(body);
        }catch (Exception e){
            log.error("[text 解析异常]\n[template:{}]", text);
            e.printStackTrace();
        }
        return builder.toString();
    }
    private void body(Context context) {
        List<Element> templates = box.templates();
        int rows = templates.size();
        //复制templates
        //只复制head之后前foot之前内容，不包含head foot本身
        List<Element> news = new ArrayList<>();
        for (int i=0; i<rows; i++) {
            Element template = templates.get(i);
            if(i == 0){
                news.addAll(copyFirst(template));
            }else if(i == rows -1){
                news.addAll(copyLast(template));
            }else{
                news.addAll(copyInner(template));
            }
        }
        try {
            TagUtil.run(doc, this, news, context);
            doc.replace(news, context);
        }catch (Exception e){
            log.error("[body 解析异常]\n[template:{}]", DocxUtil.text(templates));
            e.printStackTrace();
        }
    }
    private List<Element> copyFirst(Element template){
        List<Element> list = new ArrayList<>();
        int head_index = box.head().index();
        int foot_index = box.foot().index();
        //复制head及之后内容, 插入到foot之后
        List<Element> contents = DocxUtil.contents(template);
         if(box.head().last()){
            //head是最后一个 当前行没有其他内容 需要便利
             //下一行不要插入到foot之后，因为foot.p可能有其他标签外内容
             if(null == last()) {
                 last(box.tops().get(0));
             }
            return list;
        }
         //log.warn("copy first:{}", DocxUtil.text(template));

        if(null == last() && head_index > 0){
            last(DocxUtil.contents(box.tops().get(0)).get(head_index -1));
        }

        int end = contents.size();
        for(int i=0; i<end; i++){
            Element item = contents.get(i);
            Element copy = item.createCopy();
            //log.warn("copy first item:{}", DocxUtil.text(copy));
            list.add(copy);
            if(null != last()){
                //如果有prev就插入到prev之后 如果prev是p 则插入
                if(last().getName().equalsIgnoreCase("p")){
                    Element r = last().addElement("w:r");
                    r.add(copy);
                }else {
                    DocxUtil.after(copy, last());
                }
            }else{
                //如果没有prev就插入到head之前
                DocxUtil.before(copy, this.contents.get(0));
            }
            last(copy);
        }
        if(null != last() && tops().size() > 1) {//如果模板有多行，下一行插入到p之后(换行)
            if (!last().getName().equalsIgnoreCase("p")) {
                //如果最后一个 不是p 应该把最上层的p作为最后一个，下次从p后面开始插入
                last(DocxUtil.getParent(last(), "p"));
            }
        }
        return list;
    }
    private List<Element> copyLast(Element last){
        List<Element> list = new ArrayList<>();
        if(box.foot().index() == 0){
            //foot在开头  没有内容需要复制
            return list;
        }
        Element copy = last.createCopy();
        //找到结束结束标签 删除结束标签及之后的内容
        List<Element> contents = DocxUtil.contents(copy);
        int size = contents.size();
        for(int i=box.foot().index(); i<size; i++){
            Element item = contents.get(i);
            DocxUtil.remove(item);
        }
        list.add(copy);
        DocxUtil.after(copy, last());
        last(copy);
        return list;
    }
    private List<Element> copyInner(Element inner){
        List<Element> list = new ArrayList<>();
        Element copy = inner.createCopy();
        if(null != last()) {
            DocxUtil.after(copy, last());
        }else{
            //如果没有prev就插入到head之前
            //第一个inner
            //first中没有需要复制的内容
            DocxUtil.before(copy, contents.get(0));
        }
        list.add(copy);
        last(copy);
        return list;
    }
}
