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

package org.anyline.office.util;

import org.anyline.log.Log;
import org.anyline.log.LogProxy;
import org.anyline.office.docx.entity.WDocument;
import org.anyline.office.docx.entity.WElement;
import org.anyline.office.docx.util.DocxUtil;
import org.anyline.runtime.Context;
import org.anyline.office.tag.Tag;
import org.anyline.util.BasicUtil;
import org.anyline.util.DomUtil;
import org.anyline.util.regular.RegularUtil;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.List;

public class TagUtil {
    private static Log log = LogProxy.get(TagUtil.class);

    /**
     * 合并拆分到多个个t中标签，不限相同段落(p)<br/>
     * @param box 通常是body, p, table, tr, tc
     */
    public static void merge(WDocument doc, Element box){
        //全部t标签
        List<Element> ts = DomUtil.elements(true, box, "t");
        int size = ts.size();
        List<Element> items = new ArrayList<>();
        String full = "";
        List<Element> splits = new ArrayList<>();
        for(int i = 0; i < size; i++){
            Element t = ts.get(i);
            full += t.getText();
            if(full.contains("<")){
                items.add(t);
                if(!full.contains("<"+doc.namespace()+":") && !full.contains("</"+doc.namespace()+":")){
                    if(full.length() > 6){
                        //只有<但不是<"+doc.namespace()+":
                        full = "";
                        items.clear();
                    }
                    continue;
                }
                if(isClose(doc, full)){
                    splits.add(items.get(0));
                    //这里不需要是一个完整标签，是完整开头或完整结尾都可以
                    DocxUtil.mergeText(items);
                    //i += items.size() - 1;
                    full = "";
                    items.clear();
                }
            }else{
                full = "";
                items.clear();
            }
        }
        for(Element split:splits){
            split(split);
        }
    }

    /**
     * 拆分标签 head body foot 及前后缀拆到独立的t中
     * @param t wt
     */
    public static void split(Element t){
        String txt = t.getText();
        List<String> list = split(txt);
        int size = list.size();
        if(size > 1){
            Element ref = t;
            Element parent = t.getParent();
            for (int i=0; i<size; i++) {
                String item = list.get(i);
                if(i == 0){
                    t.setText(item);
                }else {
                    Element element = DocxUtil.addElement(parent, "t");
                    DocxUtil.after(element, ref);
                    element.setText(item);
                    ref = element;
                }
            }
        }
    }
    /**
     * 拆分标签 head body foot 及前后缀拆到独立的t中
     * @param text text
     */
    public static List<String> split(String text){
        List<String> list = new ArrayList<>();
        text = TagUtil.format(text);
        int fr = 0;
        while (true){
            if(text.isEmpty()){
                break;
            }
            int idx = text.indexOf("<", fr);
            if(idx == -1){
                list.add(text);
                break;
            }
            if(!text.startsWith("<")){
                //有前缀a<aol:   中a部分
                String prefix = text.substring(0, idx);
                if(BasicUtil.isFullString(prefix)){
                    list.add(prefix);
                    text = text.substring(idx);
                    fr = 0;
                }else{
                    fr = idx +1;
                }
            }else{
                //以<开头
                idx = text.indexOf(">", idx);
                if(idx == -1){
                    //没有结尾
                    list.add(text);
                    break;
                }
                String head = text.substring(0, idx+1);
                if(BasicUtil.isFullString(head)){
                    list.add(head);
                    text = text.substring(idx+1);
                    fr = 0;
                }else{
                    fr = idx +1;
                }
            }
        }
        return list;
    }
    public static void run(WDocument doc, Tag parent, Element box, Context context){
        List<Element> list = new ArrayList<>();
        list.add(box);
        run(doc, parent, list, context);
    }
    public static void run(WDocument doc, Tag parent, WElement box, Context context){
        List<Element> list = new ArrayList<>();
        list.add(box.getSrc());
        run(doc, parent, list, context);
    }
    public static void run(WDocument doc, Tag parent, List<Element> box, Context context){
        //全部t标签
        List<Element> contents = DocxUtil.contents(box);
        int size = contents.size();
        for(int i = 0; i < size; i++){
            Element content = contents.get(i);
            String txt = content.getText();
            if(txt.contains("<"+doc.namespace())){
                List<Element> items = new ArrayList<>(); //tag head body foot所在的t
                items.add(content);
                if(!RegularUtil.isFullTag(txt)){//如果不是完整标签(需要有开始和结束或自闭合)继续拼接下一个直到完成或失败
                    List<Element> nexts = next(doc, txt, contents, i+1);
                    if(!nexts.isEmpty()) {
                        txt = content.getText() + DocxUtil.text(nexts);
                        //removes.addAll(items);
                        Element last = nexts.get(nexts.size() - 1);
                        i = contents.indexOf(last);
                        items.addAll(nexts);
                    }else{
                        continue;
                    }
                }
                if(RegularUtil.isFullTag(txt)){
                    try {
                        run(doc, parent, items, txt, context);
                    }catch (Exception e){
                        log.error("解析异常:{}", txt);
                        e.printStackTrace();
                    }
                }
            }
        }
    }

    /**
     * 解析标签生成文本
     * @param doc WDocument
     * @param box box
     * @param context context
     * @return string
     */
    public static String parse(WDocument doc, List<Element> box, Context context){
        return parse(doc, DocxUtil.text(box), context);
    }
    public static String parse(WDocument doc, String text, Context context){
        StringBuilder builder = new StringBuilder();
        List<String> tags = RegularUtil.fetchOutTag(text, doc.namespace());
        List<String> splits = split(text, tags);
        for(String split:splits){
            if(split.startsWith("<"+doc.namespace()+":")){
                Tag instance = instance(doc, split);
                //复制占位值
                instance.init(doc);
                instance.text(split);
                instance.context(context);
                //把 "+doc.namespace()+"标签解析成html标签 下一步会解析html标签
                instance.prepare();
                try {
                    builder.append(instance.parse());
                }catch (Exception e){

                }
            }else{
                builder.append(split);
            }
        }
        return builder.toString();
    }

    /**
     * 把标签和标签中间的其他内容拆分开
     * a,tag,b,tag,c
     * @param text txt
     * @param tags tags
     * @return list
     */
    private static List<String> split(String text, List<String> tags){
        List<String> list = new ArrayList<>();
        for(String tag:tags){
            int idx = text.indexOf(tag);
            if(idx != 0){
                String prefix = text.substring(0, idx);
                list.add(prefix);
            }
            text = text.substring(idx + tag.length());
            list.add(tag);
        }
        list.add(text);
        return list;
    }


    /**
     * 开始和结束标签之间的tops
     * @param start start
     * @param end end
     * @return list
     */
    public static List<Element> tops(Element start, Element end){
        List<Element> tops = new ArrayList<>();
        Element p = DocxUtil.getParent(start, "p");
        if(null != p){
            Element pp = p.getParent();
            if(null != pp){
                List<Element> elements = pp.elements();
                boolean between = false;
                for(Element element:elements){
                    if(element == start || DomUtil.contains(element, start)){
                        between = true;
                    }
                    if(between){
                        tops.add(element);
                    }
                    if(element == end || DomUtil.contains(element, end)){
                        break;
                    }
                }
            }
        }
        return tops;
    }
    public static List<Element> tops(List<Element> elements){
        if(null != elements && !elements.isEmpty()){
            Element start = elements.get(0);
            Element end = start;
            if(elements.size() > 1){
                end = elements.get(elements.size()-1);
            }
            return tops(start, end);
        }
        return new ArrayList<>();
    }

    /**
     * 解析标签
     * @param doc doc
     * @param contents 标签所在的全部可见内容
     * @param txt 标签文本
     * @param context context
     * @throws Exception 解析异常
     */
    public static void run(WDocument doc, Tag parent, List<Element> contents, String txt, Context context) throws Exception{
        if(null == txt){
            return;
        }
        if(contents.isEmpty()){
            return;
        }
        //ts所在的及之间的所有p(tbl)
        //body部分不要根据t 因为空的p中没有t 要根据w:body.child
        txt = format(txt);
        // 不需要拆分了 split已经拆分完了
        // List<String> tags = RegularUtil.fetchOutTag(txt);
        //标签name如<"+doc.namespace()+":img 中的img
        String name = name(txt, doc.namespace()+":");

        boolean isPre = false;
        if("pre".equals(name)){
            //<"+doc.namespace()+":pre id="c"/>
            isPre = true;
        }else{
            //<"+doc.namespace()+":date pre="c"
            String preId = RegularUtil.fetchAttributeValue(txt, "pre");
            if(null != preId){
                isPre = true;
            }
        }
        if(!isPre) {
            Tag instance = instance(doc, txt);
            if (null != instance) {
                //复制占位值
                instance.init(doc);
                instance.text(txt);
                instance.contents(contents);
                instance.context(context);
                instance.parent(parent);
                //把 "+doc.namespace()+"标签解析成html标签 下一步会解析html标签
                instance.prepare();
                instance.run();
                //instance.release();
            }
            //txt = txt.replace(tag, html);
            //txt = BasicUtil.replaceFirst(txt, tag, html);
            //如果有子标签 应该在父标签中一块解析完
            /*if(txt.contains("<"+doc.namespace()+":")){
                txt = parseTag(txt, variables);
            }*/
        }
    }
    /**
     * 解析标签 生成文本
     * @param doc doc
     * @param contents 标签所在的全部可见内容
     * @param txt 标签文本
     * @param context context
     * @throws Exception 解析异常
     */
    public static String parse(WDocument doc, List<Element> contents, String txt, Context context) throws Exception{
        if(null == txt){
            return "";
        }
        if(contents.isEmpty()){
            return "";
        }
        //ts所在的及之间的所有p(tbl)
        //body部分不要根据t 因为空的p中没有t 要根据w:body.child
        txt = format(txt);
        // 不需要拆分了 split已经拆分完了
        // List<String> tags = RegularUtil.fetchOutTag(txt);
        //标签name如<"+doc.namespace()+":img 中的img
        String name = name(txt, doc.namespace()+":");

        boolean isPre = false;
        if("pre".equals(name)){
            //<"+doc.namespace()+":pre id="c"/>
            isPre = true;
        }else{
            //<"+doc.namespace()+":date pre="c"
            String preId = RegularUtil.fetchAttributeValue(txt, "pre");
            if(null != preId){
                isPre = true;
            }
        }
        if(!isPre) {
            Tag instance = instance(doc, txt);
            if (null != instance) {
                //复制占位值
                instance.init(doc);
                instance.text(txt);
                instance.contents(contents);
                instance.context(context);
                instance.prepare();
                return instance.parse();
            }
        }
        return "";
    }
/*    public static List<Element> tops(WDocument doc, List<Element> roots, List<Element> contents){
        List<Element> tops = new ArrayList<>();
        if(contents.isEmpty()){
            return tops;
        }
        List<Element> all = doc.getSrc().elements();
        Element content = contents.get(0);
        Element top = DocxUtil.getParent(roots, content, "tbl");
        if(null == top){
            top = DocxUtil.getParent(roots, content, "p");
        }
        int fr = all.indexOf(top);
        if(fr == -1){
            //有可能是copy出来的不在doc中
            for(Element item:contents){
                Element tp = DocxUtil.getParent(roots, item, "tbl");
                if(null == tp){
                    tp = DocxUtil.getParent(roots, item, "p");
                }
                if(null != tp && !tops.contains(tp)) {
                    tops.add(tp);
                }
            }
            return tops;
        }
        content = contents.get(contents.size()-1);
        top = DocxUtil.getParent(roots, content, "tbl");
        if(null == top){
            top = DocxUtil.getParent(roots, content, "p");
        }
        int to = all.indexOf(top);
        for(int i=fr; i<=to; i++){
            Element item = all.get(i);
            if(null != item && !tops.contains(item)) {
                tops.add(item);
            }
        }
        return tops;
    }*/

    public static Tag instance(WDocument doc, String tag){
        Tag instance = null;
        if(null == tag || !tag.contains("<"+doc.namespace()+":")){
            return null;
        }
        String name = name(tag, doc.namespace()+":");
        String parse = tag; //解析的标签体
        //先执行外层的 外层需要设置新变量值
        if (null == name) {
            log.error("未识别的标签格式:{}", tag);
        } else {
            //<"+doc.namespace()+":date format="" value=""/>
            instance = doc.tag(name);
        }
        String ref_text = null;
        String refId = RegularUtil.fetchAttributeValue(tag, "ref");
        if (null != refId) {
            ref_text = doc.ref(refId);
        }
        if(null == instance) {
            //<"+doc.namespace()+":c/>
            refId = name;
            String define = doc.ref(refId);
            ref_text = define;
            if (null != define) {
                //<"+doc.namespace()+":c/>
                //<"+doc.namespace()+":date ref="c" format="" value=""/>
                parse = define;
                name = name(parse, doc.namespace()+":");
                if (null == name) {
                    log.error("未识别的标签格式:{}", parse);
                } else {
                    instance = doc.tag(name);
                }
            }
        }
        if(null != instance) {
            instance.ref(ref_text);
        }
        return instance;
    }
    /**
     * 获取最外层tag所在的全部t
     * @param items 搜索范围
     * @param start 开始标记 &lt;或&lt;"+doc.namespace()+":
     * @param index 开始位置
     * @return ts
     */
    public static List<Element> next(WDocument doc, String start, List<Element> items, int index){
        List<Element> list = new ArrayList<>();
        int size = items.size();
        String flag = "<"+doc.namespace();
        String full = flag+RegularUtil.cut(start, flag, RegularUtil.TAG_END);
        for(int i=index; i<size; i++){
            Element item = items.get(i);
            list.add(item);
            String cur = item.getText();
            full += cur;
            if(BasicUtil.isEmpty(cur.trim())){
                //不影响完整性 不检测
                continue;
            }
            full = full.replace("\"", "'");
            String name = null;
            if(full.length() > 5){
                if(!full.trim().startsWith("<"+doc.namespace()+":")){
                    //不是标签
                    return new ArrayList<>();
                }
                name = RegularUtil.cut(full, doc.namespace()+":", " ");
                //有不带属性的标签
                String sname = RegularUtil.cut(full, doc.namespace()+":", ">");
                if(null != name && null != sname && sname.length() < name.length()){
                    name = sname;
                }
            }
            if(null != name){
                String head ="<"+doc.namespace()+":" + name;
                String foot_d = "</"+doc.namespace()+":"+name+">";
                String foot_s = "/>";
                int end_d = full.indexOf(foot_d);
                int end_s = full.indexOf(foot_s);
                String foot = foot_d;
                int end = end_d;
                if(end_s != -1){
                    //检测是否是单标签
                    String chk_s = full.substring(0, end_s);
                    if (!chk_s.contains(">")) {
                        //单标签结束
                        break;
                    }else{
                        //<"+doc.namespace()+":if test='${total>10}' var='if1'/>
                        //或者>在引号内
                        chk_s = chk_s.substring(0, chk_s.lastIndexOf(">"));
                        if(!BasicUtil.isFullString(chk_s)){
                            break;
                        }
                    }
                }
                if(end_d != -1){
                    int head_count = BasicUtil.charCount(full, head);
                    int foot_count = BasicUtil.charCount(full, foot);
                    if(foot_count == head_count){
                        //嵌套没有拆碎 否则说明缺少结束标签需要继续查找
                        break;
                    }
                }
            }
        }
        return list;
    }


    /**
     * 检测是否是开始或完整标签，主要检测有没有结尾&gt;
     * 如果没有标签 返回true
     * @param txt tag
     * @return boolean
     */
    public static boolean isClose(WDocument doc, String txt){
        String chk = format(txt).replace("\"", "'");
        chk = chk.replaceAll("'.*?'", "''");
        chk = chk.replaceAll("\".*?\"", "\"\"");
        //<"+doc.namespace()+":if></"+doc.namespace()+":if> <"+doc.namespace()+":number/>
        if(!chk.contains("<"+doc.namespace()+":") && !chk.contains("</"+doc.namespace()+":")){
            return true;
        }
        int idx = chk.lastIndexOf("<"+doc.namespace()+":");
        if(idx == -1){
            idx = chk.lastIndexOf("</"+doc.namespace()+":");
        }
        if(idx != -1){
            //"+doc.namespace()+":后部分
            chk = chk.substring(idx+5);

            //if test=”${xx> 100 && xx <10}
            //>不在引号内
            idx = chk.indexOf(">");
            if(idx != -1){
                chk = chk.substring(0, idx);
                if(BasicUtil.isFullString(chk)){
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 提取标签名称
     * @param text 文本
     * @param prefix 前缀 aol:
     * @return name
     */
    public static String name(String text, String prefix){
        String name = RegularUtil.cut(text, prefix, " ");
        //有不带属性的标签
        String sname = RegularUtil.cut(text, prefix, ">");
        if(null != name && null != sname && sname.length() < name.length()){
            name = sname;
        }
        if(null == name){
            name = sname;
        }
        if(null == name){
            name = RegularUtil.cut(text, prefix, "/");
        }
        return name;
    }

    /**
     * 标签文本格式化
     * @param text 文本
     * @return string
     */
    public static String format(String text){
        text = text.replace("“", "\"")
            .replace("”", "\"")
            .replace("’", "'")
            .replace("‘", "'");
        return text;
    }

    /**
     * 删除标签及标签体
     * 换成box.remove()
     * 第一个top保留标签之前内容
     * 最后个top保留标签之后内容
     * @param tops tops
     */
    public static void clear(WDocument doc, List<Element> tops){
        int size = tops.size();
        for(int i=0; i<size; i++){
            Element element = tops.get(i);
            if(i == 0){
                //删除head及之后内容
                boolean remove = false;
                List<Element> contents = DocxUtil.contents(element);
                for(Element content:contents){
                    String txt = content.getText();
                    if(txt.startsWith("<"+doc.namespace()+":")){
                        //log.warn("清空first:{}", txt);
                        DocxUtil.remove(content);
                        remove = true;
                        if(txt.endsWith("</"+doc.namespace()+":")){
                            remove = false;
                        }
                    }else if(txt.startsWith("</"+doc.namespace()+":")){
                        //log.warn("清空first:{}", txt);
                        DocxUtil.remove(content);
                        remove = false;
                    }else {
                        if (remove) {
                            //log.warn("清空first:{}", txt);
                            DocxUtil.remove(content);
                        }
                    }
                }
                //如果head之前中没有其他内容 删除整个结点(p)
                if(DocxUtil.isEmpty(element)){
                    DocxUtil.remove(element);
                }
            }else if(i == size-1){
                //删除foot及之前内容
                int remove = -1;
                List<Element> contents = DocxUtil.contents(element);
                int len = contents.size();
                for(int r = 0; r<len; r++){
                    Element t = contents.get(r);
                    String txt = t.getText();
                    if(txt.startsWith("</"+doc.namespace()+":")){
                        //找到最后一个</aot
                        remove = r;
                    }
                }
                for(int r = 0; r<=remove; r++){
                    //log.warn("清空last:{}", DocxUtil.text(ts.get(r)));
                    DocxUtil.remove(contents.get(r));
                }
                //如果foot之后中没有其他内容 删除整个结点(p)
                if(DocxUtil.isEmpty(element)){
                    DocxUtil.remove(element);
                }
            }else {
                Element top = tops.get(i);
                //log.warn("清空inner:{}", DocxUtil.text(top));
                DocxUtil.remove(top);
            }
        }
    }
}
