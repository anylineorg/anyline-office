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




package org.anyline.office.docx.util;

import org.anyline.log.Log;
import org.anyline.log.LogProxy;
import org.anyline.office.docx.entity.Context;
import org.anyline.office.docx.entity.WDocument;
import org.anyline.office.docx.tag.Tag;
import org.anyline.util.BasicUtil;
import org.anyline.util.DomUtil;
import org.anyline.util.StyleParser;
import org.anyline.util.ZipUtil;
import org.anyline.util.regular.RegularUtil;
import org.dom4j.*;

import java.io.File;
import java.util.*;

public class DocxUtil {
    private static Log log = LogProxy.get(DocxUtil.class);
    /**
     * 根据关键字查找样式列表ID
     * @param docx docx文件
     * @param key 关键字
     * @return String
     */
    public static String listStyle(File docx, String key, String charset){
        try {
            String num_xml = ZipUtil.read(docx, "word/document.xml", charset);
            Document document = DocumentHelper.parseText(num_xml);
            List<Element> ts = DomUtil.elements(document.getRootElement(),"t");
            for(Element t:ts){
                if(t.getTextTrim().contains(key)){
                    Element pr = t.getParent().getParent().element("pPr");
                    if(null != pr) {
                        Element numPr = pr.element("numPr");
                        if(null != numPr){
                            Element numId = numPr.element("numId");
                            if(null != numId){
                                String val = numId.attributeValue("val");
                                if(null != val){
                                    return val;
                                }
                            }
                        }
                    }
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 根据关键字查找样式列表ID
     * @param docx docx文件
     * @return String
     */
    public static List<String> listStyles(File docx, String charset){
        List<String> list = new ArrayList<>();
        try {
            String num_xml = ZipUtil.read(docx, "word/numbering.xml", charset);
            Document document = DocumentHelper.parseText(num_xml);
            List<Element> nums = document.getRootElement().elements("num");
            for(Element num:nums){
                list.add(num.attributeValue("numId"));
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return list;
    }

    /**
     * 合并文件(只合并内容document.xml)合并到第一个文件中
     * @param files files
     */
    public static void merge(String charset, File ... files){
        if(null != files && files.length>1){
            List<String> docs = new ArrayList<>();
            for(File file:files){
                docs.add(ZipUtil.read(file,"word/document.xml", charset));
            }
            String result = merge(docs);
            try {
                Document document = DocumentHelper.parseText(result);
                Element root = document.getRootElement().element("body");
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }
    public static String merge(List<String>  docs){
        String result = null;
        return result;
    }

    /**
     * copy的样式复制给src
     * @param src src
     * @param copy 被复制p/w或pPr/wPr
     * @param override 如果样式重复,是否覆盖原来的样式
     */
    public static void copyStyle(Element src, Element copy, boolean override){
        if(null == src || null == copy){
            return;
        }
        String name = src.getName();
        String prName = name+"Pr";
        Element srcPr = src.element(prName);
        if(override){
            src.remove(srcPr);
            srcPr = null;
        }
        Element pr = null;
        String copyName = copy.getName();
        if(copyName.equals(prName)){
            pr = copy;
        }else {
            pr = DomUtil.element(copy, prName);;
        }
        if(null != pr){
            if(null == srcPr) {
                // 如果原来没有pr新创建一个
                Element newPr = pr.createCopy();
                src.elements().add(0, newPr);
            }else{
                List<Element> items = pr.elements();
                List<Element> newItems = new ArrayList<>();
                for(Element item:items){
                    String itemName = item.getName();
                    Element srcItem = srcPr.element(itemName);
                    if(override){
                        srcPr.remove(srcItem);
                        srcItem = null;
                    }
                    if(null == srcItem){
                        // 如果原来没有这个样式条目直接复制一个
                        Element newItem = item.createCopy();
                        newItems.add(newItem);
                    }else{
                        // 如果原来有这个样式条目,在原来基础上复制属性
                        List<Attribute> attributes = item.attributes();
                        for(Attribute attribute:attributes){
                            String attributeName = attribute.getName();
                            String attributeFullName = attributeName;
                            String attributeNamespace = attribute.getNamespacePrefix();
                            if(BasicUtil.isNotEmpty(attributeNamespace)){
                                attributeFullName = attributeNamespace+":"+attributeName;
                            }
                            Attribute srcAttribute = srcItem.attribute(attributeName);
                            if(null == srcAttribute){
                                srcAttribute = srcItem.attribute(attributeFullName);
                            }
                            if(override){
                                if(null != srcAttribute){
                                    srcItem.remove(srcAttribute);
                                    srcAttribute = null;
                                }
                            }
                            if(null == srcAttribute) {
                                srcItem.attributeValue(attributeFullName, attribute.getStringValue());
                            }
                        }
                    }
                }
                srcPr.elements().addAll(newItems);
            }
        }
    }
    public static void copyStyle(Element src, Element copy){
        copyStyle(src, copy, false);
    }
    /**
     * 前一个节点
     * @param element element
     * @return element
     */
    public static Element prevByName(Element element){
        Element prev = null;
        List<Element> elements = DomUtil.elements(top(element), element.getName());
        int index = elements.indexOf(element);
        if(index > 0){
            prev = elements.get(index -1);
        }
        return prev;
    }
    public static Element prevByName(Element parent, Element element){
        Element prev = null;
        List<Element> elements = DomUtil.elements(parent, element.getName());
        int index = elements.indexOf(element);
        if(index > 0){
            prev = elements.get(index -1);
        }
        return prev;
    }
    public static Element nextByName(Element element){
        Element prev = null;
        List<Element> elements = DomUtil.elements(top(element), element.getName());
        int index = elements.indexOf(element);
        if(index < elements.size()-1 && index > 0){
            prev = elements.get(index + 1);
        }
        return prev;
    }
    public static Element nextByName(Element parent, Element element){
        Element prev = null;
        List<Element> elements = DomUtil.elements(parent, element.getName());
        int index = elements.indexOf(element);
        if(index < elements.size()-1 && index > 0){
            prev = elements.get(index + 1);
        }
        return prev;
    }
    public static Element top(Element element){
        Element top = element.getParent();
        while (null != top.getParent()){
            top = top.getParent();
        }
        return top;
    }
    /**
     * 前一个节点
     * @param element element
     * @return element
     */
    public static Element prev(Element element){
        Element prev = null;
        List<Element> elements = element.getParent().elements();
        int index = elements.indexOf(element);
        if(index>0){
            prev = elements.get(index-1);
        }
        return prev;
    }
    public static String prevName(Element element){
        Element prev = prev(element);
        if(null != prev){
            return prev.getName();
        }else{
            return "";
        }
    }
    public static Element last(Element element){
        Element last = null;
        List<Element> elements = element.getParent().elements();
        if(elements.size()>0){
            last = elements.get(elements.size()-1);
        }
        return last;
    }
    public static String lastName(Element element){
        Element last = last(element);
        if(null != last){
            return last.getName();
        }else{
            return "";
        }
    }

    /**
     * 是否有内容(表格、文本、图片)
     * @param element element
     * @return boolean
     */
    public static boolean isEmpty(Element element){
        List<Element> elements = DomUtil.elements(element, "drawing,tbl,t");
        for(Element item:elements){
            String name = item.getName();
            if(name.equalsIgnoreCase("drawing")){
                return false;
            }

            if(name.equalsIgnoreCase("tbl")){
                return false;
            }
            if(name.equalsIgnoreCase("t")){
                if(item.getTextTrim().length() > 0){
                    return false;
                }
            }
        }
        if(element.getTextTrim().length() > 0){
            return false;
        }
        return true;
    }
    private static boolean isEmpty(List<Element> elements){
        for(Element item:elements){
            String name = item.getName();
            if(name.equalsIgnoreCase("r") || name.equalsIgnoreCase("t") || name.equalsIgnoreCase("tbl")){
                return false;
            }
        }
        return true;
    }

    public static boolean hasParent(Element element, String parent){
        Element p = element.getParent();
        while(true){
            if(null == p){
                break;
            }
            if(p.getName().equalsIgnoreCase(parent)) {
                return true;
            }
            p = p.getParent();
        }
        return false;
    }

    /**
     * 获取element的上一级中第一个标签名=tag的上级
     * @param element 当前节点
     * @param tag 上级标签名 如tbl
     * @return Element
     */
    public static Element getParent(Element element, String tag){
        Element p = element.getParent();
        if(null == tag){
            return p;
        }
        while(true){
            if(null == p){
                break;
            }
            if(p.getName().equalsIgnoreCase(tag)) {
                return p;
            }
            p = p.getParent();
        }
        return null;
    }

    /**
     * wt列表文本合并
     * @param ts W:t 集合
     * @return String
     */
    public static String text(List<Element> ts){
        StringBuilder builder = new StringBuilder();
        for(Element t:ts){
            String txt = t.getText();
            if(null != txt) {
                builder.append(txt);
            }
        }
        return builder.toString();
    }
    /**
     * src插入到ref之后
     * @param src src
     * @param ref ref
     */
    public static void after(Element src, Element ref){
        if(null == ref || null == src){
            return;
        }
        if(src == ref){
            //ref取父标签后可能与src一样
            return;
        }
        Element rp = ref.getParent();
        Element sp = src.getParent();
        // 同级
        if(rp == sp || null == sp){
            List<Element> elements = ref.getParent().elements();
            int index = elements.indexOf(ref)+1;
            elements.remove(src);
            if(index > elements.size()-1){
                elements.add(src);
            }else {
                elements.add(index, src);
            }
        } else {
            // ref更下级
            after(src, ref.getParent());
        }

    }
    public static void after(List<Element> srcs, Element ref){
        if(null == ref || null == srcs){
            return;
        }
        int size = srcs.size();
        for(int i=size-1; i>=0; i--){
            Element src = srcs.get(i);
            // after(src, ref);
        }
        for(Element src:srcs){
            after(src, ref);
        }

    }
    /**
     * src插入到ref之前
     * @param src src
     * @param ref ref
     */
    public static void before(Element src, Element ref){
        if(null == ref || null == src){
            return;
        }
        List<Element> elements = ref.getParent().elements();
        int index = elements.indexOf(ref);
        while (!elements.contains(src)){
            src = src.getParent();
            if(null == src){
                return;
            }
        }
        elements.remove(src);
        elements.add(index, src);

    }
    /**
     * 当前节点在上级节点的下标
     * @param element element
     * @return index index
     */
    public static int index(Element element){
        int index = -1;
        List<Element> elements = element.getParent().elements();
        index = elements.indexOf(element);
        return index;
    }

    public static List<String> splitKey(String txt){
        return splitKey(txt, "\\$\\{.*?\\}");
    }
    /**
     * 拆分关键字
     * 拆分123${key}abc成多个w:t   abc,${key},123 最后执行解析时再反转顺序
     * @param txt txt
     * @param regex 正则
     * @return List
     */
    public static List<String> splitKey(String txt, String regex){
        List<String> list = new ArrayList<>();
        try {
            List<String> keys = RegularUtil.fetch(txt, regex);
            int size = keys.size();
            if(size>0){
                String key = keys.get(keys.size()-1);
                int index = txt.lastIndexOf(key);
                String t1 = txt.substring(0, index);
                String t2 = txt.substring(index + key.length());
                if (t2.length() > 0) {
                    list.addAll(splitKey(t2));
                }
                list.add(key);
                if (t1.length() > 0) {
                    list.addAll(splitKey(t1));
                }
                //txt = txt.substring(0, txt.length() - key.length());
            }else{
                list.add(txt);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }

    public static String tagName(String text, String prefix){
        String name = RegularUtil.cut(text, prefix, " ");;
        if(null == name){
            name = RegularUtil.cut(text, prefix, "/");
        }
        return name;
    }
    public static void parseTag(WDocument doc, Element box, Context context){
        //全部t标签
        List<Element> ts = DomUtil.elements(box, "t");
        int size = ts.size();
        List<Element> removes = new ArrayList<>();
        for(int i = 0; i < size; i++){
            Element t = ts.get(i);
            String txt = t.getText();
            if(txt.contains("<")){
                List<Element> items = new ArrayList<>(); //tag标签头 标签体 标签尾所在的t
                items.add(t);
                if(!RegularUtil.isFullTag(txt)){//如果不是完整标签(需要有开始和结束或自闭合)继续拼接下一个直到完成或失败
                    items = tagNext(txt, ts, i+1);
                    if(!items.isEmpty()) {
                        txt = t.getText() + DocxUtil.text(items);
                        removes.addAll(items);
                        Element last = items.get(items.size() - 1);
                        i = ts.indexOf(last);
                    }else{
                        continue;
                    }
                }
                try {
                    txt = parseTag(doc, items, txt, context);
                    t.setText(txt);
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        }
        remove(removes);
    }
    public static String tagFormat(String txt){
        txt = txt.replace("“", "'").replace("”", "'").replace("’", "'").replace("‘", "'");
        return txt;
    }
    /**
     * 解析标签
     * @param doc doc
     * @param ts 标签所在的全部t
     * @param txt 标签文本
     * @param context context
     * @return String
     * @throws Exception 解析异常
     */
    public static String parseTag(WDocument doc, List<Element> ts, String txt, Context context) throws Exception{
        if(null == txt){
            return "";
        }
        if(ts.isEmpty()){
            return "";
        }
        List<Element> tops = new ArrayList<>();//ts所在的最上级table或p
        for(Element t:ts){
            Element top = DocxUtil.getParent(t, "tbl");
            if(null == top){
                top = DocxUtil.getParent(t, "p");
            }
            if(null != top && !tops.contains(top)){
                tops.add(top);
            }
        }
        if(tops.size() == 1){
            return parseSimpleTag(doc, ts, txt, context);
        }
        txt = tagFormat(txt);
        /*
         * 这里 不要把内层标签单独拆出来，因为外层标签可能 会设置新变量值影响内层
         */
        List<String> tags = RegularUtil.fetchOutTag(txt);
        for(String tag:tags){
            //标签name如<aol:img 中的img
            String name = tagName(tag, "aol:");
            String head = RegularUtil.fetchTagHead(tag);
            String foot = "</aol:"+name+">";
            if(head.endsWith("/>")){
                foot = null;
            }
            //找到起止top
            Element first_top = tops.get(0);
            Element last_top = tops.get(tops.size()-1);

            String first_top_text = text(first_top);
            String last_top_text = text(last_top);
            //定位需要标签体的tops
            int body_top_index_fr = 0;
            int body_top_index_to = tops.size()-1;
            if(first_top_text.trim().endsWith(head)){
                //以标签头结尾 标签体在下一个top
                body_top_index_fr = 1;
            }
            if(null != foot && last_top_text.trim().startsWith(foot)){
                body_top_index_to = body_top_index_to -1 ;
            }
            List<Element> body_tops = new ArrayList<>();
            for(int i=body_top_index_fr; i <= body_top_index_to; i++){
                body_tops.add(tops.get(i));
            }

            boolean isPre = false;
            if("pre".equals(name)){
                //<aol:pre id="c"/>
                isPre = true;
            }else{
                //<aol:date pre="c"
                String preId = RegularUtil.fetchAttributeValue(tag, "pre");
                if(null != preId){
                    isPre = true;
                }
            }
            String html = "";
            if(!isPre) {
                Tag instance = instance(doc, tag);
                if (null == instance) {
                    log.error("未识别的标签名称:{}", name);
                }
                if (null != instance) {
                    //复制占位值
                    instance.init(doc);
                    instance.wts(ts);
                    instance.tops(body_tops);
                    instance.context(context);
                    //把 aol标签解析成html标签 下一步会解析html标签
                    html = instance.parse(tag);
                    //instance.release();
                }
            }
            //txt = txt.replace(tag, html);
            txt = BasicUtil.replaceFirst(txt, tag, html);
            //如果有子标签 应该在父标签中一块解析完
            /*if(txt.contains("<aol:")){
                txt = parseTag(txt, variables);
            }*/
        }
        return txt;
    }
    public static Tag instance(WDocument doc, String tag){
        Tag instance = null;
        String name = tagName(tag, "aol:");
        String parse = tag; //解析的标签体
        //先执行外层的 外层需要设置新变量值
        if (null == name) {
            log.error("未识别的标签格式:{}", tag);
        } else {
            //<aol:date format="" value=""/>
            instance = doc.tag(name);
        }
        String ref_text = null;
        String refId = RegularUtil.fetchAttributeValue(tag, "ref");
        if (null != refId) {
            ref_text = doc.ref(refId);
        }
        if(null == instance) {
            //<aol:c/>
            refId = name;
            String define = doc.ref(refId);
            ref_text = define;
            if (null != define) {
                //<aol:c/>
                //<aol:date ref="c" format="" value=""/>
                parse = define;
                name = tagName(parse, "aol:");
                if (null == name) {
                    log.error("未识别的标签格式:{}", parse);
                } else {
                    instance = doc.tag(name);
                }
            }
            instance.ref(ref_text);
        }
        return instance;
    }
    public static String parseSimpleTag(WDocument doc, List<Element> ts, String txt, Context context) throws Exception{
        if(null == txt){
            return "";
        }
        if(ts.isEmpty()){
            return "";
        }
        txt = tagFormat(txt);
        //String reg = "(?i)(<aol:(\\w+)[^<]*?>)[^<]*(</aol:\\2>)";
        //这里 不要把内层标签独立拆出来，因为外层标签可能 会设置新变量值影响内层
        List<String> tags = RegularUtil.fetchOutTag(txt);
        for(String tag:tags){
            //标签name如<aol:img 中的img
            String name = tagName(tag, "aol:");
            boolean isPre = false;
            if("pre".equals(name)){
                //<aol:pre id="c"/>
                isPre = true;
            }else{
                //<aol:date pre="c"
                String preId = RegularUtil.fetchAttributeValue(tag, "pre");
                if(null != preId){
                    isPre = true;
                }
            }
            String html = "";
            if(!isPre) {
                //不是预定义
                Tag instance = instance(doc, tag);
                if (null == instance) {
                    log.error("未识别的标签名称:{}", name);
                }
                if (null != instance) {
                    //复制占位值
                    instance.init(doc);
                    instance.wts(ts);
                    instance.context(context);
                    //把 aol标签解析成html标签 下一步会解析html标签
                    html = instance.parse(tag);
                    //instance.release();
                }
            }
            //txt = txt.replace(tag, html);
            txt = BasicUtil.replaceFirst(txt, tag, html);
            //如果有子标签 应该在父标签中一块解析完
            /*if(txt.contains("<aol:")){
                txt = parseTag(txt, variables);
            }*/
        }
        return txt;
    }
    /**
     * 获取最外层tag所在的全部t
     * @param items 搜索范围
     * @param start 开始标记 &lt;或&lt;aol:
     * @param index 开始位置
     * @return ts
     */
    public static List<Element> tagNext(String start, List<Element> items, int index){
        List<Element> list = new ArrayList<>();
        int size = items.size();
        String full = "<"+RegularUtil.cut(start, "<", RegularUtil.TAG_END);
        for(int i=index; i<size; i++){
            Element item = items.get(i);
            list.add(item);
            String cur = item.getText();
            full += cur;
            if(BasicUtil.isEmpty(cur.trim())){
                continue;
            }
            full = full.replace("\"", "'");
            String name = null;
            if(full.length() > 5){
                if(!full.trim().startsWith("<aol:")){
                    //不是标签
                    return new ArrayList<>();
                }
                name = RegularUtil.cut(full, "aol:", " ");
            }
            if(null != name){
                String head ="<aol:" + name;
                String foot_d = "</aol:"+name+">";
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
                        //<aol:if test='${total>10}' var='if1'/>
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
    public static void border(Element border, Map<String, String> styles){
        border(border,"top", styles);
        border(border,"right", styles);
        border(border,"bottom", styles);
        border(border,"left", styles);
        border(border,"insideH", styles);
        border(border,"insideV", styles);
        border(border,"tl2br", styles);
        border(border,"tr2bl", styles);
    }
    public static void border(Element border, String side, Map<String, String> styles){
        Element item = null;
        String width = styles.get("border-"+side+"-width");
        String style = styles.get("border-"+side+"-style");
        String color = styles.get("border-"+side+"-color");
        int dxa = DocxUtil.dxa(width);
        int line = ((int)(DocxUtil.dxa2pt(dxa)*8)/4*4);
        if(BasicUtil.isNotEmpty(width)){
            item = element(border, side);
            item.addAttribute("w:sz", line+"");
            item.addAttribute("w:val", style);
            item.addAttribute("w:color", color);
        }
    }
    public static void padding(Element margin, Map<String, String> styles){
        padding(margin,"top", styles);
        padding(margin,"start", styles);
        padding(margin,"bottom", styles);
        padding(margin,"end", styles);

    }
    public static void padding(Element margin, String side, Map<String, String> styles){
        String width = styles.get("padding-"+side);
        int dxa = DocxUtil.dxa(width);
        if(BasicUtil.isNotEmpty(width)){
            Element item = element(margin, side);
            item.addAttribute("w:w", dxa+"");
            item.addAttribute("w:type",  "dxa");
        }
    }
    public static int fontSize(String size){
        int pt = 0;
        if(fontSizes.containsKey(size)){
            pt = fontSizes.get(size);
        }else{
            if(size.endsWith("px")){
                int px = BasicUtil.parseInt(size.replace("px",""),0);
                pt = (int)DocxUtil.px2pt(px);
            }else if(size.endsWith("pt")){
                pt = BasicUtil.parseInt(size.replace("pt",""),0);
            }
        }
        return pt;
    }
    public static void font(Element pr, Map<String, String> styles){
        String fontSize = styles.get("font-size");
        if(null != fontSize){
            int pt = 0;
            if(fontSizes.containsKey(fontSize)){
                pt = fontSizes.get(fontSize);
            }else{
                if(fontSize.endsWith("px")){
                    int px = BasicUtil.parseInt(fontSize.replace("px",""),0);
                    pt = (int)DocxUtil.px2pt(px);
                }else if(fontSize.endsWith("pt")){
                    pt = BasicUtil.parseInt(fontSize.replace("pt",""),0);
                }
            }
            if(pt>0){
                // <w:sz w:val="28"/>
                element(pr, "sz","val", pt+"");
            }
        }
        // 加粗
        String fontWeight = styles.get("font-weight");
        if(null != fontWeight && fontWeight.length()>0){
            int weight = BasicUtil.parseInt(fontWeight,0);
            if(weight >=700){
                // <w:b w:val="true"/>
                element(pr, "b","val","true");
            }
        }
        // 下划线
        String underline = styles.get("underline");
        if(null != underline){
            if(underline.equalsIgnoreCase("true") || underline.equalsIgnoreCase("single")){
                // <w:u w:val="single"/>
                element(pr, "u","val","single");
            }else{
                element(pr, "u","val",underline);
                /*dash - a dashed line
                dashDotDotHeavy - a series of thick dash, dot, dot characters
                dashDotHeavy - a series of thick dash, dot characters
                dashedHeavy - a series of thick dashes
                dashLong - a series of long dashed characters
                dashLongHeavy - a series of thick, long, dashed characters
                dotDash - a series of dash, dot characters
                dotDotDash - a series of dash, dot, dot characters
                dotted - a series of dot characters
                dottedHeavy - a series of thick dot characters
                double - two lines
                none - no underline
                single - a single line
                thick - a single think line
                wave - a single wavy line
                wavyDouble - a pair of wavy lines
                wavyHeavy - a single thick wavy line
                words - a single line beneath all non-space characters
                */
            }
        }
        // 删除线
        String strike = styles.get("strike");
        if(null != strike){
            if(strike.equalsIgnoreCase("true")){
                // <w:dstrike w:val="true"/>
                element(pr, "dstrike","val","true");
            }else if("none".equalsIgnoreCase(strike) || "false".equalsIgnoreCase(strike)){
                element(pr, "dstrike","val","false");
            }
        }
        // 斜体
        String italics = styles.get("italic");
        if(null != italics){
            if(italics.equalsIgnoreCase("true")){
                // <w:dstrike w:val="true"/>
                element(pr, "i","val","true");
            }else if("none".equalsIgnoreCase(italics) || "false".equalsIgnoreCase(italics)){
                element(pr, "i","val","false");
            }
        }
        String fontFamily = styles.get("font-family");
        if(null != fontFamily){
            element(pr, "rFonts","eastAsia",fontFamily);
        }
        String fontFamilyAscii = styles.get("font-family-ascii");
        if(null != fontFamilyAscii){
            element(pr, "rFonts","ascii",fontFamilyAscii);
        }
        String fontFamilyEast = styles.get("font-family-east");
        if(null != fontFamilyEast){
            element(pr, "rFonts","eastAsia",fontFamilyEast);
        }
        fontFamilyEast = styles.get("font-family-eastAsia");
        if(null != fontFamilyEast){
            element(pr, "rFonts","eastAsia",fontFamilyEast);
        }
        String fontFamilyhAnsi = styles.get("font-family-height");
        if(null != fontFamilyhAnsi){
            element(pr, "rFonts","hAnsi",fontFamilyhAnsi);
        }
        fontFamilyhAnsi = styles.get("font-family-hAnsi");
        if(null != fontFamilyhAnsi){
            element(pr, "rFonts","hAnsi",fontFamilyhAnsi);
        }
        String fontFamilyComplex = styles.get("font-family-complex");
        if(null != fontFamilyComplex){
            element(pr, "rFonts","cs",fontFamilyComplex);
        }
        fontFamilyComplex = styles.get("font-family-cs");
        if(null != fontFamilyComplex){
            element(pr, "rFonts","cs",fontFamilyComplex);
        }

        String fontFamilyHint = styles.get("font-family-hint");
        if(null != fontFamilyHint){
            element(pr, "rFonts","hint",fontFamilyHint);
        }
        // <w:rFonts w:ascii="Adobe Gothic Std B" w:eastAsia="宋体" w:hAnsi="宋体" w:cs="宋体" w:hint="eastAsia"/>
    }

    public static void background(Element pr,Map<String, String> styles){
        String color = styles.get("background-color");
        if(null != color){
            // <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>
            DocxUtil.element(pr, "shd", "color","auto");
            DocxUtil.element(pr, "shd", "val","clear");
            DocxUtil.element(pr, "shd", "fill",color.replace("#",""));
        }
    }

    /**
     * 添加element及属性
     * @param parent parent
     * @param tag element tag
     * @param key attribute key
     * @param value attribute value
     * @return Element
     */
    public static Element element(Element parent, String tag, String key, String value){
        Element element = DocxUtil.element(parent,tag);
        addAttribute(element, key, value);
        return element;
    }

    /**
     * 添加属性值，如果属性已存在 先删除原属性
     * @param element Element
     * @param key 属性key
     * @param value 属性值
     */
    public static void addAttribute(Element element, String key, String value){
        Attribute attribute = element.attribute(key);
        if(null == attribute){
            attribute = element.attribute("w:"+key);
        }
        if(null != attribute){
            element.remove(attribute);
        }
        element.addAttribute("w:"+key, value);
    }

    /**
     * 添加节点，如果已经有了 就用原来的
     * @param parent 上级
     * @param tag name
     * @return element
     */
    public static Element element(Element parent, String tag){
        Element element = parent.element(tag);
        if(null == element){
            element = parent.addElement("w:"+tag);
        }
        return element;
    }
    public static Element addElement(Element parent, String tag){
        Element element =  parent.addElement("w:"+tag);
        return element;
    }

    public static Element next(Element parent, Element child){
        Element next = null;
        while(child.getParent() != parent){
            child = child.getParent();
            if(null == child){
                break;
            }
        }
        if(null != child){
            List<Element> elements = parent.elements();
            int index = elements.indexOf(child);
            if(index != -1){
                index ++;
                if(index >0 && index <elements.size()-1){
                    next = elements.get(index);
                }
            }
        }
        return next;
    }
    public static Element prev(Element parent, Element child){
        Element next = null;
        while(child.getParent() != parent){
            child = child.getParent();
            if(null == child){
                break;
            }
        }
        if(null != child){
            List<Element> elements = parent.elements();
            int index = elements.indexOf(child);
            if(index != -1){
                index --;
                if(index >0 && index <elements.size()-1){
                    next = elements.get(index);
                }
            }
        }
        return next;
    }
    /**
     * 当前节点下的文本
     * @param element element
     * @return String
     */
    public static String text(Element element){
        String text = "";
        Iterator<Node> nodes = element.nodeIterator();
        while (nodes.hasNext()) {
            Node node = nodes.next();
            int type = node.getNodeType();
            if(type == 3){
                text += node.getText();
            }else{
                text += text((Element)node);
            }
        }
        return text.trim();
    }
    public static boolean isBlock(String text){
        if(null != text){
            List<String> styles = RegularUtil.cuts(text,true,"<style",">","</style>");
            for(String style:styles){
                text = text.replace(style,"");
            }
            text = text.trim();
            if(text.startsWith("<div") || text.startsWith("<ul") || text.startsWith("<ol") || text.startsWith("<table")){
                return true;
            }
        }
        return false;
    }


    public static List<Element> betweens(Element bookmark, String ... tags){
        String id = bookmark.attributeValue("id");
        Element end = null;
        List<Element> ends = bookmark.getParent().elements("bookmarkEnd");
        for(Element item:ends){
            if(id.equals(item.attributeValue("id"))){
                end = item;
                break;
            }
        }
        return DomUtil.betweens(bookmark, end, tags);
    }
    public static Element bookmark(Element parent, String name){
        Element start = DomUtil.element(parent, "bookmarkStart", "name", name);
        return start;
    }

    /**
     * 合并占位符 包含ognl
     * @param box 通常是一个p标签
     */
    public static void mergePlaceholder(Element box){
        List<Element> items = DomUtil.elements(box, "t,br,bookmarkStart");
        //<w:bookmarkStart w:id="0" w:name="a"/>
        int size = items.size();
        String full = "";
        List<Element> merges = new ArrayList<>();
        for(int i=0; i<size; i++){
            Element item = items.get(i);
            String name = item.getName();
            if(name.equals("br") || name.equals("bookmarkStart")){
                merges.clear();
                continue;
            }
            String txt = item.getText();
            full += txt;
            merges.add(item);
            //a${b, c,}
            //a$, {, b, c, }

            if(!full.contains("$")){
                //没有占位符 重新计数
                full = "";
                merges.clear();
                continue;
            }
            if(full.endsWith("$")){
                continue;
            }
            String after_char = after(true, full, "$");
            if(!"{".equals(after_char)){
                merges.clear();
                full = "";
                continue;
            }
            //占位符开始位置
            int head_qty = BasicUtil.charCount(full, "$");
            int start_qty = BasicUtil.charCount(full, "{");
            int end_qty = BasicUtil.charCount(full, "}");
            if(head_qty == start_qty && head_qty == end_qty){
                //完整 占位符
                //开始合并
                full = "";
                mergeText(merges);
                //i+= merges.size()-1;
                merges.clear();
            }
        }
    }

    /**
     * 合并拆分到多个个t中标签，不限相同段落(p)<br/>
     * @param box 通常是body, p, table, tr, tc
     */
    public static void mergeTag(Element box){
        //全部t标签
        List<Element> ts = DomUtil.elements(box, "t");
        int size = ts.size();
        List<Element> items = new ArrayList<>();
        String full = "";
        List<Element> splits = new ArrayList<>();
        for(int i = 0; i < size; i++){
            Element t = ts.get(i);
            full += t.getText();
            if(full.contains("<")){
                items.add(t);
                if(checkTagClose(full)){
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
            splitTag(split);
        }
    }

    /**
     * 拆分标签 head body foot 及前后缀拆到独立的t中
     * @param t wt
     */
    public static void splitTag(Element t){
        String txt = t.getText();
        List<String> list = splitTag(txt);
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
    public static List<String> splitTag(String text){
        List<String> list = new ArrayList<>();
        text = tagFormat(text);
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
                //有前缀
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

    /**
     * 检测是否是开始或完整标签，主要检测有没有结尾&gt;
     * 如果没有标签 返回true
     * @param txt tag
     * @return boolean
     */
    public static boolean checkTagClose(String txt){
        String chk = tagFormat(txt).replace("\"", "'");
        chk = chk.replaceAll("'.*?'", "''");
        //<aol:if></aol:if> <aol:number/>
        if(!chk.contains("<aol:") && !chk.contains("</aol:")){
            return true;
        }
        int idx = chk.lastIndexOf("<aol:");
        if(idx == -1){
            idx = chk.lastIndexOf("</aol:");
        }
        if(idx != -1){
            //aol:后部分
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

    public static void mergeText(List<Element> ts){
        int size = ts.size();
        if(size > 1){
            String text = DocxUtil.text(ts);
            text = tagFormat(text);
            Element first = ts.get(0);
            first.setText(text);
            for(int i=1; i<size; i++){
                Element t = ts.get(i);
                remove(t);
            }
        }
    }
    public static void remove(List<Element> elements){
        int size = elements.size();
        for(int i=0; i<size; i++){
            Element element = elements.get(i);
            remove(element);
        }
    }
    public static void remove(Element element){
        Element parent = element.getParent();
        if(null != parent){
            parent.remove(element);
            String pn = parent.getName();
            if("r".equalsIgnoreCase(pn)){
                if(parent.elements("t").isEmpty() && parent.elements("br").isEmpty()){
                    parent.getParent().remove(parent);
                }
            }
        }else{
            log.error("重复删除:{}", element.getText());
        }
    }

    /**
     * flag后一个字符
     * @param empty 是否包含空
     * @param text 全文
     * @param flag 开始位置
     * @return char
     */
    public static String after(boolean empty, String text, String flag){
        String after = null;
        int idx = text.lastIndexOf(flag);
        if(idx != -1){
            int length = text.length();
            while (true) {
                if (idx + flag.length() < length) {
                    after = text.substring(idx + flag.length(), idx + flag.length() + 1);
                    if(!" ".equalsIgnoreCase(after) || empty){
                        break;
                    }
                }else {
                    break;
                }
                idx ++;
            }
        }
        return after;
    }

    public static Element pr(Element element, String styles){
        return pr(element, StyleParser.parse(styles));
    }
    public static Element pr(Element element, Map<String, String> styles){
        if(null == styles){
            styles = new HashMap<String, String>();
        }
        String name = element.getName();
        String prName = name+"Pr";
        Element pr = DocxUtil.element(element, prName);
        //pr需要放在第一个位置 否则样式对后面的内容可能无效
        List<Element> elements = element.elements();
        if(elements.size() > 1){
            elements.remove(pr);
            elements.add(0, pr);
        }
        if("p".equalsIgnoreCase(name)){
            for(String sk: styles.keySet()){
                String sv = styles.get(sk);
                if(BasicUtil.isEmpty(sv)){
                    continue;
                }
                if(sk.equalsIgnoreCase("list-style-type")){
                    DocxUtil.element(pr, "pStyle", "val",sv);
                }else if(sk.equalsIgnoreCase("list-lvl")){
                    Element numPr = DocxUtil.element(pr,"numPr");
                    DocxUtil.element(numPr, "ilvl", "val",sv+"");
                }else if(sk.equalsIgnoreCase("numFmt")){
                    Element numPr = DocxUtil.element(pr,"numPr");
                    DocxUtil.element(numPr, "numFmt", "val",sv+"");
                }else if ("text-align".equalsIgnoreCase(sk)) {
                    DocxUtil.element(pr, "jc","val", sv);
                }else if(sk.equalsIgnoreCase("margin-left")){
                    DocxUtil.element(pr, "ind", "left",DocxUtil.dxa(sv)+"");
                }else if(sk.equalsIgnoreCase("margin-right")){
                    DocxUtil.element(pr, "ind", "right",DocxUtil.dxa(sv)+"");
                }else if(sk.equalsIgnoreCase("margin-top")){
                    DocxUtil.element(pr, "spacing", "before",DocxUtil.dxa(sv)+"");
                }else if(sk.equalsIgnoreCase("margin-bottom")){
                    DocxUtil.element(pr, "spacing", "after",DocxUtil.dxa(sv)+"");
                }else if(sk.equalsIgnoreCase("padding-left")){
                    DocxUtil.element(pr, "ind", "left",DocxUtil.dxa(sv)+"");
                }else if(sk.equalsIgnoreCase("padding-right")){
                    DocxUtil.element(pr, "ind", "right",DocxUtil.dxa(sv)+"");
                }else if(sk.equalsIgnoreCase("padding-top")){
                    DocxUtil.element(pr, "spacing", "before",DocxUtil.dxa(sv)+"");
                }else if(sk.equalsIgnoreCase("padding-bottom")){
                    DocxUtil.element(pr, "spacing", "after",DocxUtil.dxa(sv)+"");
                }else if(sk.equalsIgnoreCase("text-indent")){
                    DocxUtil.element(pr, "ind", "firstLine",DocxUtil.dxa(sv)+"");
                }else if(sk.equalsIgnoreCase("line-height")){
                    DocxUtil.element(pr, "spacing", "line",DocxUtil.dxa(sv)+"");
                }
            }
            if(styles.containsKey("list-style-num")){
                // 如果在样式里指定了样式
                Element numPr = DocxUtil.element(pr,"numPr");
                DocxUtil.element(numPr, "numId", "val",styles.get("list-style-num"));
            }else if(styles.containsKey("list-num")){
                // 运行时自动生成
                Element numPr = DocxUtil.element(pr,"numPr");
                DocxUtil.element(numPr, "numId", "val",styles.get("list-num"));
            }

            // <div style="page-size-orient:landscape"/>
            if(styles.containsKey("page-size-orient")){
                String orient = styles.get("page-size-orient");
                if(!"landscape".equalsIgnoreCase(orient)){
                    orient = "portrait";
                }
                setOrient(pr, orient, styles);
            }

            Element border = DocxUtil.element(pr, "bdr");
            DocxUtil.border(border, styles);
            // DocxUtil.background(pr, styles);

        }else if("r".equalsIgnoreCase(name)){
            for (String sk : styles.keySet()) {
                String sv = styles.get(sk);
                if(BasicUtil.isEmpty(sv)){
                    continue;
                }
                if(sk.equalsIgnoreCase("color")){
                    element(pr, "color", "val", sv.replace("#",""));
                }else if(sk.equalsIgnoreCase("background-color")){
                    // <w:highlight w:val="yellow"/>
                    DocxUtil.element(pr, "highlight", "val",sv.replace("#",""));
                }else if(sk.equalsIgnoreCase("vertical-align")){
                    DocxUtil.element(pr,"vertAlign", "val", sv );
                }
            }
            Element border = DocxUtil.element(pr, "bdr");
            DocxUtil.border(border, styles);
            DocxUtil.font(pr, styles);
        }else if("tbl".equalsIgnoreCase(name)){

            // DocxUtil.element(pr,"tblCellSpacing","w","0");
            // DocxUtil.element(pr,"tblCellSpacing","type","nil");

            Element mar = DocxUtil.element(pr,"tblCellMar");
            /*DocxUtil.element(mar,"top","w","0");
            DocxUtil.element(mar,"top","type","dxa");
            DocxUtil.element(mar,"bottom","w","0");
            DocxUtil.element(mar,"bottom","type","dxa");
            DocxUtil.element(mar,"right","w","0"); // 新版本end
            DocxUtil.element(mar,"right","type","dxa");
            DocxUtil.element(mar,"end","w","0");
            DocxUtil.element(mar,"end","type","dxa");
            DocxUtil.element(mar,"left","w","0");//新版本用start,但07版本用start会报错
            DocxUtil.element(mar,"left","type","dxa");*/
            for (String sk : styles.keySet()) {
                String sv = styles.get(sk);
                if(BasicUtil.isEmpty(sv)){
                    continue;
                }
                if(sk.equalsIgnoreCase("width")){
                    DocxUtil.element(pr,"tblW","w", DocxUtil.dxa(sv)+"");
                    DocxUtil.element(pr,"tblW","type", DocxUtil.widthType(sv));
                }else if(sk.equalsIgnoreCase("color")){
                }else if(sk.equalsIgnoreCase("margin-left")){
                    DocxUtil.element(pr,"tblInd","w",DocxUtil.dxa(sv)+"");
                    DocxUtil.element(pr,"tblInd","type","dxa");
                }else if(sk.equalsIgnoreCase("padding-left")){
                    DocxUtil.element(mar,"left","w",DocxUtil.dxa(sv)+""); // 新版本用start,但07版本用start会报错
                    DocxUtil.element(mar,"left","type","dxa");
                }else if(sk.equalsIgnoreCase("padding-right")){
                    DocxUtil.element(mar,"right","w",DocxUtil.dxa(sv)+""); // 新版本用end
                    DocxUtil.element(mar,"right","type","dxa");
                    DocxUtil.element(mar,"end","w",DocxUtil.dxa(sv)+"");
                    DocxUtil.element(mar,"end","type","dxa");
                }else if(sk.equalsIgnoreCase("padding-top")){
                    DocxUtil.element(mar,"top","w",DocxUtil.dxa(sv)+"");
                    DocxUtil.element(mar,"top","type","dxa");
                }else if(sk.equalsIgnoreCase("padding-bottom")){
                    DocxUtil.element(mar,"bottom","w",DocxUtil.dxa(sv)+"");
                    DocxUtil.element(mar,"bottom","type","dxa");
                }
            }

            Element border = DocxUtil.element(pr,"tblBorders");
            DocxUtil.border(border, styles);
            DocxUtil.background(pr, styles);
        }else if("tr".equalsIgnoreCase(name)){
            for(String sk:styles.keySet()){
                String sv = styles.get(sk);
                if(BasicUtil.isEmpty(sv)){
                    continue;
                }
                if("repeat-header".equalsIgnoreCase(sk)){
                    DocxUtil.element(pr,"tblHeader","val","true");
                }else if("min-height".equalsIgnoreCase(sk)){
                    DocxUtil.element(pr,"trHeight","hRule","atLeast");
                    DocxUtil.element(pr,"trHeight","val",(int)DocxUtil.dxa2pt(DocxUtil.dxa(sv))*20+"");
                }else if("height".equalsIgnoreCase(sk)){
                    DocxUtil.element(pr,"trHeight","hRule","exact");
                    DocxUtil.element(pr,"trHeight","val",(int)DocxUtil.dxa2pt(DocxUtil.dxa(sv))*20+"");
                }
            }
        }else if("tc".equalsIgnoreCase(name)){
            for(String sk:styles.keySet()){
                String sv = styles.get(sk);
                if(BasicUtil.isEmpty(sv)){
                    continue;
                }

                Element mar = DocxUtil.element(pr,"tcMar");
                /*DocxUtil.element(mar,"top","w","0");
                DocxUtil.element(mar,"top","type","dxa");
                DocxUtil.element(mar,"bottom","w","0");
                DocxUtil.element(mar,"bottom","type","dxa");
                DocxUtil.element(mar,"right","w","0"); // 新版本end
                DocxUtil.element(mar,"right","type","dxa");
                DocxUtil.element(mar,"end","w","0");
                DocxUtil.element(mar,"end","type","dxa");
                DocxUtil.element(mar,"left","w","0");//新版本用start,但07版本用start会报错
                DocxUtil.element(mar,"left","type","dxa");*/
                if("vertical-align".equalsIgnoreCase(sk)){
                    DocxUtil.element(pr,"vAlign", "val", sv );
                }else if("text-align".equalsIgnoreCase(sk)){
                    DocxUtil.element(pr, "jc","val", sv);
                }else if(sk.equalsIgnoreCase("white-space")){
                    DocxUtil.element(pr,"noWrap");
                }else if(sk.equalsIgnoreCase("width")){
                    DocxUtil.element(pr,"tcW","w",DocxUtil.dxa(sv)+"");
                    DocxUtil.element(pr,"tcW","type",DocxUtil.widthType(sv));
                }else if(sk.equalsIgnoreCase("padding-left")){
                    DocxUtil.element(mar,"left","w",DocxUtil.dxa(sv)+""); // 新版本用start,但07版本用start会报错
                    DocxUtil.element(mar,"left","type","dxa");
                }else if(sk.equalsIgnoreCase("padding-right")){
                    DocxUtil.element(mar,"right","w",DocxUtil.dxa(sv)+""); // 新版本用end
                    DocxUtil.element(mar,"right","type","dxa");
                    DocxUtil.element(mar,"end","w",DocxUtil.dxa(sv)+"");
                    DocxUtil.element(mar,"end","type","dxa");
                }else if(sk.equalsIgnoreCase("padding-top")){
                    DocxUtil.element(mar,"top","w",DocxUtil.dxa(sv)+"");
                    DocxUtil.element(mar,"top","type","dxa");
                }else if(sk.equalsIgnoreCase("padding-bottom")){
                    DocxUtil.element(mar,"bottom","w",DocxUtil.dxa(sv)+"");
                    DocxUtil.element(mar,"bottom","type","dxa");
                }
            }
            //
            Element padding = DocxUtil.element(pr,"tcMar");
            DocxUtil.padding(padding, styles);
            Element border = DocxUtil.element(pr,"tcBorders");
            DocxUtil.border(border, styles);
            DocxUtil.background(pr, styles);
        }
        if(pr.elements().size()==0){
            element.remove(pr);
        }
        return pr;
    }

    // 插入排版方向
    public static void setOrient(Element pr, String orient, Map<String, String> styles){
        String w = styles.get("page-size-w");
        String h = styles.get("page-size-h");
        String top = styles.get("page-margin-top");
        String right = styles.get("page-margin-right");
        String bottom = styles.get("page-margin-bottom");
        String left = styles.get("page-margin-left");
        String header = styles.get("page-margin-left");
        String footer = styles.get("page-margin-left");

        header = BasicUtil.evl(header, "851").toString();
        footer = BasicUtil.evl(footer, "992").toString();
        if("portrait".equalsIgnoreCase(orient)){
            // 竖板<w:pgMar w:top="1440" w:right="1134" w:bottom="1440" w:left="1531" w:header="851" w:footer="992" w:gutter="0"/>
            w = BasicUtil.evl(w, "11906").toString();
            h = BasicUtil.evl(h, "16838").toString();
            top = BasicUtil.evl(top, "1440").toString();
            right = BasicUtil.evl(right, "1134").toString();
            bottom = BasicUtil.evl(bottom, "1440").toString();
            left = BasicUtil.evl(left, "1531").toString();
        }else {
            // 横板
            // <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
            // <w:pgMar w:top="1531" w:right="1440" w:bottom="1134" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>
            w = BasicUtil.evl(w, "16838").toString();
            h = BasicUtil.evl(h, "11906").toString();
            top = BasicUtil.evl(top, "1531").toString();
            right = BasicUtil.evl(right, "1134").toString();
            bottom = BasicUtil.evl(bottom, "1440").toString();
            left = BasicUtil.evl(left, "1531").toString();
        }
        Element sectPr = DocxUtil.element(pr,"sectPr");
        DocxUtil.element(sectPr,"pgSz","w", w);
        DocxUtil.element(sectPr,"pgSz","h", h);
        DocxUtil.element(sectPr,"pgSz","orient", orient);

        DocxUtil.element(sectPr,"pgMar","top", top);
        DocxUtil.element(sectPr,"pgMar","right", right);
        DocxUtil.element(sectPr,"pgMar","bottom", bottom);
        DocxUtil.element(sectPr,"pgMar","left", left);
        DocxUtil.element(sectPr,"pgMar","header", header);
        DocxUtil.element(sectPr,"pgMar","footer", footer);

    }
    public static void removeAttribute(Element element, String attribute){
        Attribute att = element.attribute("w:"+attribute);
        if(null != att){
            element.remove(att);
        }
    }

    public static void removeContent(Element parent){
        List<Element> ts = DomUtil.elements(parent,"t");
        for(Element t:ts){
            t.getParent().remove(t);
        }
        List<Element> imgs = DomUtil.elements(parent,"drawing");
        for(Element img:imgs){
            img.getParent().remove(img);
        }
        List<Element> brs = DomUtil.elements(parent,"br");
        for(Element br:brs){
            br.getParent().remove(br);
        }
    }
    public static void removeElement(Element parent, String element){
        List<Element> elements = DomUtil.elements(parent, element);
        for(Element item:elements){
            item.getParent().remove(item);
        }
    }

    /**
     * 替换占位符
     * @param src Element
     * @param replaces replaces
     */
    public static void replace(Element src, Map<String, String> replaces){
        List<Element> ts = DomUtil.elements(src, "t");
        for(Element t:ts){
            String txt = t.getTextTrim();
            List<String> flags = DocxUtil.splitKey(txt);
            if(flags.size() == 0){
                continue;
            }
            Collections.reverse(flags);
            Element r = t.getParent();
            List<Element> elements = r.elements();
            int index = elements.indexOf(t);
            for(int i=0; i<flags.size(); i++){
                String flag = flags.get(i);
                String content = flag;
                String key = null;
                if(flag.startsWith("${") && flag.endsWith("}")) {
                    key = flag.substring(2, flag.length() - 1);
                    content = replaces.get(key);
                    if(null == content){
                        content = replaces.get(flag);
                    }
                }else if(flag.startsWith("{") && flag.endsWith("}")){
                    key = flag.substring(1, flag.length() - 1);
                    content = replaces.get(key);
                    if(null == content){
                        content = replaces.get(flag);
                    }
                }
                txt = txt.replace(flag, content);
            }
            t.setText(txt);
        }
    }
    /**
     * 宽度计算
     * @param src width
     * @return dxa
     */
    public static int dxa(String src){
        int dxa = 0;
        if(null != src){
            src = src.trim().toLowerCase();
            if(src.endsWith("px")){
                src = src.replace("px","");
                dxa = px2dxa(BasicUtil.parseInt(src,0));
            }else if(src.endsWith("cm")){
                src = src.replace("cm","");
                dxa = cm2dxa(BasicUtil.parseDouble(src,0d));
            }else if(src.endsWith("厘米")){
                src = src.replace("厘米","");
                dxa = cm2dxa(BasicUtil.parseDouble(src,0d));
            }else if(src.endsWith("pt")){
                src = src.replace("pt","");
                dxa = pt2dxa(BasicUtil.parseInt(src,0));
            }else if(src.endsWith("%")){
                dxa = (int)(BasicUtil.parseDouble(src.replace("%",""),0d)/100*5000);
            }else if(src.endsWith("dxa")){
                dxa = BasicUtil.parseInt(src.replace("dxa",""),0);
            }else{
                dxa = px2dxa(BasicUtil.parseInt(src,0));
            }
        }
        return dxa;
    }
    public static String widthType(String width){
        if(null != width && width.trim().endsWith("%")){
            return "pct";
        }
        if(null != width && width.trim().endsWith("dxa")){
            return "dxa";
        }
        return "dxa";
    }
    public static double PT_PER_PX = 0.75;
    public static int IN_PER_PT = 72;
    public static double CM_PER_PT = 28.3;
    public static double MM_PER_PT = 2.83;
    public static int EMU_PER_PX = 9525;
    public static int EMU_PER_DXA = 635;
    public static int px2dxa(int px){
        return pt2dxa(px2pt(px));
    }
    public static int px2dxa(double px){
        return pt2dxa(px2pt(px));
    }
    public static int pt2dxa(double pt){
        return (int)(pt*20);
    }
    public static double dxa2pt(double dxa){
        return  dxa/20;
    }
    public static double dxa2px(double dxa){
        return  pt2px(dxa2pt(dxa));
    }
    public static int px2emu(int px) {
        return px* EMU_PER_PX;
    }
    public static double dxa2emu(double dxa){
        return dxa * EMU_PER_DXA;
    }
    public static double emu2px(double emu) {
        return (emu/EMU_PER_PX);
    }

    public static double pt2px(double pt) {
        return (pt/PT_PER_PX);
    }

    public static double in2px(double in) {
        return (in2pt(in)*PT_PER_PX);
    }

    public static double px2in(double px) {
        return pt2in(px2pt(px));
    }

    public static double cm2px(double cm) {
        return (cm2pt(cm)*PT_PER_PX);
    }

    public static double px2cm(double px) {
        return pt2cm(px2pt(px));
    }

    public static double mm2px(double mm) {
        return (mm2pt(mm)*PT_PER_PX);
    }

    public static double px2mm(double px) {
        return pt2mm(px2pt(px));
    }

    public static double pt2in(double pt) {
        return (pt/IN_PER_PT);
    }

    public static double pt2mm(double mm) {
        return (mm/MM_PER_PT);
    }

    public static double pt2cm(double in) {
        return (in/CM_PER_PT);
    }

    public static double px2pt(double px) {
        return (px*PT_PER_PX);
    }

    public static double in2pt(double in) {
        return (in*IN_PER_PT);
    }

    public static double mm2pt(double mm) {
        return (mm*MM_PER_PT);
    }

    public static double cm2pt(double cm) {
        return (cm*CM_PER_PT);
    }

    public static int cm2dxa(double cm) {
        return px2dxa(cm2px(cm));
    }

    private static Map<String, Integer> fontSizes = new HashMap<String, Integer>() {
        {
            put("初号", 84);
            put("小初", 72);
            put("一号", 52);
            put("小一", 48);
            put("二号", 44);
            put("小二", 36);
            put("三号", 33);
            put("小三", 30);
            put("四号", 28);
            put("小四", 24);
            put("五号", 21);
            put("小五", 18);
            put("六号", 15);
            put("小六", 13);
            put("七号", 11);
            put("八号", 10);
        }
    };

}
