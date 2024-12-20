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
import org.anyline.office.docx.entity.WElement;
import org.anyline.office.util.TagUtil;
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
            List<Element> ts = DomUtil.elements(true, document.getRootElement(),"t");
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
     * 前一个节点
     * @param element element
     * @return element
     */
    public static Element prevByName(Element element){
        Element prev = null;
        List<Element> elements = DomUtil.elements(true, top(element), element.getName());
        int index = elements.indexOf(element);
        if(index > 0){
            prev = elements.get(index -1);
        }
        return prev;
    }
    public static Element prevByName(Element parent, Element element){
        Element prev = null;
        List<Element> elements = DomUtil.elements(true, parent, element.getName());
        int index = elements.indexOf(element);
        if(index > 0){
            prev = elements.get(index -1);
        }
        return prev;
    }
    public static Element nextByName(Element element){
        Element prev = null;
        List<Element> elements = DomUtil.elements(true, top(element), element.getName());
        int index = elements.indexOf(element);
        if(index < elements.size()-1 && index > 0){
            prev = elements.get(index + 1);
        }
        return prev;
    }
    public static Element nextByName(Element parent, Element element){
        Element prev = null;
        List<Element> elements = DomUtil.elements(true, parent, element.getName());
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
     * @param elements 集合
     * @return String
     */
    public static String text(List<Element> elements){
        StringBuilder builder = new StringBuilder();
        for(Element element:elements){
            String txt = text(element);
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
            List<Element> elements = rp.elements();
            int index = elements.indexOf(ref)+1;
            elements.remove(src);
            if(index > elements.size()-1){
                elements.add(src);
            }else {
                elements.add(index, src);
            }
            src.setParent(rp);
        } else {
            // ref更下级
            after(src, ref.getParent());
        }
    }

    public static void after(WElement src, Element ref){
        after(src.getSrc(), ref);
    }
    public static void after(WElement src, WElement ref){
        after(src.getSrc(), ref.getSrc());
    }
    public static void after(Element src, WElement ref){
        after(src, ref.getSrc());
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
        if(src == ref){
            //ref取父标签后可能与src一样
            return;
        }
        Element rp = ref.getParent();
        Element sp = src.getParent();

        // 同级
        if(rp == sp || null == sp){
            List<Element> elements = rp.elements();
            int index = elements.indexOf(ref);
            elements.remove(src);
            elements.add(index, src);
            src.setParent(rp);
        } else {
            // ref更下级
            before(src, ref.getParent());
        }


        ///////////////
        /*List<Element> elements = ref.getParent().elements();
        int index = elements.indexOf(ref);
        while (!elements.contains(src)){
            src = src.getParent();
            if(null == src){
                return;
            }
        }
        elements.remove(src);
        elements.add(index, src);
*/
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
    public static String attributeValue(Element element, String key){
        String value = element.attributeValue(key);
        if(null == value){
            value = element.attributeValue("w:"+key);
        }
        return value;
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
        String name = tag;
        if(!name.startsWith("w:")){
            name = "w:" + name;
        }
        Element element =  parent.addElement(name);
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
            }else if(node instanceof Element){
                text += text((Element)node);
            }
        }
        return text;
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
        List<Element> items = contents(box);
        //<w:bookmarkStart w:id="0" w:name="a"/>
        int size = items.size();
        String full = "";
        List<Element> merges = new ArrayList<>();
        for(int i=0; i<size; i++){
            Element item = items.get(i);
            String name = item.getName();
            if(name.equals("br") || name.contains("bookmark")){
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

    public static void mergeText(List<Element> ts){
        int size = ts.size();
        if(size > 1){
            String text = DocxUtil.text(ts);
            text = TagUtil.format(text);
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
        //log.warn("删除:{}", text(element));
        if(null != parent){
            parent.remove(element);
            if(DocxUtil.isEmpty(parent)) {
                remove(parent);
            }
        }/*else{
            log.error("重复删除:{}", element.getText());
        }*/
    }

    /**
     * 是否有内容(表格、文本、图片)
     * @param element element
     * @return boolean
     */
    public static boolean isEmpty(Element element){
        List<Element> elements = contents(element);;
        for(Element item:elements){
            String name = item.getName();
            if(name.equalsIgnoreCase("drawing")){
                return false;
            }
            if(name.contains("bookmark")){
                return false;
            }
            if(name.equalsIgnoreCase("br")){
                return false;
            }
            if(name.equalsIgnoreCase("tbl")){
                return false;
            }
            if(name.equalsIgnoreCase("t")){
                if(item.getText().length() > 0){
                    return false;
                }
            }
        }
        if(element.getText().length() > 0){
            return false;
        }
        return true;
    }
    public static List<Element> contents(Element element){
        return DomUtil.elements(true, element, "drawing,tbl,t,br,bookmarkStart,bookmarkEnd,sectPr");
    }
    public static List<Element> contents(List<Element> elements){
        List<Element> list = new ArrayList<>();
        for(Element element:elements){
            list.addAll(contents(element));
        }
        return list;
    }

    public static List<Element> contents(WElement element){
        return contents(element.getSrc());
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
            StyleUtil.border(border, styles);
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
            StyleUtil.border(border, styles);
            StyleUtil.font(pr, styles);
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
            StyleUtil.border(border, styles);
            StyleUtil.background(pr, styles);
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
            StyleUtil.padding(padding, styles);
            Element border = DocxUtil.element(pr,"tcBorders");
            StyleUtil.border(border, styles);
            StyleUtil.background(pr, styles);
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
        List<Element> ts = DomUtil.elements(true, parent,"t");
        for(Element t:ts){
            t.getParent().remove(t);
        }
        List<Element> imgs = DomUtil.elements(true, parent,"drawing");
        for(Element img:imgs){
            img.getParent().remove(img);
        }
        List<Element> brs = DomUtil.elements(true, parent,"br");
        for(Element br:brs){
            br.getParent().remove(br);
        }
    }
    public static void removeElement(Element parent, String element){
        List<Element> elements = DomUtil.elements(true, parent, element);
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
        List<Element> ts = DomUtil.elements(true, src, "t");
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

}
