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

import org.anyline.handler.Uploader;
import org.anyline.office.docx.util.DocxUtil;
import org.anyline.util.BasicUtil;
import org.anyline.util.DomUtil;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class WElement {

    protected WDocument root;
    protected Element src;
    public void reload(){

    }
    /**
     * 替换限制范围内占位符
     * @param replaces replaces
     */
    public void replace(Map<String, String> replaces){
        DocxUtil.replace(getSrc(), replaces);
    }
    /**
     * 删除行内文本内容
     */
    public void removeContent(){
        DocxUtil.removeContent(getSrc());
    }
    public WDocument getDoc() {
        return root;
    }

    public void setDoc(WDocument doc) {
        this.root = doc;
    }

    public Element getSrc() {
        if(null == src){
            reload();
        }
        return src;
    }

    public void setSrc(Element src) {
        this.src = src;
    }

    /**
     * 所有书签
     * @return list
     */
    public List<Element> bookmarks(){
        List<Element> bookmarks = DomUtil.elements(getSrc(), "bookmarkStart");
        return bookmarks;
    }

    public List<String> placeholders(){
        return placeholders("\\$\\{.*?\\}");
    }

    /**
     * 占位符所在元素(w:t标签)
     * @param element 是否返回占位符所在元素
     * @return List
     */
    public List<Element> placeholders(boolean element){
        return placeholders(element, "\\$\\{.*?\\}");
    }
    public List<Element> placeholders(boolean element, String regex){
        List<Element> list = new ArrayList<>();
        try {
            List<Element> ts = DomUtil.elements(getSrc(), "t");
            for(Element t:ts){
                String txt = t.getTextTrim();
                List<String> flags = DocxUtil.splitKey(txt, regex);
                if(flags.isEmpty()){
                    continue;
                }
                list.add(t);
            }

        }catch (Exception e){
            e.printStackTrace();
        }
        return list;
    }
    /**
     * 所有${key}格式的占位符
     * @param regex 正则
     * @return list 不包括前后缀标识
     */
    public List<String> placeholders(String regex){
        List<String> list = new ArrayList<>();
        try {
            List<Element> ts = DomUtil.elements(getSrc(), "t");
            for(Element t:ts){
                String txt = t.getTextTrim();
                List<String> flags = DocxUtil.splitKey(txt, regex);
                if(flags.isEmpty()){
                    continue;
                }
                for(String flag:flags){
                    String key = null;
                    if(flag.startsWith("${") && flag.endsWith("}")) {
                        key = flag.substring(2, flag.length() - 1);
                        list.add(key);
                    }
                }
            }

        }catch (Exception e){
            e.printStackTrace();
        }
        return list;
    }

    /**
     * 获取每个t标签中的文本，注意经常会发生word自己把文字拆成多个t标签，
     * 可以调用getTexts()返回一个完整文本
     * @return List
     */
    public List<String> getTextList(){
        List<String> texts = new ArrayList<>();
        List<Element> ts = DomUtil.elements(getSrc(), "t");
        for(Element t:ts){
            texts.add(t.getTextTrim());
        }
        return texts;
    }

    /**
     * 合并所有t标签文本成一个文本返回,如果要分开返回可以调用getTextList
     * @return String
     */
    public String getTexts(){
        List<String> list = getTextList();
        StringBuilder builder = new StringBuilder();
        for(String item:list){
            if(null != item){
                builder.append(item);
            }
        }
        return builder.toString();
    }

    /**
     * 查找直接子级t标签文本，tc中可能有多个 t标签，可以调用getTextList或getText返回所有t标签文本(不限层级)
     * @return String
     */
    public String getText(){
        Element t = getSrc().element("t");
        if(null != t){
            return t.getText();
        }
        return null;
    }

    /**
     * 计算下标
     * @param index 下标 从0开始 -1表示最后一行 -2表示倒数第2行
     * @param size 总行数
     * @return 最终下标
     */
    protected int index(Integer index, int size){
        if(null == index){
            return 0;
        }
        return BasicUtil.index(index, size);
    }

    public Element child(Element parent, String ... tags){
        if(null != tags){
            Element element = parent;
            for(String tag:tags){
                if(null != element){
                    element = element.element(tag);
                }
            }
            return element;
        }
        return null;
    }

    protected void t(StringBuilder builder, int lvl){
        for(int i=0; i<lvl; i++){
            builder.append("\t");
        }
    }
    public LinkedHashMap<String, String> styles(){
        LinkedHashMap<String, String> styles = new LinkedHashMap<>();
        return styles;
    }
    public LinkedHashMap<String, String> styles(StringBuilder builder){
        LinkedHashMap<String, String> styles = styles();
        if(!styles.isEmpty()){
            builder.append(" style = '");
            for (String key:styles.keySet()){
                builder.append(key).append(":").append(styles.get(key)).append(";");
            }
            builder.append("'");
        }
        return styles;
    }
    public String color(String color){
        return color;
    }
    public String html(){
        return html(null, 0);
    }
    public String html(int lvl){
        return html(null, lvl);
    }
    public String html(Uploader uploader){
        return html(uploader, 0);
    }
    public String html(Uploader uploader, int lvl){
        return "";
    }
}
