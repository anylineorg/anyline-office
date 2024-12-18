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

import org.anyline.entity.DataSet;
import org.anyline.log.Log;
import org.anyline.log.LogProxy;
import org.anyline.office.docx.entity.WDocument;
import org.anyline.office.docx.util.DocxUtil;
import org.anyline.office.util.Context;
import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;
import org.anyline.util.ConfigTable;
import org.anyline.util.regular.RegularUtil;
import org.dom4j.Element;

import java.io.File;
import java.util.*;

public abstract class AbstractTag implements Tag {
    protected static Log log = LogProxy.get(AbstractTag.class);
    protected List<Tag> children = new ArrayList<>();
    protected WDocument doc;
    protected List<Element> tops = new ArrayList<>(); // 标签所在顶层p或table(包括head body foot)
    protected List<Element> ts = new ArrayList<>();
    protected Context context = new Context();
    protected String text;
    protected String valueKey = ConfigTable.DEFAULT_PRIMARY_KEY;
    protected String textKey = "NM";
    protected String var;
    protected String ref;
    protected Object data;
    protected String property;
    protected String head;

    public void init(WDocument doc) {
        this.doc = doc;
        this.context = doc.context().clone();
    }
    public void prepare(){
        head = RegularUtil.fetchTagHead(text);
        String vk = fetchAttributeString(head, "valueKey", "vk");
        if(BasicUtil.isNotEmpty(vk)){
            valueKey = vk;
        }
        String tk = fetchAttributeString(text, "textKey", "tk");
        if(BasicUtil.isNotEmpty(tk)){
            textKey = tk;
        }

        var = fetchAttributeString(head, "var");
        property = fetchAttributeString(head, "property", "p");
        /*String name = TagUtil.name(text, "aol:");
        String foot = "</aol:"+name+">";
        if(head.endsWith("/>")){
            foot = null;
        }
        //标签起止top
        Element first_top = tops.get(0);
        Element last_top = tops.get(tops.size()-1);
        //标签起止top 文本
        String first_top_text = DocxUtil.text(first_top);
        String last_top_text = DocxUtil.text(last_top);

        //定位标签体所在的tops
        int body_top_first_index = 0;
        int body_top_last_index = tops.size()-1;
        if(first_top_text.trim().endsWith(head)){
            //以标签头结尾 标签体在下一个top
            body_top_first_index = 1;
        }
        if(null != foot && last_top_text.trim().startsWith(foot)){
            //以标签必尾开头 标签体截止到上一个top
            body_top_last_index = body_top_last_index -1 ;
        }
        //标签体所在tops
        for(int i=body_top_first_index; i <= body_top_last_index; i++){
            //inners.add(tops.get(i));
        }*/

    }

    public void release(){
        ts.clear();
        children.clear();
        context = new Context();
    }
    public void ref(String ref){
        this.ref = ref;
    }
    public String ref(){
        return ref;
    }

    public void context(Context context) {
        this.context = context;
    }

    public Context context() {
        return context;
    }

    public void variable(String key, Object value) {
        context.variable(key, value);
    }

    public void variable(Map<String, Object> values) {
        context.variable(values);
    }

    public void wts(List<Element> wts) {
        this.ts = wts;
    }

    public List<Element> wts() {
        return ts;
    }

    /**
     * 标签内的wt所在的顶层p或table
     * 注意如果是与标签在同一个wp中的 设置top=wt
     * @return list
     */
    public List<Element> tops() {
        return tops;
    }
    public void tops(List<Element> tops) {
        this.tops = tops;
    }
    /**
     * 设置占位符替换值 在调用save时执行替换<br/>
     * 注意如果不解析的话 不会添加自动${}符号 按原文替换,是替换整个文件的纯文件，包括标签名在内
     *
     * @param key     占位符
     * @param content 替换值
     */
    public Tag replace(String key, String content) {
        context.replace(key, content);
        return this;
    }

    public Tag replace(boolean parse, String key, File... words) {
        return replace(parse, key, BeanUtil.array2list(words));
    }

    public Tag replace(String key, File... words) {
        return replace(true, key, BeanUtil.array2list(words));
    }

    public Tag replace(boolean parse, String key, List<File> words) {
        if (null != words) {
            StringBuilder content = new StringBuilder();
            for (File word : words) {
                content.append("<word>").append(word.getAbsolutePath()).append("</word>");
            }
            context.replace(parse, key, content.toString());
        }
        return this;
    }

    public Tag replace(String key, List<File> words) {
        return replace(true, key, words);
    }

    protected String fetchAttributeString(String text, String ... attributes){
        for(String attribute:attributes){
            String value = RegularUtil.fetchAttributeValue(text, attribute);
            if(null == value && null != ref){
                value = RegularUtil.fetchAttributeValue(ref, attribute);
            }
            if(null != value){
                if(value.contains("${") && value.contains("}")){
                    List<String> ks = DocxUtil.splitKey(value);
                    Collections.reverse(ks);
                    value = "";
                    for(String k:ks){
                        if(BasicUtil.checkEl(k)){
                            String v = context.string(false, k);
                            if(null == v){
                                v = "";
                            }
                            value += v;
                        }else{
                            value += k;
                        }
                    }
                }
                return value;
            }
        }
        return null;
    }
    protected Object fetchAttributeData(String text, String ... attributes){
        for(String attribute:attributes){
            String value = RegularUtil.fetchAttributeValue(text, attribute);
            if(null == value && null != ref){
                value = RegularUtil.fetchAttributeValue(ref, attribute);
            }
            if(null != value){
                if(BasicUtil.checkEl(value)){
                    Object data = context.data(value);
                    if(null != data){
                        return data;
                    }
                }
                return value;
            }
        }
        return null;
    }
    protected String body(String text, String name){
        String body = null;
        try {
            body = RegularUtil.fetchTagBody(text, "aol:"+name);
            if (null == body && null != ref) {
                body = RegularUtil.fetchTagBody(ref, "aol:pre");
                if(null == body){
                    body = RegularUtil.fetchTagBody(ref, "aol:"+name);
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return body;
    }
    protected Object data(){
        Object data = fetchAttributeData(head, "data", "d", "items", "is");
        if(null == data){
            return null;
        }
        String distinct = fetchAttributeString(head, "distinct", "ds");
        Integer index = BasicUtil.parseInt(fetchAttributeString(head, "index", "i"), null);
        Integer begin = BasicUtil.parseInt(fetchAttributeString(head, "begin", "b"), null);
        Integer end = BasicUtil.parseInt(fetchAttributeString(head, "end", "e"), null);
        Integer qty = BasicUtil.parseInt(fetchAttributeString(head, "qty", "q"), null);
        String selector = fetchAttributeString(head, "selector","st");

        if(data instanceof Collection) {
            Collection items = (Collection) data;
            if(BasicUtil.isNotEmpty(selector)) {
                items = BeanUtil.select(items,selector.split(","));
            }
            if(index != null) {
                int i = 0;
                data = null;
                for(Object item:items) {
                    if(index ==i) {
                        data = item;
                        break;
                    }
                    i ++;
                }
            }else{
                int[] range = BasicUtil.range(begin, end, qty, items.size());
                if(items instanceof DataSet) {
                    data = ((DataSet) items).cuts(range[0], range[1]);
                }else {
                    data = BeanUtil.cuts(items, range[0], range[1]);
                }
            }
            if(null != distinct && data instanceof Collection) {
                if(data instanceof DataSet){
                    DataSet set = (DataSet) data;
                    data = set.distinct(distinct.split(","));
                }else{
                    data = BeanUtil.distinct((Collection) data, distinct.split(","));
                }
            }
        }
        return data;
    }
    public String text(){
        return text;
    }
    public void text(String text){
        this.text = text;
    }

    /**
     * 输出文本
     * 输出到第一个t,清空其他t
     * @param result 输出内容
     */
    protected void output(Object result){
        int size = ts.size();
        Element t = ts.get(0);
        if(null == result){
            result = doc.getPlaceholderDefault();
        }

        if(BasicUtil.isNotEmpty(var)){
            doc.variable(var, result);
            DocxUtil.remove(t);
        }else{
            t.setText(result.toString());
        }
        if(size > 1) {
            for (int i = 1; i < size; i++) {
                DocxUtil.remove(ts.get(i));
            }
        }
    }

}