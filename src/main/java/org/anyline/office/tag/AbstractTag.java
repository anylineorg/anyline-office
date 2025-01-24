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
import org.anyline.office.util.TagUtil;
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
    protected List<Element> contents = new ArrayList<>();
    protected List<Element> tops = new ArrayList<>();
    protected Context context = new Context();
    protected String text;
    protected String valueKey = ConfigTable.DEFAULT_PRIMARY_KEY;
    protected String textKey = "NM";
    protected String var;
    protected String ref;
    protected Object data;
    protected String property;
    protected TagBox box;
    protected Tag parent;
    protected Element last;

    public void init(WDocument doc) {
        this.doc = doc;
        this.context = doc.context().clone();
    }
    public void prepare(){
        tops = TagUtil.tops(contents);
        box = new TagBox(doc);
        box.contents(contents);
        TagElement head = new TagElement();
        TagElement foot = new TagElement();
        box.head(head);
        Element t0 = null;
        Element t1 = null;
        if(!contents.isEmpty()){
            t0 = contents.get(0);
            head.element(t0);
            List<Element> head_contents = DocxUtil.contents(tops.get(0));
            int index = head_contents.indexOf(t0);
            head.index(index);
            head.first(index == 0);
            head.last(index == head_contents.size()-1);
            head.text(t0.getText());
            t0.setText("");

            if(this.contents.size() > 0){
                t1 = this.contents.get(this.contents.size()-1);
                foot.element(t1);
                // foot在top中的下标 注意区分是在首行(中有一行)还是尾行
                List<Element> foot_contents = DocxUtil.contents(tops.get(tops.size()-1));
                index = foot_contents.indexOf(t1);
                foot.index(index);
                foot.first(index == 0);
                foot.last(index == foot_contents.size()-1);
                foot.text(t1.getText());
                t1.setText("");
                box.foot(foot);
            }
        }else{
            String head_text = RegularUtil.fetchTagHead(text);
            head.text(head_text);
        }

        box.tops(tops);
        String vk = fetchAttributeString("valueKey", "vk");
        if(BasicUtil.isNotEmpty(vk)){
            valueKey = vk;
        }
        String tk = fetchAttributeString("textKey", "tk");
        if(BasicUtil.isNotEmpty(tk)){
            textKey = tk;
        }

        var = fetchAttributeString("var");
        property = fetchAttributeString("property", "p");
    }

    public void release(){
        contents.clear();
        children.clear();
        context = new Context();
    }
    @Override
    public String parse() throws Exception {
        return "";
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

    public void contents(List<Element> contents) {
        this.contents = contents;
    }

    public List<Element> contents() {
        return contents;
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

    protected String fetchAttributeString(String ... attributes){
        String text = box.head().text();
        text = TagUtil.format(text);
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
    protected Object fetchAttributeData(String ... attributes){
        String text = box.head().text();
        for(String attribute:attributes){
            String value = RegularUtil.fetchAttributeValue(text, attribute);
            if(null == value && null != ref){
                value = RegularUtil.fetchAttributeValue(ref, attribute);
            }
            if(null != value){
                if(BasicUtil.checkEl(value)){
                    return context.data(value);
                }
                return value;
            }
        }
        return null;
    }
    protected String body(String text, String name){
        String body = null;
        try {
            body = RegularUtil.fetchTagBody(text, doc.namespace()+":"+name);
            if (null == body && null != ref) {
                body = RegularUtil.fetchTagBody(ref, doc.namespace()+":pre");
                if(null == body){
                    body = RegularUtil.fetchTagBody(ref, doc.namespace()+":"+name);
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return body;
    }
    protected Object data(){
        return data(true);
    }

    /**
     *
     * @param filter 是否根据begin end过滤
     * @return Object
     */
    protected Object data(boolean filter){
        Object data = fetchAttributeData("data", "d", "items", "is");
        if(null == data){
            return null;
        }
        String distinct = fetchAttributeString("distinct", "ds");
        Integer index = BasicUtil.parseInt(fetchAttributeString("index", "i"), null);
        Integer begin = BasicUtil.parseInt(fetchAttributeString("begin", "start", "b"), null);
        Integer end = BasicUtil.parseInt(fetchAttributeString("end", "e"), null);
        Integer qty = BasicUtil.parseInt(fetchAttributeString("qty", "q"), null);
        String selector = fetchAttributeString("selector","st");

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
                if(filter) {
                    int[] range = BasicUtil.range(begin, end, qty, items.size());
                    if (items instanceof DataSet) {
                        data = ((DataSet) items).cuts(range[0], range[1]);
                    } else {
                        data = BeanUtil.cuts(items, range[0], range[1]);
                    }
                }
            }
            if(null != distinct && data instanceof Collection) {
                if(data instanceof DataSet){
                    DataSet set = (DataSet) data;
                    data = set.distinct(false, distinct.split(","));
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
        int size = contents.size();
        Element t = contents.get(0);

        if(BasicUtil.isNotEmpty(var)){
            context.variable(var, result);
            Context pc = context.parent();
            if(null == pc){
                pc = doc.context();
            }
            pc.variable(var, result);
            DocxUtil.remove(t);
        }else{
            if(null == result){
                result = doc.getPlaceholderDefault();
            }
            t.setText(result.toString());
        }
        if(size > 1) {
            for (int i = 1; i < size; i++) {
                DocxUtil.remove(contents.get(i));
            }
        }
    }

    @Override
    public Tag parent() {
        return parent;
    }

    @Override
    public void parent(Tag parent) {
        this.parent = parent;
    }
    @Override
    public Element last() {
        if(null == last && null != parent){
            return parent.last();
        }
        return last;
    }

    @Override
    public void last(Element last) {
        if(null != parent){
            parent.last(last);
        } else {
            this.last = last;
        }
    }
    public List<Element> tops(){
        return tops;
    }
    public void tops(List<Element> tops){
        this.tops = tops;
    }
}