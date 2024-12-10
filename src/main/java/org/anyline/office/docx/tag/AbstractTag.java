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
import org.anyline.office.docx.entity.WDocument;
import org.anyline.util.BeanUtil;
import org.anyline.util.regular.RegularUtil;
import org.dom4j.Element;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public abstract class AbstractTag implements Tag {
    protected List<Tag> children = new ArrayList<>();
    protected WDocument doc;
    protected List<Element> tops = new ArrayList<>();
    protected List<Element> wts = new ArrayList<>();
    protected Context context = new Context();
    protected String ref;

    public void init(WDocument doc) {
        this.doc = doc;
        this.context = doc.context().clone();
    }
    public void release(){
        wts.clear();
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
        this.wts = wts;
    }

    public List<Element> wts() {
        return wts;
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


    public String parse(String text) throws Exception {
        return text;
    }
    protected String fetchAttributeValue(String text, String ... attributes){
        for(String attribute:attributes){
            String value = RegularUtil.fetchAttributeValue(text, attribute);
            if(null == value && null != ref){
                value = RegularUtil.fetchAttributeValue(ref, attribute);
            }
            if(null != attribute){
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
}