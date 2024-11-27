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

import org.anyline.office.docx.entity.WDocument;
import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;
import org.anyline.util.regular.RegularUtil;
import org.dom4j.Element;

import java.io.File;
import java.util.*;

public abstract class AbstractTag implements Tag{
    protected List<Tag> children = new ArrayList<>();
    protected WDocument doc;
    protected Element wt;
    protected LinkedHashMap<String, String> replaces = new LinkedHashMap<>();
    /**
     * 与replaces不同的是values中可以包含复杂结构
     */
    protected LinkedHashMap<String, Object> variables = new LinkedHashMap<>();
    /**
     * 文本原样替换，不解析原文没有${}的也不要添加
     */
    protected LinkedHashMap<String, String> txt_replaces = new LinkedHashMap<>();

    public void init(WDocument doc){
        this.doc = doc;
        this.replaces.putAll(doc.getReplaces());
        this.txt_replaces.putAll(doc.getTextReplaces());
        this.variables.putAll(doc.variables());
    }
    public void variable(String key, Object value) {
        variables.put(key, value);
    }
    public void variable(Map<String, Object> values) {
        variables.putAll(values);
    }
    public void wt(Element wt){
        this.wt = wt;
    }
    public Element wt(){
        return wt;
    }
    /**
     * 设置占位符替换值 在调用save时执行替换<br/>
     * 注意如果不解析的话 不会添加自动${}符号 按原文替换,是替换整个文件的纯文件，包括标签名在内
     * @param parse 是否解析标签 true:解析HTML标签 false:直接替换文本
     * @param key 占位符
     * @param content 替换值
     */
    public Tag replace(boolean parse, String key, String content){
        if(null == key && key.trim().length()==0){
            return this;
        }
        if(parse) {
            replaces.put(key, content);
        }else{
            txt_replaces.put(key, content);
        }
        return this;
    }
    public Tag replace(String key, String content){
        return replace(true, key, content);
    }
    public Tag replace(boolean parse, String key, File... words){
        return replace(parse, key, BeanUtil.array2list(words));
    }
    public Tag replace(String key, File ... words){
        return replace(true, key, BeanUtil.array2list(words));
    }
    public Tag replace(boolean parse, String key, List<File> words){
        if(null != words) {
            StringBuilder content = new StringBuilder();
            for(File word:words) {
                content.append("<word>").append(word.getAbsolutePath()).append("</word>");
            }
            if(parse) {
                replaces.put(key, content.toString());
            }else{
                txt_replaces.put(key, content.toString());
            }
        }
        return this;
    }

    public Tag replace(String key, List<File> words){
        return replace(true, key, words);
    }
    public LinkedHashMap<String, Object> variables(){
        return variables;
    }
    public String parse(String text) throws Exception{
        return text;
    }

    public Object data(String key){
        if(null == key){
            return null;
        }
        Object data = key;
        if(BasicUtil.checkEl(key)){
            //${users}
            key = key.substring(2, key.length() - 1);
            data = variables.get(key);
            if(null == data){
                data = replaces.get(key);
            }
            if(null == data){
                data = txt_replaces.get(key);
            }

            if(null == data){
                if(key.contains(".")){
                    data = BeanUtil.value(variables, key);
                }
            }
        }else if(key.startsWith("{") && key.endsWith("}")){
            key = key.replace("{", "").replace("}", "");
            data = key;
            if(key.contains(",")){
                String[] ks = key.split(",");
                List<String> list = new ArrayList<>();
                for(String k:ks){
                    //{0:关,1:开}
                    if(k.contains(":")){
                        String[] kv = k.split(":");
                        if(kv.length ==2){
                            Map map = new HashMap();
                            map.put(kv[0], kv[1]);
                        }
                    }else {
                        //{FI,CO,HR}
                        list.add(k);
                    }
                }
                data = list;
            }
        }
        return data;
    }
    /**
     * 替换占位符
     * @param text 原文
     * @return String
     */
    public String placeholder(String text){
        String result = text;
        for(String key:replaces.keySet()){
            String value = replaces.get(key);
            if(null == value){
                value = "";
            }
            result = result.replace("${" + key + "}", value);
        }
        for(String key:txt_replaces.keySet()){
            String value = txt_replaces.get(key);
            if(null == value){
                value = "";
            }
            result = result.replace("${" + key + "}", value);
        }
        for(String key:variables.keySet()){
            Object value = variables.get(key);
            if(null == value){
                value = "";
            }
            result = result.replace("${" + key + "}", value.toString());
        }
        //检测复合占位符
        try {
            List<String> ks = RegularUtil.fetch(text, "\\$\\{.*?\\}");
            for(String k:ks){
                Object value = data(k);
                result = result.replace(k, value.toString());
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        if(null == result){
            result = "";
        }
        return result;
    }
}
