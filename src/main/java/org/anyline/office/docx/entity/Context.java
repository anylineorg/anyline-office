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

import org.anyline.adapter.KeyAdapter;
import org.anyline.entity.DataRow;
import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;
import org.anyline.util.regular.RegularUtil;

import java.io.File;
import java.util.*;

public class Context {
    /**
     * 文本原样替换，不解析原文中的标签,没有${}的也不要添加
     */
    private LinkedHashMap<String, String> htmls = new LinkedHashMap<>();
    private LinkedHashMap<String, String> texts = new LinkedHashMap<>();
    /**
     * 与replaces不同的是values中可以包含复杂结构
     */
    private LinkedHashMap<String, Object> variables = new LinkedHashMap<>();
    public LinkedHashMap<String, String> replaces(){
        return htmls;
    }
    public LinkedHashMap<String, String> texts(){
        return texts;
    }
    public LinkedHashMap<String, Object> variables(){
        return variables;
    }

    /**
     * 设置占位符替换值 在调用save时执行替换<br/>
     * 注意如果不解析的话 不会添加自动${}符号 按原文替换,是替换整个文件的纯文件，包括标签名在内
     * @param parse 是否解析标签 true:解析HTML标签 false:直接替换文本
     * @param key 占位符
     * @param content 替换值
     */
    public Context replace(boolean parse, String key, String content){
        if(null == key && key.trim().length()==0){
            return this;
        }
        if(parse) {
            htmls.put(key, content);
        }else{
            texts.put(key, content);
        }
        return this;
    }

    public Context variable(String key, Object value){
        if(null != key) {
            variables.put(key, value);
        }
        return this;
    }

    public Context variable(Map<String, Object> values) {
        values.putAll(values);
        return this;
    }
    public Context replace(String key, String content){
        return replace(true, key, content);
    }
    public Context replace(boolean parse, String key, File... words){
        return replace(parse, key, BeanUtil.array2list(words));
    }
    public Context replace(String key, File ... words){
        return replace(true, key, BeanUtil.array2list(words));
    }
    public Context replace(boolean parse, String key, List<File> words){
        if(null != words) {
            StringBuilder content = new StringBuilder();
            for(File word:words) {
                content.append("<word>").append(word.getAbsolutePath()).append("</word>");
            }
            if(parse) {
                htmls.put(key, content.toString());
            }else{
                texts.put(key, content.toString());
            }
        }
        return this;
    }

    public void replace(String key, List<File> words){
        replace(true, key, words);
    }

    public String string(String key){
        Object data = data(key);
        if(null != data){
            return data.toString();
        }
        return null;
    }
    public Object data(String key){
        if(null == key){
            return "";
        }
        key = key.trim();
        Object data = key;
        if(BasicUtil.checkEl(key)){
            //${users}
            key = key.substring(2, key.length() - 1);
            data = variables.get(key);
            if(null == data){
                data = htmls.get(key);
            }
            if(null == data){
                data = texts.get(key);
            }
            if(null == data){
                if(key.contains(".")){
                    //user.dept.name
                    String[] ks = key.split("\\.");
                    int size = ks.length;
                    if(size > 1) {
                        data = variables.get(ks[0]);
                        for (int i = 1; i < size; i++) {
                            String k = ks[i];
                            if(null == data){
                                break;
                            }
                            if(data instanceof Collection && "size".equals(k)){
                                data = ((Collection)data).size();
                            }else {
                                if(data instanceof String){
                                    String str = (String)data;
                                    if(str.startsWith("{") && str.endsWith("}")){
                                        try{
                                            data = DataRow.parseJson(KeyAdapter.KEY_CASE.SRC, str);
                                        }catch (Exception e){
                                            e.printStackTrace();
                                        }
                                    }
                                }
                                data = BeanUtil.getFieldValue(data, k);
                            }
                        }
                    }
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
        if(null == data){
            data = "";
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
        for(String key: htmls.keySet()){
            String value = htmls.get(key);
            if(null == value){
                value = "";
            }
            result = result.replace("${" + key + "}", value);
        }
        for(String key:texts.keySet()){
            String value = texts.get(key);
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
            List<String> ks = RegularUtil.fetch(result, "\\$\\{.*?\\}");
            for(String k:ks){
                Object value = data(k);
                if(null == value){
                    value = "";
                }
                if(BasicUtil.isEmpty(value)){
                    if(k.startsWith("${__")){
                        continue;
                    }
                }
                result = result.replace(k, value.toString());
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        if(null == result){
            if(text.startsWith("__")){
                return text;
            }
            result = "";
        }
        return result;
    }
    public Context clone(){
        Context clone = new Context();
        clone.htmls.putAll(htmls);
        clone.texts.putAll(texts);
        clone.variables.putAll(variables);
        return clone;
    }
}
