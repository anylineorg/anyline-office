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

import ognl.Ognl;
import ognl.OgnlContext;
import org.anyline.adapter.KeyAdapter;
import org.anyline.entity.DataRow;
import org.anyline.entity.DataSet;
import org.anyline.log.Log;
import org.anyline.log.LogProxy;
import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;
import org.anyline.util.DefaultOgnlMemberAccess;
import org.anyline.util.regular.RegularUtil;

import java.io.File;
import java.util.*;

public class Context {
    private static Log log = LogProxy.get(Context.class);
    private Context parent = null;
    /**
     * 文本原样替换，不解析原文中的标签,没有${}的也不要添加
     */
    private LinkedHashMap<String, String> htmls = new LinkedHashMap<>();
    private LinkedHashMap<String, String> texts = new LinkedHashMap<>();
    private String placeholderDefault = "";
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

    public String getPlaceholderDefault() {
        return placeholderDefault;
    }

    public void setPlaceholderDefault(String placeholderDefault) {
        this.placeholderDefault = placeholderDefault;
    }

    public String string(boolean def, String key){
        Object data = data(key);
        if(null == data && def){
            data = placeholderDefault;
        }
        if(null != data){
            return data.toString();
        }
        return null;
    }
    public String string(String key){
        return string(true, key);
    }
    public Object data(String key){
        if(null == key){
            return null;
        }
        //a:b:c:123
        //a:b:c:'123'
        //list[list.size-1]
        //list[0].name

        String kk = key.trim();
        Object data = null;
        if(BasicUtil.checkEl(kk)) {
            kk = kk.substring(2, kk.length() - 1).trim();
        }
        if(kk.startsWith("aov:")){
            return aov(kk);
        }
        if(kk.startsWith("{") && kk.endsWith("}")) {
            kk = kk.replace("{", "").replace("}", "");
            if(kk.contains(",")){
                String[] ks = kk.split(",");
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
            if(null != data){
                return data;
            }
        }
        //多变量顺位 取第一个非空
        if(kk.contains(":")){
            kk = TagUtil.format(kk);
            String[] ks = kk.split(":");
            int len = ks.length;
            for(int idx = 0; idx <len; idx ++){
                String k = ks[idx];
                k = k.trim();
                if(k.isEmpty()){
                    continue;
                }
                Object v = data(k);
                if(null != v){
                    return v;
                }
            }
        }
        data = variables.get(kk);
        if(null == data){
            data = htmls.get(kk);
        }
        if(null == data){
            data = texts.get(kk);
        }
        if(null == data){
            //Collection[index] ognl不支持
            //list[list.size-1].USER.NAME
            try{
                OgnlContext ognl = new OgnlContext(null, null, new DefaultOgnlMemberAccess(true));
                Map map = new HashMap();
                for(String k:variables.keySet()){
                    Object v = variables.get(k);
                    if(v instanceof DataSet){
                        v = ((DataSet)v).getRows();
                    }
                    map.put(k, v);
                }
                data = Ognl.getValue(kk, ognl, map);
            }catch (Exception ignored){
            }
        }
        if(null == data && null != parent){
            data = parent.data(key);
        }
        if(data instanceof String){
            String str = (String)data;
            try {
                if (str.startsWith("{") && str.endsWith("}")) {
                    data = DataRow.parseJson(str);
                }
                if (str.startsWith("[") && str.endsWith("]")) {
                    data = DataSet.parseJson(str);
                }
            }catch (Exception ignored){

            }
        }
        return data;
    }

    public Object data1(String key) {
        if(null == key){
            return null;
        }
        String kk = key.trim();
        Object data = null;
        if(BasicUtil.checkEl(kk)){
            //${users}
            kk = kk.substring(2, kk.length() - 1).trim();
            if(BasicUtil.isNumber(kk)){
                return BasicUtil.parseDecimal(kk, null);
            } else if(BasicUtil.isBoolean(kk)){
                return BasicUtil.parseBoolean(kk, false);
            } else if(kk.startsWith("aov:")){
                return aov(kk);
            }else{
                //多变量顺位 取第一个非空
                if(kk.contains(":")){
                    kk = TagUtil.format(kk);
                    String[] ks = kk.split(":");
                    int len = ks.length;
                    for(int idx = 0; idx <len; idx ++){
                        String k = ks[idx];
                        k = k.trim();
                        if(k.isEmpty()){
                            continue;
                        }
                        if(idx == len -1) {
                            //最后一位默认值
                            //${v1:v2:'abc'}
                            //${v1:v2:123}
                            if (k.startsWith("'") && k.endsWith("'")) {
                                return k.substring(1, k.length() - 1);
                            }
                            if (k.startsWith("\"") && k.endsWith("\"")) {
                                return k.substring(1, k.length() - 1);
                            }
                            if(BasicUtil.isNumber(k)){
                                return BasicUtil.parseInt(k, null);
                            }
                        }
                        Object v = data(k);
                        if(null != v){
                            return v;
                        }
                    }
                }
            }
            data = variables.get(kk);
            if(null == data){
                data = htmls.get(kk);
            }
            if(null == data){
                data = texts.get(kk);
            }
            if(null == data){
                if(kk.contains(".")){
                    //user.dept.name
                    //user[0].name
                    String[] ks = kk.split("\\.");
                    int size = ks.length;
                    if(size > 1) {
                        data = data("${"+ks[0]+"}");
                        for (int i = 1; i < size; i++) {
                            String k = ks[i];
                            if(null == data){
                                break;
                            }

                            if(data instanceof Collection){
                                Collection cols = ((Collection)data);
                                if("size".equals(k)){
                                    data = cols.size();
                                }else if("empty".equals(k)){
                                    data = cols.isEmpty();
                                }
                            } else if(data instanceof Map){
                                Map map = ((Map)data);
                                if("size".equals(k)){
                                    data = map.size();
                                }else if("empty".equals(k)){
                                    data = map.isEmpty();
                                }else{
                                    if(data instanceof DataRow){
                                        data = ((DataRow)data).get(k);
                                    }else {
                                        data = map.get(k);
                                    }
                                }
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
            if(null == data){
                if(kk.contains("+") || kk.contains("-") || kk.contains("*") || kk.contains("/") || kk.contains("%") || kk.contains(">") || kk.contains("<") || kk.contains("=")){
                    try{
                        OgnlContext ognl = new OgnlContext(null, null, new DefaultOgnlMemberAccess(true));
                        data = Ognl.getValue(kk, ognl, variables);
                    }catch (Exception ignored){
                    }
                }
            }
        } else if(kk.startsWith("{") && kk.endsWith("}")) {
            kk = kk.replace("{", "").replace("}", "");
            data = kk;
            if(kk.contains(",")){
                String[] ks = kk.split(",");
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
        if(null == data && null != parent){
            data = parent.data(key);
        }
        return data;
    }

    /**
     * 当前时间 ${aov:now}
     * 随机8位字符${aov:random:8} ${aov:string:8}
     * 随机8位数字${aov:number:8}
     * 随机10-100数字${aov:number:10:100}
     * UUID  ${aov:uuid}
     * @param key key
     * @return Object
     */
    private Object aov(String key){
        //内置常量
        String[] tmps = key.split(":");
        if(tmps.length > 1){
            String var = tmps[1];
            //当前时间
            //ao:now
            if(var.equalsIgnoreCase("now")){
                return new Date();
            }
            //随机字符
            if(var.equalsIgnoreCase("random") || var.equalsIgnoreCase("string")){
                int len = 8;
                if(tmps.length> 2){
                    //随机8位
                    //ao:random:8(默认8位)
                    len = BasicUtil.parseInt(tmps[2], len);
                }
                return BasicUtil.getRandomString(len);
            }
            //随机数字
            if(var.equalsIgnoreCase("number")){
                if(tmps.length> 3){
                    //随机8位
                    //ao:number:0:100
                    int min = BasicUtil.parseInt(tmps[2], 0);
                    int max = BasicUtil.parseInt(tmps[3], 0);
                    return BasicUtil.getRandomNumber(min, max);
                }
                int len = 8;
                if(tmps.length> 2){
                    //随机8位
                    //ao:number:8(默认8位)
                    len = BasicUtil.parseInt(tmps[2], len);
                }
                return BasicUtil.getRandomNumberString(len);
            }
            if(var.equalsIgnoreCase("uuid")){
                return UUID.randomUUID().toString();
            }
        }
        return null;
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
            if(null != value) {
                result = result.replace("${" + key + "}", value);
            }
        }
        for(String key:texts.keySet()){
            String value = texts.get(key);
            if(null != value) {
                result = result.replace("${" + key + "}", value);
            }
        }
        for(String key:variables.keySet()){
            Object value = variables.get(key);
            if(null != value) {
                result = result.replace("${" + key + "}", value.toString());
            }
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
    public Context parent(){
        return parent;
    }
    public void parent(Context parent){
        this.parent = parent;
    }
    public Context clone(){
        Context clone = new Context();
        clone.htmls.putAll(htmls);
        clone.texts.putAll(texts);
        clone.variables.putAll(variables);
        clone.placeholderDefault = placeholderDefault;
        clone.parent = this;
        return clone;
    }
}
