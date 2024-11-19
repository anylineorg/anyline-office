package org.anyline.office.docx.tag;

import org.anyline.util.BasicUtil;
import org.anyline.util.DateUtil;
import org.anyline.util.regular.RegularUtil;

import java.util.Date;

public class DateFormat extends AbstractTag implements Tag{
    @Override
    public String parse(String text) {
        String result = text;
        //<aol:date format="yyyy-MM-dd HH:mm:ss" value="${current_time}"></aol:date>
        String value = RegularUtil.fetchAttributeValue(text, "value");
        //空值时 是否取当前时间
        String evl = RegularUtil.fetchAttributeValue(text, "evl");
        String format = RegularUtil.fetchAttributeValue(text, "format");

        if(BasicUtil.checkEl(format)){
            format = placeholder(format);
        }

        Date date = null;
        if(null == value){
            if("true".equalsIgnoreCase(evl) || "1".equalsIgnoreCase(evl)){
                date = new Date();
            }else {
                return "";
            }
        }
        if(null == date){

        }
        value = value.trim();
        if (BasicUtil.checkEl(value)) {
            //占位符
            value = placeholder(value);
            if(BasicUtil.isNotEmpty(value)) {
                date = DateUtil.parse(value);
            }else{
                String key = value.substring(2, value.length() - 1);
                Object val = variables.get(key);
                if(null != val){
                    date = DateUtil.parse(val);
                }
            }
        }
        if(null == date){
            return "";
        }
        result = DateUtil.format(date, format);
        return result;
    }
}
