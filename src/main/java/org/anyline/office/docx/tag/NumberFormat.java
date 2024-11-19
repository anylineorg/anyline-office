package org.anyline.office.docx.tag;

import org.anyline.util.BasicUtil;
import org.anyline.util.NumberUtil;
import org.anyline.util.regular.RegularUtil;

public class NumberFormat extends AbstractTag implements Tag{
    @Override
    public String parse(String text) {
        String result = text;
        //<aol:number format="###,##0.00" value="${total}"></aol:number>
        String value = RegularUtil.fetchAttributeValue(text, "value");
        String format = RegularUtil.fetchAttributeValue(text, "format");
        if(null == value){
            return "";
        }
        value = value.trim();
        if (BasicUtil.checkEl(value)) {
            String key = value.substring(2, value.length() - 1);
            value = replaces.get(key);
        }
        result = NumberUtil.format(value, format);
        return result;
    }
}
