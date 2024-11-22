package org.anyline.office.docx.tag;

import org.anyline.util.BasicUtil;
import org.anyline.util.NumberUtil;
import org.anyline.util.regular.RegularUtil;

public class NumberFormat extends AbstractTag implements Tag{
    @Override
    public String parse(String text) {
        String result = "";
        //<aol:number format="###,##0.00" value="${total}"></aol:number>
        String key = RegularUtil.fetchAttributeValue(text, "value");
        String format = RegularUtil.fetchAttributeValue(text, "format");
        if(BasicUtil.isEmpty(key) || BasicUtil.isEmpty(format)){
            return "";
        }
        Object data = data(key.trim());
        if(BasicUtil.isNotEmpty(data)) {
            result = NumberUtil.format(data.toString(), format);
        }
        return result;
    }
}
