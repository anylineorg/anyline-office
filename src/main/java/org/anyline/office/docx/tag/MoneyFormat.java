package org.anyline.office.docx.tag;

import org.anyline.util.BasicUtil;
import org.anyline.util.MoneyUtil;
import org.anyline.util.NumberUtil;
import org.anyline.util.regular.RegularUtil;

public class MoneyFormat extends AbstractTag implements Tag{
    @Override
    public String parse(String text) {
        String result = "";
        //<aol:money value="${total}"></aol:money>
        String key = RegularUtil.fetchAttributeValue(text, "value");
        if(BasicUtil.isEmpty(key)){
            return "";
        }
        Object data = data(key.trim());
        if(BasicUtil.isNotEmpty(data)) {
            double d = BasicUtil.parseDouble(data, 0d);
            result = MoneyUtil.format(d);
        }
        return result;
    }
}
