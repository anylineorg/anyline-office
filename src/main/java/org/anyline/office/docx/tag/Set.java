package org.anyline.office.docx.tag;

import org.anyline.entity.DataSet;
import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;
import org.anyline.util.regular.RegularUtil;

import java.util.*;

public class Set extends AbstractTag implements Tag{

    private String var;
    private Object data;
    private String selector;
    private Integer index = null;
    private Integer begin = null;
    private Integer end = null;
    private Integer qty = null;
    public String parse(String text){
        String html = "";
        String key = RegularUtil.fetchAttributeValue(text, "data");
        var = RegularUtil.fetchAttributeValue(text, "var");
        selector = RegularUtil.fetchAttributeValue(text, "selector");
        index = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "index"), null);
        begin = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "begin"), null);
        end = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "end"), null);
        qty = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "qty"), null);
        if(BasicUtil.isEmpty(key) || BasicUtil.isEmpty(var)){
            return "";
        }
        data = data(key);
        if (BasicUtil.isNotEmpty(data)) {
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
            }
            doc.variable(var, data);
            doc.replace(var, data.toString());
        }else{
            doc.replace(var, "");
            doc.variable(var, null);
        }
        return html;
    }
}
