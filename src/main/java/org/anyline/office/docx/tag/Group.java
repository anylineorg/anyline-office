package org.anyline.office.docx.tag;

import org.anyline.entity.DataSet;
import org.anyline.util.BasicUtil;
import org.anyline.util.regular.RegularUtil;

public class Group extends AbstractTag implements Tag{
    private String var;
    private Object data;
    private String by;

    public String parse(String text){
        String key = RegularUtil.fetchAttributeValue(text, "data");
        var = RegularUtil.fetchAttributeValue(text, "data");
        by = RegularUtil.fetchAttributeValue(text, "by");
        if(BasicUtil.isEmpty(key) || BasicUtil.isEmpty(var) || BasicUtil.isEmpty(by)){
            return "";
        }
        data = data(key);
        if(data instanceof DataSet){
            DataSet set = (DataSet) data;
            set.group(by.split(","));
            DataSet groups = null;
            variables.put(var, groups);
        }
        return "";
    }
}
