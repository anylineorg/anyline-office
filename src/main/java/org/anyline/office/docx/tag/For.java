package org.anyline.office.docx.tag;

import ognl.Ognl;
import ognl.OgnlContext;
import ognl.OgnlException;
import org.anyline.util.BasicUtil;
import org.anyline.util.DefaultOgnlMemberAccess;
import org.anyline.util.regular.RegularUtil;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

public class For extends AbstractTag implements Tag {
    private Object items;
    private String var;
    private String status;
    private Integer begin;
    private Integer end;
    public String parse(String text) throws Exception {
        StringBuilder html = new StringBuilder();
        String items_key = RegularUtil.fetchAttributeValue(text, "items");
        if(null != items_key) {
            items = data(items_key);
        }
        var = RegularUtil.fetchAttributeValue(text, "var");
        status = RegularUtil.fetchAttributeValue(text, "status");
        begin = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "begin"), 0);
        end = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "end"));
        String body = RegularUtil.fetchTagBody(text, "aol:for", true);
        if(null != items) {
            if (items instanceof Collection) {
                Collection list = (Collection) items;
                int index = 0;
                Map<String, Object> map = new HashMap<>();
                for (Object item : list) {
                    if (null != begin && index < begin) {
                        continue;
                    }
                    if (null != end && index > end) {
                        break;
                    }
                    map.put("index", index);
                    variables.put(var, item);
                    variables.put(status, map);
                    String parse = doc.parseTag(body, variables);
                    parse = placeholder(parse);
                    html.append(parse);
                    index++;
                }
            }
        }else{
            if(null != end){
                Map<String, Object> map = new HashMap<>();
                int index = 0;
                for(int i=begin; i<=end; i++){
                    map.put("index", index);
                    variables.put(var, i);
                    variables.put(status, map);
                    String parse = doc.parseTag(body, variables);
                    parse = placeholder(parse);
                    html.append(parse);
                    index++;
                }
            }
        }
        return html.toString();
    }
}
