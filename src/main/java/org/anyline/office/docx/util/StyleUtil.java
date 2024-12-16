package org.anyline.office.docx.util;

import org.anyline.util.BasicUtil;
import org.anyline.util.DomUtil;
import org.dom4j.Attribute;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class StyleUtil {
    public static Map<String, Integer> fontSizes = new HashMap<String, Integer>() {
        {
            put("初号", 84);
            put("小初", 72);
            put("一号", 52);
            put("小一", 48);
            put("二号", 44);
            put("小二", 36);
            put("三号", 33);
            put("小三", 30);
            put("四号", 28);
            put("小四", 24);
            put("五号", 21);
            put("小五", 18);
            put("六号", 15);
            put("小六", 13);
            put("七号", 11);
            put("八号", 10);
        }
    };
    /**
     * copy的样式复制给src
     * @param src src
     * @param copy 被复制p/w或pPr/wPr
     * @param override 如果样式重复,是否覆盖原来的样式
     */
    public static void copy(Element src, Element copy, boolean override){
        if(null == src || null == copy){
            return;
        }
        String name = src.getName();
        String prName = name+"Pr";
        Element srcPr = src.element(prName);
        if(override){
            src.remove(srcPr);
            srcPr = null;
        }
        Element pr = null;
        String copyName = copy.getName();
        if(copyName.equals(prName)){
            pr = copy;
        }else {
            pr = DomUtil.element(copy, prName);;
        }
        if(null != pr){
            if(null == srcPr) {
                // 如果原来没有pr新创建一个
                Element newPr = pr.createCopy();
                src.elements().add(0, newPr);
            }else{
                List<Element> items = pr.elements();
                List<Element> newItems = new ArrayList<>();
                for(Element item:items){
                    String itemName = item.getName();
                    Element srcItem = srcPr.element(itemName);
                    if(override){
                        srcPr.remove(srcItem);
                        srcItem = null;
                    }
                    if(null == srcItem){
                        // 如果原来没有这个样式条目直接复制一个
                        Element newItem = item.createCopy();
                        newItems.add(newItem);
                    }else{
                        // 如果原来有这个样式条目,在原来基础上复制属性
                        List<Attribute> attributes = item.attributes();
                        for(Attribute attribute:attributes){
                            String attributeName = attribute.getName();
                            String attributeFullName = attributeName;
                            String attributeNamespace = attribute.getNamespacePrefix();
                            if(BasicUtil.isNotEmpty(attributeNamespace)){
                                attributeFullName = attributeNamespace+":"+attributeName;
                            }
                            Attribute srcAttribute = srcItem.attribute(attributeName);
                            if(null == srcAttribute){
                                srcAttribute = srcItem.attribute(attributeFullName);
                            }
                            if(override){
                                if(null != srcAttribute){
                                    srcItem.remove(srcAttribute);
                                    srcAttribute = null;
                                }
                            }
                            if(null == srcAttribute) {
                                srcItem.attributeValue(attributeFullName, attribute.getStringValue());
                            }
                        }
                    }
                }
                srcPr.elements().addAll(newItems);
            }
        }
    }
    public static void copy(Element src, Element copy){
        copy(src, copy, false);
    }

    public static void border(Element border, Map<String, String> styles){
        border(border,"top", styles);
        border(border,"right", styles);
        border(border,"bottom", styles);
        border(border,"left", styles);
        border(border,"insideH", styles);
        border(border,"insideV", styles);
        border(border,"tl2br", styles);
        border(border,"tr2bl", styles);
    }
    public static void border(Element border, String side, Map<String, String> styles){
        Element item = null;
        String width = styles.get("border-"+side+"-width");
        String style = styles.get("border-"+side+"-style");
        String color = styles.get("border-"+side+"-color");
        int dxa = DocxUtil.dxa(width);
        int line = ((int)(DocxUtil.dxa2pt(dxa)*8)/4*4);
        if(BasicUtil.isNotEmpty(width)){
            item = DocxUtil.element(border, side);
            item.addAttribute("w:sz", line+"");
            item.addAttribute("w:val", style);
            item.addAttribute("w:color", color);
        }
    }
    public static void padding(Element margin, Map<String, String> styles){
        padding(margin,"top", styles);
        padding(margin,"start", styles);
        padding(margin,"bottom", styles);
        padding(margin,"end", styles);

    }
    public static void padding(Element margin, String side, Map<String, String> styles){
        String width = styles.get("padding-"+side);
        int dxa = DocxUtil.dxa(width);
        if(BasicUtil.isNotEmpty(width)){
            Element item = DocxUtil.element(margin, side);
            item.addAttribute("w:w", dxa+"");
            item.addAttribute("w:type",  "dxa");
        }
    }
    public static int fontSize(String size){
        int pt = 0;
        if(fontSizes.containsKey(size)){
            pt = fontSizes.get(size);
        }else{
            if(size.endsWith("px")){
                int px = BasicUtil.parseInt(size.replace("px",""),0);
                pt = (int)DocxUtil.px2pt(px);
            }else if(size.endsWith("pt")){
                pt = BasicUtil.parseInt(size.replace("pt",""),0);
            }
        }
        return pt;
    }
    public static void font(Element pr, Map<String, String> styles){
        String fontSize = styles.get("font-size");
        if(null != fontSize){
            int pt = 0;
            if(fontSizes.containsKey(fontSize)){
                pt = fontSizes.get(fontSize);
            }else{
                if(fontSize.endsWith("px")){
                    int px = BasicUtil.parseInt(fontSize.replace("px",""),0);
                    pt = (int)DocxUtil.px2pt(px);
                }else if(fontSize.endsWith("pt")){
                    pt = BasicUtil.parseInt(fontSize.replace("pt",""),0);
                }
            }
            if(pt>0){
                // <w:sz w:val="28"/>
                DocxUtil.element(pr, "sz","val", pt+"");
            }
        }
        // 加粗
        String fontWeight = styles.get("font-weight");
        if(null != fontWeight && fontWeight.length()>0){
            int weight = BasicUtil.parseInt(fontWeight,0);
            if(weight >=700){
                // <w:b w:val="true"/>
                DocxUtil.element(pr, "b","val","true");
            }
        }
        // 下划线
        String underline = styles.get("underline");
        if(null != underline){
            if(underline.equalsIgnoreCase("true") || underline.equalsIgnoreCase("single")){
                // <w:u w:val="single"/>
                DocxUtil.element(pr, "u","val","single");
            }else{
                DocxUtil.element(pr, "u","val",underline);
                /*dash - a dashed line
                dashDotDotHeavy - a series of thick dash, dot, dot characters
                dashDotHeavy - a series of thick dash, dot characters
                dashedHeavy - a series of thick dashes
                dashLong - a series of long dashed characters
                dashLongHeavy - a series of thick, long, dashed characters
                dotDash - a series of dash, dot characters
                dotDotDash - a series of dash, dot, dot characters
                dotted - a series of dot characters
                dottedHeavy - a series of thick dot characters
                double - two lines
                none - no underline
                single - a single line
                thick - a single think line
                wave - a single wavy line
                wavyDouble - a pair of wavy lines
                wavyHeavy - a single thick wavy line
                words - a single line beneath all non-space characters
                */
            }
        }
        // 删除线
        String strike = styles.get("strike");
        if(null != strike){
            if(strike.equalsIgnoreCase("true")){
                // <w:dstrike w:val="true"/>
                DocxUtil.element(pr, "dstrike","val","true");
            }else if("none".equalsIgnoreCase(strike) || "false".equalsIgnoreCase(strike)){
                DocxUtil.element(pr, "dstrike","val","false");
            }
        }
        // 斜体
        String italics = styles.get("italic");
        if(null != italics){
            if(italics.equalsIgnoreCase("true")){
                // <w:dstrike w:val="true"/>
                DocxUtil.element(pr, "i","val","true");
            }else if("none".equalsIgnoreCase(italics) || "false".equalsIgnoreCase(italics)){
                DocxUtil.element(pr, "i","val","false");
            }
        }
        String fontFamily = styles.get("font-family");
        if(null != fontFamily){
            DocxUtil.element(pr, "rFonts","eastAsia",fontFamily);
        }
        String fontFamilyAscii = styles.get("font-family-ascii");
        if(null != fontFamilyAscii){
            DocxUtil.element(pr, "rFonts","ascii",fontFamilyAscii);
        }
        String fontFamilyEast = styles.get("font-family-east");
        if(null != fontFamilyEast){
            DocxUtil.element(pr, "rFonts","eastAsia",fontFamilyEast);
        }
        fontFamilyEast = styles.get("font-family-eastAsia");
        if(null != fontFamilyEast){
            DocxUtil.element(pr, "rFonts","eastAsia",fontFamilyEast);
        }
        String fontFamilyhAnsi = styles.get("font-family-height");
        if(null != fontFamilyhAnsi){
            DocxUtil.element(pr, "rFonts","hAnsi",fontFamilyhAnsi);
        }
        fontFamilyhAnsi = styles.get("font-family-hAnsi");
        if(null != fontFamilyhAnsi){
            DocxUtil.element(pr, "rFonts","hAnsi",fontFamilyhAnsi);
        }
        String fontFamilyComplex = styles.get("font-family-complex");
        if(null != fontFamilyComplex){
            DocxUtil.element(pr, "rFonts","cs",fontFamilyComplex);
        }
        fontFamilyComplex = styles.get("font-family-cs");
        if(null != fontFamilyComplex){
            DocxUtil.element(pr, "rFonts","cs",fontFamilyComplex);
        }

        String fontFamilyHint = styles.get("font-family-hint");
        if(null != fontFamilyHint){
            DocxUtil.element(pr, "rFonts","hint",fontFamilyHint);
        }
        // <w:rFonts w:ascii="Adobe Gothic Std B" w:eastAsia="宋体" w:hAnsi="宋体" w:cs="宋体" w:hint="eastAsia"/>
    }

    public static void background(Element pr,Map<String, String> styles){
        String color = styles.get("background-color");
        if(null != color){
            // <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>
            DocxUtil.element(pr, "shd", "color","auto");
            DocxUtil.element(pr, "shd", "val","clear");
            DocxUtil.element(pr, "shd", "fill",color.replace("#",""));
        }
    }
}
