import org.anyline.office.xlsx.entity.XSheet;
import org.anyline.office.xlsx.entity.XWorkBook;
import org.anyline.util.DateUtil;
import org.anyline.util.FileUtil;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class XlsxTest {
    public static void main(String[] args) {
        //replace();
        insert();
    }
    public static void insert(){
        XWorkBook book = book();
        //String内容 会生成ShareString
        book.replace("ymd", DateUtil.format("yyyy/MM/dd"));
        //Number类型直接插入到单元格
        book.replace("price", "200.00");
        XSheet sheet = book.sheet(0);
        List<Object> values = new ArrayList<>();
        values.add("A_APPEND");
        values.add("B_APPEND");
        values.add("C_APPEND");
        sheet.append(values);
        values = new ArrayList<>();
        values.add("A_INSERT");
        values.add("B_INSERT");
        values.add("C_INSERT");
        sheet.insert(4, values);
        book.save();
    }
    public static void replace(){
        XWorkBook book = book();
        //String内容 会生成ShareString
        book.replace("ymd", DateUtil.format("yyyy-MM-dd"));
        //Number类型直接插入到单元格
        book.replace("price", "100.00");
        book.save();
    }
    public static XWorkBook book(){
        //模板中插入占位符${key}
        //模板文件
        File template = new File("E:\\template\\excel\\a.xlsx");
        //复制模板
        File copy = new File(template.getParentFile(), "copy_"+ DateUtil.format("yyyy_MM_dd_HH_mm_ss") +".xlsx");
        FileUtil.copy(template, copy);
        return new XWorkBook(copy);
    }
}
