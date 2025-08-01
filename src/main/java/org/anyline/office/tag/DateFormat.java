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

package org.anyline.office.tag;

import org.anyline.entity.DataRow;
import org.anyline.entity.DataSet;
import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;
import org.anyline.util.DateUtil;

import java.util.Date;

public class DateFormat extends AbstractTag implements Tag{
    public void release(){
        super.release();
    }
    @Override
    public void run() {
        Object result = null;
        //<aot:date format="yyyy-MM-dd HH:mm:ss" value="${current_time}"></aot:date>

        //空值时 是否取当前时间
        String evl = fetchAttributeString("evl");
        String property = fetchAttributeString("property", "p");
        String format = fetchAttributeString("format", "f");
        Date date = null;
        Object data = fetchAttributeData("value");
        if(null == data){
            data = body(text, "date");
        }
        if(BasicUtil.isEmpty(data)){
            if("true".equalsIgnoreCase(evl) || "1".equalsIgnoreCase(evl)){
                data = new Date();
            }else {
                data = null;
            }
        }
        if(null != data) {
            if(BasicUtil.isNotEmpty(property)){
                //对象属性格式化
                result = data;
                if(data instanceof DataRow){
                    DataRow row = (DataRow)data;
                    row.format.date(format, property);
                } else if(data instanceof DataSet){
                    DataSet set = (DataSet)data;
                    set.format.date(format, property);
                } else {
                    Object v = BeanUtil.getFieldValue(data, property);
                    result = v;
                    if(null != v){
                        Date d = DateUtil.parse(v);
                        result = DateUtil.format(d, format);
                    }
                }
                //TODO 处理 List<Object> Map List<Map>格式
            }else {
                date = DateUtil.parse(data);
                result = DateUtil.format(date, format);
            }
        }
        output(result);
    }
}
