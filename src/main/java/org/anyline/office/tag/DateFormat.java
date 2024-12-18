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

import org.anyline.util.BasicUtil;
import org.anyline.util.DateUtil;

import java.util.Date;

public class DateFormat extends AbstractTag implements Tag{
    public void release(){
        super.release();
    }
    @Override
    public void run() {
        String result = null;
        //<aol:date format="yyyy-MM-dd HH:mm:ss" value="${current_time}"></aol:date>

        //空值时 是否取当前时间
        String evl = fetchAttributeString(head, "evl");
        String format = fetchAttributeString(head, "format", "f");
        Date date = null;
        Object data = fetchAttributeData(head, "value");
        if(null == data){
            data = body(text, "date");
        }
        if(BasicUtil.isNotEmpty(data)){
            if("true".equalsIgnoreCase(evl) || "1".equalsIgnoreCase(evl)){
                data = new Date();
            }else {
                data = null;
            }
        }
        if(null != data) {
            date = DateUtil.parse(data);
            result = DateUtil.format(date, format);
        }
        output(result);
    }
}
