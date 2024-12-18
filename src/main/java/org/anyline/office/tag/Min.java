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

public class Min extends AbstractTag implements Tag{

    public void release(){
        super.release();
    }
    public void run() throws Exception{
        Object data = data();
        Object result = null;
        if(data instanceof DataSet){
            DataSet set = (DataSet)data;
            DataRow min = set.min(property);
            if (null != min) {
                if(BasicUtil.isEmpty(var)) {
                    //直接输出
                    result = min.getString(property);
                }else{
                    result = min;
                }
            }
        }
        output(result);
    }
}