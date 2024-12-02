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

package org.anyline.office.docx.tag;

import org.anyline.entity.DataSet;
import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;
import org.anyline.util.regular.RegularUtil;

public class Concat  extends AbstractTag implements Tag{
	private static final long serialVersionUID = 1L;
	private Object data;
	private String split = ",";
 
	public String parse(String text) {
		String result = "";
		try {
			String key = RegularUtil.fetchAttributeValue(text, "data");
			String property = RegularUtil.fetchAttributeValue(text, "data");
			String var = RegularUtil.fetchAttributeValue(text, "data");
			String distinct = RegularUtil.fetchAttributeValue(text, "distinct");
			String split = RegularUtil.fetchAttributeValue(text, "split");
			if(BasicUtil.isEmpty(key)){
				return "";
			}
			data = context.data(key);

			if(data instanceof String) {
				String[] items = data.toString().split(",");
				result = BeanUtil.concat(items, split);
			}else if(data instanceof DataSet){
				DataSet set = (DataSet) data;
				if(BasicUtil.isNotEmpty(distinct)){
					set = set.distinct(distinct.split(","));
				}
				if(null != property){
					result = set.concat(property, split);
				}

			}
			if(BasicUtil.isNotEmpty(var)){
				doc.context().variable(var, result);
				result = "";
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}

}
