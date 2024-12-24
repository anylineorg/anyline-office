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

import org.anyline.entity.DataSet;

public class Concat  extends AbstractTag implements Tag{
	private static final long serialVersionUID = 1L;
	private String split = ",";

	public void release(){
		super.release();
		data = null;
		split = ",";
	}
	public void run() {
		String result = "";
		try {
			String split = fetchAttributeString("split", "s");
			if(null == split){
				split = ",";
			}
			data = data();
			if(data instanceof DataSet){
				DataSet set = (DataSet) data;
				if(null != property){
					result = set.concat(property, split);
				}
			}
			output(result);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
