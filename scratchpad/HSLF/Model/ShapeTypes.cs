/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSLF.Model;

using NPOI.HSLF.exceptions.HSLFException;




/**
 * Contains all known shape types in PowerPoint
 *
 * @author Yegor Kozlov
 */
public class ShapeTypes : NPOI.sl.usermodel.ShapeTypes {
    /**
     * Return name of the shape by id
     * @param type  - the id of the shape, one of the static constants defined in this class
     * @return  the name of the shape
     */
    public static String typeName(int type) {
        String name = (String)types.Get(Int32.ValueOf(type));
        return name;
    }

    public static HashMap types;
    static {
        types = new HashMap();
        try {
            Field[] f = NPOI.sl.usermodel.ShapeTypes.class.GetFields();
            for (int i = 0; i < f.Length; i++){
                Object val = f[i].Get(null);
                if (val is int) {
                    types.Put(val, f[i].GetName());
                }
            }
        } catch (IllegalAccessException e){
            throw new HSLFException("Failed to Initialize shape types");
        }
    }

}





