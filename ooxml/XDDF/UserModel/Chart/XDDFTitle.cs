/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

namespace NPOI.XDDF.UserModel.Chart
{
    using NPOI.XDDF.UserModel.Text;
    using NPOI.OpenXmlFormats.Dml.Chart;
    /// <summary>
    /// </summary>
    /// <remarks>
    /// @since 4.0.1
    /// </remarks>
    public class XDDFTitle
    {
        private  CT_Title title;
        private  ITextContainer parent;

        public XDDFTitle(ITextContainer parent, CT_Title title)
        {
            this.parent = parent;
            this.title = title;
        }

        public XDDFTextBody Body
        {
            get
            {
                if(!title.IsSetTx())
                {
                    title.AddNewTx();
                }
                CT_Tx tx = title.tx;
                if(tx.IsSetStrRef())
                {
                    tx.UnsetStrRef();
                }
                if(!tx.IsSetRich())
                {
                    tx.AddNewRich();
                }
                return new XDDFTextBody(parent, tx.rich);
            }
        }

        public void SetText(string text)
        {
            if(!title.IsSetLayout())
            {
                title.AddNewLayout();
            }
            Body.SetText(text);
        }

        public void SetOverlay(bool? overlay)
        {
            if(!overlay.HasValue)
            {
                if(title.IsSetOverlay())
                {
                    title.UnsetOverlay();
                }
            }
            else
            {
                if(title.IsSetOverlay())
                {
                    title.overlay.val = overlay.Value ? 1 : 0;
                }
                else
                {
                    title.AddNewOverlay().val = overlay.Value ? 1 : 0;
                }
            }
        }

    }
}
