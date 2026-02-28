using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.OpenXmlFormats.Shared
{
    public interface IOMathContainer
    {
        ArrayList Items { get; }
        /// <summary>
        /// Add new Run
        /// </summary>
        /// <returns></returns>
        CT_R AddNewR();
        /// <summary>
        /// Add new Accent
        /// </summary>
        /// <returns></returns>
        CT_Acc AddNewAcc();
        /// <summary>
        /// Add new n-ary Operator
        /// </summary>
        /// <returns></returns>
        CT_Nary AddNewNary();
        /// <summary>
        /// Add new Subscript
        /// </summary>
        /// <returns></returns>
        CT_SSub AddNewSSub();
        /// <summary>
        /// Add new Superscript
        /// </summary>
        /// <returns></returns>
        CT_SSup AddNewSSup();
        /// <summary>
        /// Add new Fraction
        /// </summary>
        /// <returns></returns>
        CT_F AddNewF();
        /// <summary>
        /// Add new Radical
        /// </summary>
        /// <returns></returns>
        CT_Rad AddNewRad();
    }
}
