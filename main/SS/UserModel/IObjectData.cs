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

namespace NPOI.SS.UserModel
{
    using NPOI.POIFS.FileSystem;

    /// <summary>
    /// <para>
    /// Common interface for OLE shapes, i.e. shapes linked to embedded documents
    /// </para>
    /// <para>
    /// </para>
    /// </summary>
    /// @since POI 3.16-beta2
    public interface IObjectData : ISimpleShape
    {
        /// <summary>
        /// data portion, for an ObjectData that doesn't have an associated POIFS Directory Entry
        /// </summary>
        byte[] ObjectData { get; }


        /// <summary>does this ObjectData have an associated POIFS Directory Entry?
        ///(Not all do, those that don't have a data portion)
        /// </summary>
        bool HasDirectoryEntry();

        /// <summary>
        /// <para>
        /// Gets the object data. Only call for ones that have
        /// data though. See {@link #hasDirectoryEntry()}.
        /// The caller has to close the corresponding POIFSFileSystem
        /// </para>
        /// </summary>
        /// <return>object data as an OLE2 directory./// </return>
        /// <throws name="IOException">if there was an error Reading the data. </throws>
        DirectoryEntry Directory { get; }


        /// <summary>
        /// </summary>
        /// <return>the OLE2 Class Name of the object/// </return>
        string OLE2ClassName { get; }

        /// <summary>
        /// </summary>
        /// <return>a filename suggestion - inspecting/interpreting the Directory object probably gives a better result</return>
        string FileName { get; }

        /// <summary>
        /// </summary>
        /// <return>the preview picture
        /// </return>
        IPictureData PictureData { get; }
    }
}
