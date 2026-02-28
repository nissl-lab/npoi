/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

using System;
using System.Collections.Generic;

using NPOI.DDF;

namespace NPOI.HSSF.Model
{
    /// <summary>
    /// Provides utilities to manage drawing Groups.
    /// </summary>
    /// <remarks>
    /// Glen Stampoultzis (glens at apache.org) 
    /// </remarks>
    public class DrawingManager2
    {
        private readonly EscherDggRecord dgg;
        private readonly List<EscherDgRecord> drawingGroups = [];


        public DrawingManager2(EscherDggRecord dgg)
        {
            this.dgg = dgg;
        }

        /// <summary>
        /// Clears the cached list of drawing Groups
        /// </summary>
        public void ClearDrawingGroups()
        {
            drawingGroups.Clear();
        }

        public virtual EscherDgRecord CreateDgRecord()
        {
            EscherDgRecord dg = new EscherDgRecord();
            dg.RecordId = EscherDgRecord.RECORD_ID;
            short dgId = FindNewDrawingGroupId();
            dg.Options=(short)(dgId << 4);
            dg.NumShapes=0;
            dg.LastMSOSPID=(-1);
            drawingGroups.Add(dg);
            dgg.AddCluster(dgId, 0);
            dgg.DrawingsSaved=dgg.DrawingsSaved + 1;
            return dg;
        }
        /// <summary>
        /// Allocates new shape id for the new drawing Group id.
        /// </summary>
        /// <param name="drawingGroupId"></param>
        /// <returns>a new shape id.</returns>
        [Obsolete("deprecated in POI 3.17-beta2, use AllocateShapeId(EscherDgRecord) ")]
        public virtual int AllocateShapeId(short drawingGroupId)
        {
            foreach (EscherDgRecord dg in drawingGroups)
            {
                if (dg.DrawingGroupId == drawingGroupId) 
                {
                    return AllocateShapeId(dg);
                }
            }
            throw new InvalidOperationException("Drawing group id "+drawingGroupId+" doesn't exist.");
        }
        /// <summary>
        /// Allocates new shape id for the new drawing group id.
        /// </summary>
        /// <param name="drawingGroupId"></param>
        /// <param name="dg"></param>
        /// <returns>a new shape id.</returns>
        [Obsolete("deprecated in POI 3.17-beta2, use allocateShapeId(EscherDgRecord) ")]
        public virtual int AllocateShapeId(short drawingGroupId, EscherDgRecord dg)
        {
            return AllocateShapeId(dg);
        }

        /// <summary>
        /// Allocates new shape id for the drawing group
        /// </summary>
        /// <param name="dg">the EscherDgRecord which receives the new shape</param>
        /// <returns>a new shape id.</returns>
        public int AllocateShapeId(EscherDgRecord dg) {
            return dgg.AllocateShapeId(dg, true);
        }
        /// <summary>
        /// Finds the next available (1 based) drawing Group id
        /// </summary>
        /// <returns></returns>
        public short FindNewDrawingGroupId()
        {
            return dgg.FindNewDrawingGroupId();
        }

        public EscherDggRecord Dgg
        {
            get
            {
                return dgg;
            }
        }
        public void IncrementDrawingsSaved()
        {
            dgg.DrawingsSaved = dgg.DrawingsSaved + 1;
        }
    }
}