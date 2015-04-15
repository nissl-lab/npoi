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

namespace NPOI.HSSF.Model
{
    using System;
    using NPOI.DDF;
    using System.Collections;

    /**
     * Provides utilities to manage drawing Groups.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class DrawingManager2
    {
        EscherDggRecord dgg;
        IList drawingGroups = new ArrayList();


        public DrawingManager2(EscherDggRecord dgg)
        {
            this.dgg = dgg;
        }

        /**
         * Clears the cached list of drawing Groups
         */
        public void ClearDrawingGroups()
        {
            drawingGroups.Clear();
        }

        public EscherDgRecord CreateDgRecord()
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

        /**
         * Allocates new shape id for the new drawing Group id.
         *
         * @return a new shape id.
         */
        public int AllocateShapeId(short drawingGroupId)
        {
            EscherDgRecord dg = GetDrawingGroup(drawingGroupId);
            return AllocateShapeId(drawingGroupId, dg);
        }
        /**
 * Allocates new shape id for the new drawing group id.
 *
 * @return a new shape id.
 */
        public int AllocateShapeId(short drawingGroupId, EscherDgRecord dg)
        {
            dgg.NumShapesSaved=(dgg.NumShapesSaved + 1);

            // Add to existing cluster if space available
            for (int i = 0; i < dgg.FileIdClusters.Length; i++)
            {
                EscherDggRecord.FileIdCluster c = dgg.FileIdClusters[i];
                if (c.DrawingGroupId == drawingGroupId && c.NumShapeIdsUsed != 1024)
                {
                    int result = c.NumShapeIdsUsed + (1024 * (i + 1));
                    c.IncrementShapeId();
                    dg.NumShapes=(dg.NumShapes + 1);
                    dg.LastMSOSPID=(result);
                    if (result >= dgg.ShapeIdMax)
                        dgg.ShapeIdMax=(result + 1);
                    return result;
                }
            }

            // Create new cluster
            dgg.AddCluster(drawingGroupId, 0);
            dgg.FileIdClusters[dgg.FileIdClusters.Length - 1].IncrementShapeId();
            dg.NumShapes=(dg.NumShapes + 1);
            int result2 = (1024 * dgg.FileIdClusters.Length);
            dg.LastMSOSPID = (result2);
            if (result2 >= dgg.ShapeIdMax)
                dgg.ShapeIdMax = (result2 + 1);
            return result2;
        }

        /**
         * Finds the next available (1 based) drawing Group id
         */
        public short FindNewDrawingGroupId()
        {
            short dgId = 1;
            while (DrawingGroupExists(dgId))
                dgId++;
            return dgId;
        }

        EscherDgRecord GetDrawingGroup(int drawingGroupId)
        {
            return (EscherDgRecord)drawingGroups[drawingGroupId - 1];
        }

        bool DrawingGroupExists(short dgId)
        {
            for (int i = 0; i < dgg.FileIdClusters.Length; i++)
            {
                if (dgg.FileIdClusters[i].DrawingGroupId == dgId)
                    return true;
            }
            return false;
        }

        int FindFreeSPIDBlock()
        {
            int max = dgg.ShapeIdMax;
            int next = ((max / 1024) + 1) * 1024;
            return next;
        }

        public EscherDggRecord GetDgg()
        {
            return dgg;
        }

    }
}