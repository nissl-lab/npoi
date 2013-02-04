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
    using NPOI.DDF;
    using System.Collections;


    /**
     * Provides utilities to manage drawing Groups.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class DrawingManager
    {
        EscherDggRecord dgg;
        Hashtable dgMap = new Hashtable(); // key = Short(drawingId), value=EscherDgRecord

        public DrawingManager(EscherDggRecord dgg)
        {
            this.dgg = dgg;
        }

        public EscherDgRecord CreateDgRecord()
        {
            EscherDgRecord dg = new EscherDgRecord();
            dg.RecordId=(EscherDgRecord.RECORD_ID);
            short dgId = FindNewDrawingGroupId();
            dg.Options=((short)(dgId << 4));
            dg.NumShapes=(0);
            dg.LastMSOSPID=(-1);
            dgg.AddCluster(dgId, 0);
            dgg.DrawingsSaved=(dgg.DrawingsSaved + 1);
            dgMap[dgId]= dg;
            return dg;
        }

        /**
         * Allocates new shape id for the new drawing Group id.
         *
         * @return a new shape id.
         */
        
        public int AllocateShapeId(short drawingGroupId)
        {
            // Get the last shape id for this drawing Group.
            EscherDgRecord dg = (EscherDgRecord)dgMap[drawingGroupId];
            int lastShapeId = dg.LastMSOSPID;


            // Have we run out of shapes for this cluster?
            int newShapeId = 0;
            if (lastShapeId % 1024 == 1023)
            {
                // Yes:
                // Find the starting shape id of the next free cluster
                newShapeId = FindFreeSPIDBlock();
                // Create a new cluster in the dgg record.
                dgg.AddCluster(drawingGroupId, 1);
            }
            else
            {
                // No:
                // Find the cluster for this drawing Group with free space.
                for (int i = 0; i < dgg.FileIdClusters.Length; i++)
                {
                    EscherDggRecord.FileIdCluster c = dgg.FileIdClusters[i];
                    if (c.DrawingGroupId == drawingGroupId)
                    {
                        if (c.NumShapeIdsUsed != 1024)
                        {
                            // Increment the number of shapes used for this cluster.
                            c.IncrementShapeId();
                        }
                    }
                    // If the last shape id = -1 then we know to Find a free block;
                    if (dg.LastMSOSPID == -1)
                    {
                        newShapeId = FindFreeSPIDBlock();
                    }
                    else
                    {
                        // The new shape id to be the last shapeid of this cluster + 1
                        newShapeId = dg.LastMSOSPID + 1;
                    }
                }
            }
            // Increment the total number of shapes used in the dgg.
            dgg.NumShapesSaved=(dgg.NumShapesSaved + 1);
            // Is the new shape id >= max shape id for dgg?
            if (newShapeId >= dgg.ShapeIdMax)
            {
                // Yes:
                // Set the max shape id = new shape id + 1
                dgg.ShapeIdMax=newShapeId + 1;
            }
            // Set last shape id for this drawing Group.
            dg.LastMSOSPID=newShapeId;
            // Increased the number of shapes used for this drawing Group.
            dg.IncrementShapeCount();


            return newShapeId;
        }

        
        public short FindNewDrawingGroupId()
        {
            short dgId = 1;
            while (DrawingGroupExists(dgId))
                dgId++;
            return dgId;
        }

        public bool DrawingGroupExists(short dgId)
        {
            for (int i = 0; i < dgg.FileIdClusters.Length; i++)
            {
                if (dgg.FileIdClusters[i].DrawingGroupId == dgId)
                    return true;
            }
            return false;
        }

        public int FindFreeSPIDBlock()
        {
            int max = dgg.ShapeIdMax;
            int next = ((max / 1024) + 1) * 1024;
            return next;
        }

        public EscherDggRecord Dgg
        {
            get
            {
                return dgg;
            }
        }

    }
}