using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HWPF.Model;

namespace NPOI.HWPF.UserModel
{
    public partial class Picture
    {
        /**
         * @return Horizontal scaling factor supplied by user expressed in .001%
         *         units
         */
        public int HorizontalScalingFactor
        {
            get { return mx; }
        }
        /**
         * @return Vertical scaling factor supplied by user expressed in .001% units
         */
        public int VerticalScalingFactor
        {
            get { return my; }
        }

        /**
         * Gets the initial width of the picture, in twips, prior to cropping or
         * scaling.
         * 
         * @return the initial width of the picture in twips
         */
        public int DxaGoal
        {
            get
            {
                return dxaGoal;
            }
        }

        /**
         * Gets the initial height of the picture, in twips, prior to cropping or
         * scaling.
         * 
         * @return the initial width of the picture in twips
         */
        public int DyaGoal
        {
            get
            {
                return dyaGoal;
            }
        }

        /**
         * @return The amount the picture has been cropped on the left in twips
         */
        public int DxaCropLeft
        {
            get
            {
                return dxaCropLeft;
            }
        }

        /**
         * @return The amount the picture has been cropped on the top in twips
         */
        public int DyaCropTop
        {
            get
            {
                return dyaCropTop;
            }
        }

        /**
         * @return The amount the picture has been cropped on the right in twips
         */
        public int DxaCropRight
        {
            get
            {
                return dxaCropRight;
            }
        }

        /**
         * @return The amount the picture has been cropped on the bottom in twips
         */
        public int DyaCropBottom
        {
            get
            {
                return dyaCropBottom;
            }
        }

    }
}
