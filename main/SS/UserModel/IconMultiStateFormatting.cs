using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.SS.UserModel
{
    public interface IIconMultiStateFormatting
    {

        /// <summary>
        ///Get or set the Icon Set used
        /// </summary>
        IconSet IconSet { get; set; }

        /// <summary>
        /// <para>
        /// Changes the Icon Set used
        /// </para>
        /// <para>
        /// If the new Icon Set has a different number of
        ///  icons to the old one, you <em>must</em> update the
        ///  thresholds before saving!
        /// </para>
        /// </summary>

        /// <summary>
        /// Should Icon + Value be displayed, or only the Icon?
        /// </summary>
        bool IsIconOnly { get; set; }

        bool IsReversed { get; set; }
        /// <summary>
        /// Gets the list of thresholds
        /// </summary>
        /// <summary>
        /// Sets the of thresholds. The number must match
        ///  <see cref="IconSet.num" /> for the current <see cref="getIconSet()" />
        /// </summary>

        IConditionalFormattingThreshold[] Thresholds { get; set; }
        /// <summary>
        /// Creates a new, empty Threshold
        /// </summary>
        IConditionalFormattingThreshold CreateThreshold();
    }

    public class IconSet
    {
        /// <summary>
        /// Green Up / Yellow Side / Red Down arrows */
        /// </summary>
        public static IconSet GYR_3_ARROW = new IconSet(0, 3, "3Arrows");
        /// <summary>
        /// Grey Up / Side / Down arrows */
        /// </summary>
        public static IconSet GREY_3_ARROWS = new IconSet(1, 3, "3ArrowsGray");
        /// <summary>
        /// Green / Yellow / Red flags */
        /// </summary>
        public static IconSet GYR_3_FLAGS = new IconSet(2, 3, "3Flags");
        /// <summary>
        /// Green / Yellow / Red traffic lights (no background). Default */
        /// </summary>
        public static IconSet GYR_3_TRAFFIC_LIGHTS = new IconSet(3, 3, "3TrafficLights1");
        /// <summary>
        /// Green / Yellow / Red traffic lights on a black square background.
        /// Note, MS-XLS docs v20141018 say this is id=5 but seems to be id=4 */
        /// </summary>
        public static IconSet GYR_3_TRAFFIC_LIGHTS_BOX = new IconSet(4, 3, "3TrafficLights2");
        /// <summary>
        /// Green Circle / Yellow Triangle / Red Diamond.
        /// Note, MS-XLS docs v20141018 say this is id=4 but seems to be id=5 */
        /// </summary>
        public static IconSet GYR_3_SHAPES = new IconSet(5, 3, "3Signs");
        /// <summary>
        /// Green Tick / Yellow ! / Red Cross on a circle background */
        /// </summary>
        public static IconSet GYR_3_SYMBOLS_CIRCLE = new IconSet(6, 3, "3Symbols");
        /// <summary>
        /// Green Tick / Yellow ! / Red Cross (no background) */
        /// </summary>
        public static IconSet GYR_3_SYMBOLS = new IconSet(7, 3, "3Symbols2");
        /// <summary>
        /// Green Up / Yellow NE / Yellow SE / Red Down arrows */
        /// </summary>
        public static IconSet GYR_4_ARROWS = new IconSet(8, 4, "4Arrows");
        /// <summary>
        /// Grey Up / NE / SE / Down arrows */
        /// </summary>
        public static IconSet GREY_4_ARROWS = new IconSet(9, 4, "4ArrowsGray");
        /// <summary>
        /// Red / Light Red / Grey / Black traffic lights */
        /// </summary>
        public static IconSet RB_4_TRAFFIC_LIGHTS = new IconSet(0xA, 4, "4RedToBlack");
        public static IconSet RATINGS_4 = new IconSet(0xB, 4, "4Rating");
        /// <summary>
        /// Green / Yellow / Red / Black traffic lights */
        /// </summary>
        public static IconSet GYRB_4_TRAFFIC_LIGHTS = new IconSet(0xC, 4, "4TrafficLights");
        public static IconSet GYYYR_5_ARROWS = new IconSet(0xD, 5, "5Arrows");
        public static IconSet GREY_5_ARROWS = new IconSet(0xE, 5, "5ArrowsGray");
        public static IconSet RATINGS_5 = new IconSet(0xF, 5, "5Rating");
        public static IconSet QUARTERS_5 = new IconSet(0x10, 5, "5Quarters");


        protected static IconSet DEFAULT_ICONSET = IconSet.GYR_3_TRAFFIC_LIGHTS;

        /// <summary>
        /// Numeric ID of the icon set */
        /// </summary>
        public int id;
        /// <summary>
        /// How many icons in the set */
        /// </summary>
        public int num;
        /// <summary>
        /// Name (system) of the set */
        /// </summary>
        public String name;

        private static List<IconSet> values = new List<IconSet>() {
            GYR_3_ARROW, GREY_3_ARROWS, GYR_3_FLAGS, GYR_3_TRAFFIC_LIGHTS, GYR_3_TRAFFIC_LIGHTS_BOX,
            GYR_3_SHAPES, GYR_3_SYMBOLS_CIRCLE, GYR_3_SYMBOLS, GYR_4_ARROWS, GREY_4_ARROWS,
            RB_4_TRAFFIC_LIGHTS, RATINGS_4, GYRB_4_TRAFFIC_LIGHTS, GYYYR_5_ARROWS, GREY_5_ARROWS,
            RATINGS_5, QUARTERS_5
        };
        public static List<IconSet> Values()
        {
            return values;
        }
        public override String ToString()
        {
            return id + " - " + name;
        }

        public static IconSet ById(int id)
        {
            return Values()[id];
        }
        public static IconSet ByName(String name)
        {
            foreach(IconSet set in Values())
            {
                if(set.name.Equals(name))
                    return set;
            }
            return null;
        }
        public static IconSet ByOOXMLName(String name)
        {
            if(name.StartsWith("Item"))
                name = name.Remove(0, 4);
            return ByName(name);
        }
        private IconSet(int id, int num, String name)
        {
            this.id = id;
            this.num = num;
            this.name = name;
        }
    }

}
