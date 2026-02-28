using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.SS.UserModel
{
    /**
     * Represents a type of a conditional formatting rule
     */
    public class ConditionType
    {
        private static Dictionary<int, ConditionType> lookup = new Dictionary<int, ConditionType>();

        /**
         * This conditional formatting rule compares a cell value
         * to a formula calculated result, using an operator
         */
        public static ConditionType CellValueIs = new ConditionType(1, "cellIs");

        /**
         * This conditional formatting rule contains a formula to evaluate.
         * When the formula result is true, the cell is highlighted.
         */
        public static ConditionType Formula = new ConditionType(2, "expression");

        /**
         * This conditional formatting rule contains a color scale,
         * with the cell background set according to a gradient.
         */
        public static ConditionType ColorScale = new ConditionType(3, "colorScale");

        /**
         * This conditional formatting rule sets a data bar, with the
         *  cell populated with bars based on their values
         */
        public static ConditionType DataBar = new ConditionType(4, "dataBar");

        /**
         * This conditional formatting rule that files the values
         */
        public static ConditionType Filter = new ConditionType(5, null);

        /**
         * This conditional formatting rule sets a data bar, with the
         *  cell populated with bars based on their values
         */
        public static ConditionType IconSet = new ConditionType(6, "iconSet");


        public byte Id { get; set; }
        public string Type { get; set; }

        public override string ToString()
        {
            return Id + " - " + Type;
        }

        public override int GetHashCode()
        {
            return Id.GetHashCode() ^ Type.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (obj == null || obj is not ConditionType other)
                return false;
            return this.Id == other.Id && this.Type == other.Type;
        }

        public static ConditionType ForId(byte id)
        {
            return ForId((int)id);
        }
        public static ConditionType ForId(int id)
        {
            return lookup[(id)];
        }

        private ConditionType(int id, String type)
        {
            this.Id = (byte)id; this.Type = type;
            lookup.Add(id, this);
        }
    }
}
