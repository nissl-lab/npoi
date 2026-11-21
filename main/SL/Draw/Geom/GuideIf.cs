using System;

namespace NPOI.SL.Draw.Geom
{
    enum Op
    {
        muldiv, addsub, adddiv, ifelse, val, abs, sqrt, max, min, at2, sin, cos, tan, cat2, sat2, pin, mod
    }

    public class GuideIf : Formula
    {
        public string Name { get; set; }
        public string Fmla { get; set; }
        public override double Evaluate(Context ctx)
        {
            return evaluateGuide(ctx);
        }
        double evaluateGuide(Context ctx)
        {
            Op op;
            if (string.IsNullOrWhiteSpace(Fmla))
                throw new ArgumentException("A geometry formula is required.", nameof(Fmla));
            String[] operands = Fmla.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
            switch(operands[0])
            {
                case "*/":
                    op = Op.muldiv;
                    break;
                case "+-":
                    op = Op.addsub;
                    break;
                case "+/":
                    op = Op.adddiv;
                    break;
                case "?:":
                    op = Op.ifelse;
                    break;
                default:
                    if (!Enum.TryParse(operands[0], out op) || !Enum.IsDefined(typeof(Op), op) ||
                        operands[0] != op.ToString())
                        throw new ArgumentException("Unknown geometry formula operator: " + operands[0], nameof(Fmla));
                    break;
            }

            int operandCount;
            switch (op)
            {
                case Op.val:
                case Op.abs:
                case Op.sqrt:
                    operandCount = 1;
                    break;
                case Op.max:
                case Op.min:
                case Op.at2:
                case Op.sin:
                case Op.cos:
                case Op.tan:
                    operandCount = 2;
                    break;
                default:
                    operandCount = 3;
                    break;
            }
            if (operands.Length != operandCount + 1)
                throw new ArgumentException("Incorrect number of geometry formula operands.", nameof(Fmla));

            double x = (operands.Length > 1) ? ctx.GetValue(operands[1]) : 0;
            double y = (operands.Length > 2) ? ctx.GetValue(operands[2]) : 0;
            double z = (operands.Length > 3) ? ctx.GetValue(operands[3]) : 0;
            switch(op)
            {
                case Op.abs:
                    // Absolute Value Formula
                    return Math.Abs(x);
                case Op.adddiv:
                    // Add Divide Formula
                    return (z == 0) ? 0 : (x + y) / z;
                case Op.addsub:
                    // Add Subtract Formula
                    return (x + y) - z;
                case Op.at2:
                    // ArcTan Formula: "at2 x y" = arctan( y / z ) = value of this guide
                    return RadianToDegree(Math.Atan2(y, x)) * OOXML_DEGREE;
                case Op.cos:
                    // Cosine Formula: "cos x y" = (x * cos( y )) = value of this guide
                    return x * Math.Cos(DegreeToRadian(y / OOXML_DEGREE));
                case Op.cat2:
                    // Cosine ArcTan Formula: "cat2 x y z" = (x * cos(arctan(z / y) )) = value of this guide
                    return x * Math.Cos(Math.Atan2(z, y));
                case Op.ifelse:
                    // If Else Formula: "?: x y z" = if (x > 0), then y = value of this guide,
                    // else z = value of this guide
                    return x > 0 ? y : z;
                case Op.val:
                    // Literal Value Expression
                    return x;
                case Op.max:
                    // Maximum Value Formula
                    return Math.Max(x, y);
                case Op.min:
                    // Minimum Value Formula
                    return Math.Min(x, y);
                case Op.mod:
                    // Modulo Formula: "mod x y z" = sqrt(x^2 + b^2 + c^2) = value of this guide
                    return Math.Sqrt(x*x + y*y + z*z);
                case Op.muldiv:
                    // Multiply Divide Formula
                    return (z == 0) ? 0 : (x * y) / z;
                case Op.pin:
                    // Pin To Formula: "pin x y z" = if (y < x), then x = value of this guide
                    // else if (y > z), then z = value of this guide
                    // else y = value of this guide
                    return Math.Max(x, Math.Min(y, z));
                case Op.sat2:
                    // Sine ArcTan Formula: "sat2 x y z" = (x*sin(arctan(z / y))) = value of this guide
                    return x *  Math.Sin(Math.Atan2(z, y));
                case Op.sin:
                    // Sine Formula: "sin x y" = (x * sin( y )) = value of this guide
                    return x *  Math.Sin(DegreeToRadian(y / OOXML_DEGREE));
                case Op.sqrt:
                    // Square Root Formula: "sqrt x" = sqrt(x) = value of this guide
                    return Math.Sqrt(x);
                case Op.tan:
                    // Tangent Formula: "tan x y" = (x * tan( y )) = value of this guide
                    return x *  Math.Tan(DegreeToRadian(y / OOXML_DEGREE));
                default:
                    throw new InvalidOperationException("Unsupported geometry formula operator: " + op);
            }
        }
    }
}
