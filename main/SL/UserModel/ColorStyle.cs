using SixLabors.ImageSharp;

namespace NPOI.SL.UserModel
{
    public interface ColorStyle
    {
        Color GetColor();
        int Alpha { get; }
        int HueOff { get; }
        int HueMod { get; }
        int SatOff { get; }
        int SatMod { get; }
        int LumOff { get; }
        int LumMod { get; }
        int Shade { get; }
        int Tint { get; }
    }
}