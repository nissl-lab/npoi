using NPOI.HSLF.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace NPOI.SL.UserModel
{
	public class BaseBulletStyle : BulletStyle
	{
		private HSLFTextParagraph _textParagraph;
		public BaseBulletStyle(HSLFTextParagraph textParagraph)
		{
			_textParagraph = textParagraph;
		}

		//@Override

		public String GetBulletCharacter()
		{
			Character chr = _textParagraph.GetBulletChar();
			return (chr == null || chr == 0) ? "" : "" + chr;
		}

		//@Override

		public String GetBulletFont()
		{
			return _textParagraph.GetBulletFont();
		}

		//@Override

		public Double GetBulletFontSize()
		{
			return _textParagraph.GetBulletSize();
		}

		//@Override

		public void SetBulletFontColor(Color color)
		{
			_textParagraph.SetBulletFontColor(DrawPaint.createSolidPaint(color));
		}

		//@Override

		public void SetBulletFontColor(PaintStyle color)
		{
			if (!(color is SolidPaint))
			{
				throw new InvalidOperationException("HSLF only supports SolidPaint");
			}
			SolidPaint sp = (SolidPaint)color;
			Color col = DrawPaint.applyColorTransform(sp.getSolidColor());
			_textParagraph.setBulletColor(col);
		}

		//@Override
		public PaintStyle GetBulletFontColor()
		{
			Color col = _textParagraph.GetBulletColor();
			return DrawPaint.createSolidPaint(col);
		}

		//@Override
		public AutoNumberingScheme GetAutoNumberingScheme()
		{
			return _textParagraph.GetAutoNumberingScheme();
		}

		//@Override
		public int GetAutoNumberingStartAt()
		{
			return _textParagraph.GetAutoNumberingStartAt();
		}
	}
}
