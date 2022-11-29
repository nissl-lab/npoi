using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace NPOI.SL.UserModel
{
	public class GenericBulletStyle : BulletStyle
	{
		private Func<AutoNumberingScheme> autoNumberingScheme;
		private Func<int> autoNumberingStartAt;
		private Func<int> autoNumberingEndAt;
		private Func<string> bulletCharacter;
		private Func<string> bulletFont;
		private Func<PaintStyle> bulletFontColor;
		private Func<double> bulletFontSize;

		public GenericBulletStyle(Func<AutoNumberingScheme> autoNumberingScheme, Func<int> autoNumberingStartAt, Func<int> autoNumberingEndAt, Func<string> bulletCharacter, Func<string> bulletFont, Func<PaintStyle> bulletFontColor, Func<double> bulletFontSize)
		{
			this.autoNumberingScheme = autoNumberingScheme;
			this.autoNumberingStartAt = autoNumberingStartAt;
			this.autoNumberingEndAt = autoNumberingEndAt;
			this.bulletCharacter = bulletCharacter;
			this.bulletFont = bulletFont;
			this.bulletFontColor = bulletFontColor;
			this.bulletFontSize = bulletFontSize;	
		}

		public AutoNumberingScheme getAutoNumberingScheme()
		{
			return this.autoNumberingScheme();
		}

		public int getAutoNumberingStartAt()
		{
			return this.autoNumberingStartAt();
		}

		public string getBulletCharacter()
		{
			return this.bulletCharacter();
		}

		public string getBulletFont()
		{
			return this.bulletFont();
		}

		public PaintStyle getBulletFontColor()
		{
			return this.bulletFontColor();
		}

		public double getBulletFontSize()
		{
			return this.bulletFontSize();
		}

		public void setBulletFontColor(Color color)
		{
			throw new NotImplementedException();
		}

		public void setBulletFontColor(PaintStyle color)
		{
			throw new NotImplementedException();
		}
	}
}
