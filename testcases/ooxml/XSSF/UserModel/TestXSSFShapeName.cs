using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;using NUnit.Framework.Legacy;
using System.Text;

namespace TestCases.XSSF.UserModel
{
	[TestFixture]
	internal class TestXSSFShapeName
	{
		[Test]
		public void TestNameAndID()
		{
			int id = 0;

			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
			XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();

			XSSFClientAnchor anchor = new XSSFClientAnchor(0,0,0,0, 1,1,2,2);

			XSSFConnector cxn = drawing.CreateConnector(anchor);
			cxn.Name = "cxn1";
			ClassicAssert.AreEqual("cxn1", cxn.Name);
			ClassicAssert.AreEqual(++id, cxn.ID);

			//XSSFGraphicFrame gf = drawing.CreateGraphicFrame(anchor);
			//gf.Name = "graphic frame 1";
			//ClassicAssert.AreEqual("graphic frame 1", gf.Name);
			//ClassicAssert.AreEqual(++id, gf.ID);

			XSSFShapeGroup sg = drawing.CreateGroup(anchor);
			sg.Name = "shape group 1";
			ClassicAssert.AreEqual("shape group 1", sg.Name);
			ClassicAssert.AreEqual(++id, sg.ID);

			byte[] jpegData = Encoding.UTF8.GetBytes("test jpeg data");
			int jpegIdx = wb.AddPicture(jpegData, PictureType.JPEG);
			XSSFPicture shape = (XSSFPicture)drawing.CreatePicture(anchor, jpegIdx);
			shape.Name = "name test pic 1";
			ClassicAssert.AreEqual("name test pic 1", shape.Name, "ID={0}", shape.ID);
			ClassicAssert.AreEqual(++id, shape.ID);

			XSSFSimpleShape ss = drawing.CreateSimpleShape(anchor);
			ss.Name = "simple shape 1";
			ClassicAssert.AreEqual("simple shape 1", ss.Name);
			ClassicAssert.AreEqual(++id, ss.ID);
		}
	}
}
