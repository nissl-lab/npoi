using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
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
			Assert.AreEqual("cxn1", cxn.Name);
			Assert.AreEqual(++id, cxn.ID);

			//XSSFGraphicFrame gf = drawing.CreateGraphicFrame(anchor);
			//gf.Name = "graphic frame 1";
			//Assert.AreEqual("graphic frame 1", gf.Name);
			//Assert.AreEqual(++id, gf.ID);

			XSSFShapeGroup sg = drawing.CreateGroup(anchor);
			sg.Name = "shape group 1";
			Assert.AreEqual("shape group 1", sg.Name);
			Assert.AreEqual(++id, sg.ID);

			byte[] jpegData = Encoding.UTF8.GetBytes("test jpeg data");
			int jpegIdx = wb.AddPicture(jpegData, PictureType.JPEG);
			XSSFPicture shape = (XSSFPicture)drawing.CreatePicture(anchor, jpegIdx);
			shape.Name = "name test pic 1";
			Assert.AreEqual("name test pic 1", shape.Name, "ID={0}", shape.ID);
			Assert.AreEqual(++id, shape.ID);

			XSSFSimpleShape ss = drawing.CreateSimpleShape(anchor);
			ss.Name = "simple shape 1";
			Assert.AreEqual("simple shape 1", ss.Name);
			Assert.AreEqual(++id, ss.ID);
		}
	}
}
