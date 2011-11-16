using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HWPF.Model;
using System.Collections.ObjectModel;

namespace NPOI.HWPF.UserModel
{
    public interface Fields
    {
        Field GetFieldByStartOffset(FieldsDocumentPart documentPart, int offset);

        Collection<Field> GetFields(FieldsDocumentPart part);
    }
}
