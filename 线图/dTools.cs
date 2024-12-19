using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ZwSoft.ZwCAD.ApplicationServices;
using ZwSoft.ZwCAD.DatabaseServices;
using ZwSoft.ZwCAD.EditorInput;
using ZwSoft.ZwCAD.Geometry;
using ZwSoft.ZwCAD.Runtime;
using Application = ZwSoft.ZwCAD.ApplicationServices.Application;

namespace 线图
{
    public class dTools
    {

        #region Editor-导出多段线线图数据-ZWTT
        [CommandMethod("ZWTT")]
        public void ZWTT()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\nHello World");
        }
        #endregion


    }


}
