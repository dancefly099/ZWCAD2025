using Aspose.Cells;
using ZwSoft.ZwCAD.ApplicationServices;
using ZwSoft.ZwCAD.DatabaseServices;
using ZwSoft.ZwCAD.EditorInput;
using ZwSoft.ZwCAD.Geometry;
using ZwSoft.ZwCAD.GraphicsInterface;
using IniParser;
using IniParser.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Polyline = ZwSoft.ZwCAD.DatabaseServices.Polyline;

namespace 线图
{
    public static class AddEntityTools
    {


        #region 添加单个实体到模型空间-有doc
        /// <summary>
        /// 添加单个实体到模型空间
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="ent">单个图元</param>
        /// <returns></returns>
        public static ObjectId AddEntityToModelSpaceHasDoc(this Database db, Entity ent)
        {
            Document doc = Application.DocumentManager.GetDocument(db);
            DocumentLock docLock = doc.LockDocument();
            ObjectId entid = ObjectId.Null;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                entid = btr.AppendEntity(ent);
                trans.AddNewlyCreatedDBObject(ent, true);
                db.TransactionManager.QueueForGraphicsFlush();
                trans.Commit();
            }
            docLock.Dispose();
            return entid;

        }
        #endregion

        #region 添加多个实体到模型空间-有doc
        /// <summary>
        /// 添加多个实体到模型空间
        /// </summary>
        /// <param name="db">图元数据库</param>
        /// <param name="ents">图元列表</param>
        /// <returns></returns>
        public static List<ObjectId> AddEntityToModelSpaceHasDoc(this Database db, List<Entity> ents)
        {
            Document doc = Application.DocumentManager.GetDocument(db);
            DocumentLock docLock = doc.LockDocument();
            List<ObjectId> entid = new List<ObjectId>();
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                foreach (Entity ent in ents)
                {
                    entid.Add(btr.AppendEntity(ent));
                    trans.AddNewlyCreatedDBObject(ent, true);
                }

                trans.Commit();
            }
            docLock.Dispose();
            return entid;

        }
        #endregion

        #region 根据图元-EntList删除图元
        public static void DeleteEntityOfEntList(this Database db, List<Entity> entList)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead, true) as BlockTable;
                foreach (Entity bkRefEnt in entList)
                {
                    ObjectId bkObjId = bkRefEnt.ObjectId;
                    Entity ent = bkObjId.GetObject(OpenMode.ForWrite) as Entity;
                    ent.Erase();
                }
                trans.Commit();
            }

        }
        #endregion

        #region 从选择集获取多个实体对象
        /// <summary>
        /// 从选择集获取多个实体对象
        /// </summary>
        /// <param name="mulitSelectResault">选择集结果</param>
        /// <returns></returns>
        public static List<Entity> GetMulitEntityBySelect(this PromptSelectionResult mulitSelectResault)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Entity ent = null;
            List<Entity> entCol = new List<Entity>();
            if (mulitSelectResault.Status == PromptStatus.OK)
            {
                using (Transaction trans = doc.TransactionManager.StartTransaction())
                {
                    SelectionSet ss = mulitSelectResault.Value;
                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        ent = (Entity)trans.GetObject(id, OpenMode.ForRead, true);
                        Type tb = ent.GetType();
                        if (ent != null)
                        {
                            entCol.Add(ent);
                        }
                    }

                }
            }
            return entCol;
        }
        #endregion

        #region 多段线转样条曲线

        public static Spline ToSpline(this Polyline polyline, bool isQuadratic = false)
        {
            using (Polyline2d poly2d = polyline.ConvertTo(false))
            {
                if (isQuadratic)
                    poly2d.ConvertToPolyType(Poly2dType.QuadSplinePoly);

                Spline sp = new Spline();              

                return poly2d.Spline;
            }
        }

        #endregion

        #region 样条曲线转多段线
        public static Polyline ToPolyline2(this Spline sp) {

            return sp.ToPolyline() as Polyline;
        }

        #endregion


    }






}
