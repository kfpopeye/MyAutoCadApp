/// This app was designed for AutoCAd 2017.
/// It creates a list of all dwg files in the "fixtures" directory on a users desktop then
/// It loads each dwg into autocad and does the following:
///     1) get the block which has the same name as the dwg file
///     2) explode the block and gether the id's of the geometry
///     3) erase the block
///     4) scan the geometry and if any of it does not have "ByLayer" or "Continuous" linetype

using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using System;

namespace MyAutoCadApp
{
    public class Class1
    {
        [CommandMethod("PROCBLOCK")]
        public void ProcessBlock()
        {
            double insPnt_X = 0d;
            double insPnt_Y = 0d;
            Document bDwg = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.DatabaseServices.TransactionManager bTransMan = bDwg.TransactionManager;
            Editor ed = bDwg.Editor;
            string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string[] familyFiles = Directory.GetFiles(path + "\\fixtures", "*.dwg", SearchOption.TopDirectoryOnly);
            ed.WriteMessage("\nUsing directory {0}", path + "\\fixtures");

            try
            {
                DocumentCollection acDocMgr = Application.DocumentManager;
                foreach (string familyFile in familyFiles)
                {
                    string equipmentNumber = Path.GetFileNameWithoutExtension(familyFile);
                    Document doc = acDocMgr.Open(familyFile, false);
                    if (doc != null)
                    {
                        doc.LockDocument();
                        ed.WriteMessage("\nProcessing {0}", equipmentNumber);
                        ExplodeBlockByNameCommand(ed, doc, equipmentNumber);
                        doc.CloseAndSave(familyFile);
                        doc.Dispose();
                    }
                    else
                        ed.WriteMessage("\nCould not open {0}", equipmentNumber);

                    using (bDwg.LockDocument())
                    using (Transaction bTrans = bTransMan.StartTransaction())
                    {
                        Database blkDb = new Database(false, true);
                        blkDb.ReadDwgFile(familyFile, System.IO.FileShare.Read, true, ""); //Reading block source file
                        string name = Path.GetFileNameWithoutExtension(familyFile);
                        ObjectId oid = bDwg.Database.Insert(name, blkDb, true);

                        using (BlockReference acBlkRef = new BlockReference(new Point3d(insPnt_X, insPnt_Y, 0), oid))
                        {
                            BlockTableRecord acCurSpaceBlkTblRec;
                            acCurSpaceBlkTblRec = bTrans.GetObject(bDwg.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                            bTrans.AddNewlyCreatedDBObject(acBlkRef, true);
                        }
                        bTrans.Commit();
                    }
                    insPnt_X += 240d;
                    if (insPnt_X > 2400)
                    {
                        insPnt_X = 0;
                        insPnt_Y += 240d;
                    }
                }
            }
            catch (System.Exception err)
            {
                ed.WriteMessage("\nSomething went wrong in main process.");
                File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\error_main.log", err.ToString());
            }
            finally
            {
            }
        }

        public void ExplodeBlockByNameCommand(Editor ed, Document doc, string blockToExplode)
        {
            Document bDwg = doc;
            Database db = bDwg.Database;
            Database olddb = HostApplicationServices.WorkingDatabase;
            HostApplicationServices.WorkingDatabase = db;
            Autodesk.AutoCAD.DatabaseServices.TransactionManager bTransMan = bDwg.TransactionManager;

            using (Transaction bTrans = bTransMan.StartTransaction())
            {
                try
                {
                    LayerTable lt = (LayerTable)bTrans.GetObject(db.LayerTableId, OpenMode.ForRead); 
                    BlockTable bt = (BlockTable)bTrans.GetObject(db.BlockTableId, OpenMode.ForRead);

                    if (bt.Has(blockToExplode))
                    {
                        ObjectId blkId = bt[blockToExplode];
                        BlockTableRecord btr = (BlockTableRecord)bTrans.GetObject(blkId, OpenMode.ForRead);
                        ObjectIdCollection blkRefs = btr.GetBlockReferenceIds(true, true);

                        foreach (ObjectId blkXId in blkRefs)
                        {
                            //create collection for exploded objects
                            DBObjectCollection objs = new DBObjectCollection();

                            //handle as entity and explode
                            Entity ent = (Entity)bTrans.GetObject(blkXId, OpenMode.ForRead);
                            ent.Explode(objs);

                            //erase Block
                            ent.UpgradeOpen();
                            ent.Erase();

                            BlockTableRecord btrCs = (BlockTableRecord)bTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                            foreach (DBObject obj in objs)
                            {
                                Entity ent2 = (Entity)obj;
                                if(!ent2.Linetype.Equals("ByLayer", StringComparison.CurrentCultureIgnoreCase) &&
                                   !ent2.Linetype.Equals("Continuous", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    string layer = "EQUIPMENT-" + ent2.Linetype;
                                    ObjectId oid;
                                    if (!lt.Has(layer))
                                    {
                                        using (Transaction bTrans2 = bTransMan.StartTransaction())
                                        {
                                            ed.WriteMessage("\nCreating layer {0}", layer);
                                            using (LayerTableRecord ltr = new LayerTableRecord())
                                            {
                                                LayerTable lt2 = (LayerTable)bTrans2.GetObject(db.LayerTableId, OpenMode.ForWrite);
                                                ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                                                ltr.Name = layer;
                                                ltr.LinetypeObjectId = ent2.LinetypeId;
                                                oid = lt2.Add(ltr);
                                                bTrans2.AddNewlyCreatedDBObject(ltr, true);
                                                bTrans2.Commit();
                                            }
                                        }
                                        lt = (LayerTable)bTrans.GetObject(db.LayerTableId, OpenMode.ForRead);
                                    }
                                    else
                                        oid = lt[layer];
                                    ed.WriteMessage("\nSetting entity properties.");
                                    ent2.LayerId = oid;
                                    ent2.Linetype = "ByLayer";
                                }
                                btrCs.AppendEntity(ent2);
                                bTrans.AddNewlyCreatedDBObject(ent2, true);
                            }
                        }

                        //purge block
                        ObjectIdCollection blockIds = new ObjectIdCollection();
                        blockIds.Add(btr.ObjectId);
                        db.Purge(blockIds);
                        btr.UpgradeOpen();
                        btr.Erase();
                    }
                    else
                        ed.WriteMessage("\nCould not find block named {0}", blockToExplode);

                    bTrans.Commit();
                }
                catch (System.Exception err)
                {
                    ed.WriteMessage("\nSomething went wrong: {0}", blockToExplode);
                    File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\error.log", err.ToString());
                }
                finally
                {
                }
                bTrans.Dispose();
                bTransMan.Dispose();
                HostApplicationServices.WorkingDatabase = olddb;
                ed.WriteMessage("\n");
            }
        }

    }
}
