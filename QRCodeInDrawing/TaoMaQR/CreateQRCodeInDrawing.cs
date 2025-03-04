using QRCoder;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Teigha.ApplicationServices;
using Teigha.DatabaseServices;
using Teigha.EditorInput;
using Teigha.Geometry;
using Teigha.Runtime;

namespace QRCodeInDrawing.TaoMaQR
{
    public class CreateQRCodeInDrawing
    {
        [CommandMethod("InsertQRCode")]
        public void InsertQRCode()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database dat = doc.Database;

            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                #region bo loc SeihinBango
                TypedValue[] TypValArrSeihinbango = new TypedValue[2];
                TypValArrSeihinbango.SetValue(new TypedValue((int)DxfCode.Start, "TEXT,MTEXT"), 0);//loc doi tuong text
                TypValArrSeihinbango.SetValue(new TypedValue((int)DxfCode.LayerName, "256_TKK_Waku_meisyou"), 1);//loc doi tuong co layer la 256_TKK_Waku_meisyou
                #endregion

                #region bo loc Shuzai
                TypedValue[] TypValArrShuzai = new TypedValue[2];
                TypValArrShuzai.SetValue(new TypedValue((int)DxfCode.Start, "TEXT,MTEXT"), 0);//loc doi tuong text
                TypValArrShuzai.SetValue(new TypedValue((int)DxfCode.LayerName, "257_TKK_Waku_syuzai"), 1);//loc doi tuong co layer la 257_TKK_Waku_syuzai
                #endregion

                #region bo loc Shiage
                TypedValue[] TypValArrShiage = new TypedValue[2];
                TypValArrShiage.SetValue(new TypedValue((int)DxfCode.Start, "TEXT,MTEXT"), 0);//loc doi tuong text
                TypValArrShiage.SetValue(new TypedValue((int)DxfCode.LayerName, "258_TKK_Waku_shiage"), 1);//loc doi tuong co layer la 258_TKK_Waku_shiage
                #endregion

                #region bo loc SuuRyo
                TypedValue[] TypValArrSuuRyo = new TypedValue[2];
                TypValArrSuuRyo.SetValue(new TypedValue((int)DxfCode.Start, "TEXT,MTEXT"), 0);//loc doi tuong text
                TypValArrSuuRyo.SetValue(new TypedValue((int)DxfCode.LayerName, "259_TKK_Waku_suuryou"), 1);//loc doi tuong co layer la 259_TKK_Waku_suuryou
                #endregion

                #region bo loc ZumenBango
                TypedValue[] TypValArrZumenBango = new TypedValue[2];
                TypValArrZumenBango.SetValue(new TypedValue((int)DxfCode.Start, "TEXT,MTEXT"), 0);//loc doi tuong text
                TypValArrZumenBango.SetValue(new TypedValue((int)DxfCode.LayerName, "255_TKK_Waku_zumenbangou"), 1);//loc doi tuong co layer la 255_TKK_Waku_zumenbangou
                #endregion

                //tao bo loc
                SelectionFilter SelFilSehinBango = new SelectionFilter(TypValArrSeihinbango);
                SelectionFilter SelFilShuzai = new SelectionFilter(TypValArrShuzai);
                SelectionFilter SelFilShiage = new SelectionFilter(TypValArrShiage);
                SelectionFilter SelFilSuuRyo = new SelectionFilter(TypValArrSuuRyo);
                SelectionFilter SelFilZumenBango=new SelectionFilter(TypValArrZumenBango);

                #region List thong tin
                List<Point3d> minPtArr = new List<Point3d>();
                List<Point3d> maxPtArr = new List<Point3d>();
                List<string> SeihinBangoArr = new List<string>();
                List<string> ShuzaiArr = new List<string>();
                List<string> ShiageArr = new List<string>();
                List<string> SuuRyoArr = new List<string>();
                List<string> ZumenBangoArr = new List<string>();
                List<double> blockLenArr= new List<double>();
                List<double> blockHeightArr = new List<double>();
                #endregion

                #region Getblock khung ban ve va Min/MaxPoint
                var ids = GetBlockId();
                List<BlockReference> BlockColl = new List<BlockReference>();
                //Line line1 = new Line();
                var sortedBlock = new List<BlockReference>();
                if (ids != null)
                {
                    using (Transaction tr2 = dat.TransactionManager.StartTransaction())
                    {
                        foreach (ObjectId id in ids)
                        {
                            BlockReference blRef = tr2.GetObject(id, OpenMode.ForRead) as BlockReference;
                            if (blRef != null)
                            {
                                BlockColl.Add(blRef);
                            }
                        }
                        //BlockColl.Sort((a,b)=>a.Position.X.CompareTo(b.Position.X));
                        sortedBlock = BlockColl.OrderBy(block => block.Position.X)
                                               .ThenByDescending(block => block.Position.Y)
                                               .ToList();
                    }
                }
                //lay list minpoint, maxpoint cua cac block chua thong tin
                for (int iChay = 0; iChay < sortedBlock.Count; iChay += 1)
                {
                    minPtArr.Add(sortedBlock[iChay].GeometricExtents.MinPoint);
                    maxPtArr.Add(sortedBlock[iChay].GeometricExtents.MaxPoint);
                    double deltaX = sortedBlock[iChay].GeometricExtents.MaxPoint.X - sortedBlock[iChay].GeometricExtents.MinPoint.X;
                    double deltaY = sortedBlock[iChay].GeometricExtents.MaxPoint.Y - sortedBlock[iChay].GeometricExtents.MinPoint.Y;
                    blockLenArr.Add(deltaX);
                    blockHeightArr.Add(deltaY);
                }
                #endregion

                #region tao thu muc QR Picture trong o C va xoa het cac anh dang luu trong do
                // Create the directory if it doesn't exist
                string directoryPath = @"C:\QR Picture";
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }
                // Delete existing files in the directory
                var existingFiles = Directory.GetFiles(directoryPath, "*.png");
                foreach (var file in existingFiles)
                {
                    File.Delete(file);
                }
                #endregion

                for (int iChay = 0; iChay < sortedBlock.Count; iChay += 1)
                {
                    #region lay list SeihinBango
                    PromptSelectionResult promSeRes2 = doc.Editor.SelectCrossingWindow(minPtArr[iChay], maxPtArr[iChay], SelFilSehinBango);
                    if (promSeRes2.Status == PromptStatus.OK)
                    {
                        SelectionSet Selset2 = promSeRes2.Value;
                        foreach (ObjectId objId2 in Selset2.GetObjectIds())
                        {
                            //xet 2 truong hop text va mtext
                            DBText text2 = tr.GetObject(objId2, OpenMode.ForRead) as DBText;
                            if (text2 != null)
                            {
                                SeihinBangoArr.Add(text2.TextString);
                            }

                            MText text2a = tr.GetObject(objId2, OpenMode.ForRead) as MText;
                            if (text2a != null)
                            {
                                SeihinBangoArr.Add(text2a.Text);
                            }
                        }
                    }
                    else
                    {
                        SeihinBangoArr.Add("00000");
                    }
                    #endregion

                    #region lay list Shuzai
                    PromptSelectionResult promSeRes3 = doc.Editor.SelectCrossingWindow(minPtArr[iChay], maxPtArr[iChay], SelFilShuzai);
                    if (promSeRes3.Status == PromptStatus.OK)
                    {
                        SelectionSet Selset3 = promSeRes3.Value;
                        foreach (ObjectId objId3 in Selset3.GetObjectIds())
                        {
                            //xet 2 truong hop text va mtext
                            DBText text3 = tr.GetObject(objId3, OpenMode.ForRead) as DBText;
                            if (text3 != null)
                            {
                                ShuzaiArr.Add(text3.TextString);
                            }

                            MText text3a = tr.GetObject(objId3, OpenMode.ForRead) as MText;
                            if (text3a != null)
                            {
                                ShuzaiArr.Add(text3a.Text);
                            }
                        }
                    }
                    else
                    {
                        ShuzaiArr.Add("00000");
                    }
                    #endregion

                    #region lay list Shiage
                    PromptSelectionResult promSeRes4 = doc.Editor.SelectCrossingWindow(minPtArr[iChay], maxPtArr[iChay], SelFilShiage);
                    if (promSeRes4.Status == PromptStatus.OK)
                    {
                        SelectionSet Selset4 = promSeRes4.Value;
                        foreach (ObjectId objId4 in Selset4.GetObjectIds())
                        {
                            //xet 2 truong hop text va mtext
                            DBText text4 = tr.GetObject(objId4, OpenMode.ForRead) as DBText;
                            if (text4 != null)
                            {
                                ShiageArr.Add(text4.TextString);
                            }

                            MText text4a = tr.GetObject(objId4, OpenMode.ForRead) as MText;
                            if (text4a != null)
                            {
                                ShiageArr.Add(text4a.Text);
                            }
                        }
                    }
                    else
                    {
                        ShiageArr.Add("00000");
                    }
                    #endregion

                    #region lay list SuuRyo
                    PromptSelectionResult promSeRes5 = doc.Editor.SelectCrossingWindow(minPtArr[iChay], maxPtArr[iChay], SelFilSuuRyo);
                    if (promSeRes5.Status == PromptStatus.OK)
                    {
                        SelectionSet Selset5 = promSeRes5.Value;
                        foreach (ObjectId objId5 in Selset5.GetObjectIds())
                        {
                            //xet 2 truong hop text va mtext
                            DBText text5 = tr.GetObject(objId5, OpenMode.ForRead) as DBText;
                            if (text5 != null)
                            {
                                SuuRyoArr.Add(text5.TextString);
                            }

                            MText text5a = tr.GetObject(objId5, OpenMode.ForRead) as MText;
                            if (text5a != null)
                            {
                                SuuRyoArr.Add(text5a.Text);
                            }
                        }
                    }
                    else
                    {
                        SuuRyoArr.Add("0");
                    }
                    #endregion

                    #region lay list Zumenbango
                    //lay  tu diem minPt cua block xuong mot khoang de tao ra vung quet, lay thong tin zumenbango
                    double xMin2 = minPtArr[iChay].X + 2.7 * blockLenArr[iChay];
                    double yMin2 = minPtArr[iChay].Y-80*blockHeightArr[iChay];
                    Point3d minPt2 = new Point3d(xMin2, yMin2, 0);

                    PromptSelectionResult promSeRes6 = doc.Editor.SelectCrossingWindow(minPtArr[iChay], minPt2, SelFilZumenBango);
                    if (promSeRes6.Status == PromptStatus.OK)
                    {
                        SelectionSet Selset6 = promSeRes6.Value;
                        foreach (ObjectId objId6 in Selset6.GetObjectIds())
                        {
                            //xet 2 truong hop text va mtext
                            DBText text6 = tr.GetObject(objId6, OpenMode.ForRead) as DBText;
                            if (text6 != null)
                            {
                                ZumenBangoArr.Add(text6.TextString);
                            }

                            MText text6a = tr.GetObject(objId6, OpenMode.ForRead) as MText;
                            if (text6a != null)
                            {
                                ZumenBangoArr.Add(text6a.Text);
                            }
                        }
                    }
                    else
                    {
                        ZumenBangoArr.Add("0000");
                    }

                    #endregion

                    // Lay thong tin can dua vao QR Code
                    string extractedInfo = SeihinBangoArr[iChay].ToString() + "#"  + ShuzaiArr[iChay].ToString() + "#" + ShiageArr[iChay].ToString() + "#" + SuuRyoArr[iChay].ToString() + "#" + ZumenBangoArr[iChay].ToString();

                    // Generate QR code
                    Bitmap qrCodeImage = GenerateQRCode(extractedInfo);

                    // Save QR code image to file
                    string qrCodeFilePath = Path.Combine(directoryPath, $"QRCode_{iChay}.png");
                    qrCodeImage.Save(qrCodeFilePath, System.Drawing.Imaging.ImageFormat.Png);

                    #region diem dat QR Code
                    //tim diem dat QR Code
                    double blockLen = blockLenArr[iChay];
                    double x1=minPtArr[iChay].X+blockLen-blockLen * 45 /375;
                    double y1 = minPtArr[iChay].Y - blockLen * 45 / 375;
                    Point3d basePt = new Point3d(x1, y1, 0);
                    #endregion

                    // Insert QR code into drawing
                    InsertQRCodeIntoDrawing(doc, qrCodeFilePath, basePt, blockLen);
                }
                tr.Commit();
            }
        }
        #region ham getblock
        public static ObjectIdCollection GetBlockId()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database dat = doc.Database;

            ObjectIdCollection objIdCol = new ObjectIdCollection();
            string blName = "部品枠";

            using (Transaction tr3 = doc.TransactionManager.StartTransaction())
            {
                BlockTable blTb = tr3.GetObject(dat.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord blTbRec = tr3.GetObject(blTb[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                foreach (ObjectId objId in blTbRec)
                {
                    Entity entity1 = tr3.GetObject(objId, OpenMode.ForRead) as Entity;
                    if (entity1 is BlockReference)
                    {
                        BlockReference blRef = entity1 as BlockReference;
                        if (blRef.Name.Contains(blName))
                        {
                            objIdCol.Add(blRef.ObjectId);
                        }
                    }
                }
            }
            return objIdCol;
        }
        #endregion

        #region Ham GenerateQRCode
        private Bitmap GenerateQRCode(string info)
        {
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(info, QRCodeGenerator.ECCLevel.Q);
            QRCode qrCode = new QRCode(qrCodeData);
            return qrCode.GetGraphic(5);
        }
        #endregion

        #region ham InsertQRIntoCAD
        private void InsertQRCodeIntoDrawing(Document doc, string qrCodeFilePath, Point3d basePoint,double blockLen)
        {
            Database db = doc.Database;
            ObjectId acImgDefId = ObjectId.Null;

            using (Transaction tr1 = doc.TransactionManager.StartTransaction())
            {
                // Get the image dictionary
                ObjectId acImgDctID = RasterImageDef.GetImageDictionary(db);

                // Check to see if the dictionary does not exist, if not then create it
                if (acImgDctID.IsNull)
                {
                    acImgDctID = RasterImageDef.CreateImageDictionary(db);
                }

                // Open the image dictionary
                DBDictionary acImgDict = tr1.GetObject(acImgDctID, OpenMode.ForRead) as DBDictionary;
                RasterImageDef acRasterDef;
                //bool bRasterDefCreated;

                // Check to see if the image definition already exists
                // DWGファイルの名前には  <>/\":;?*|,=` の文字が使えません
                var name = System.IO.Path.GetFileNameWithoutExtension(qrCodeFilePath);
                if (acImgDict.Contains(name))
                {
                    acImgDefId = acImgDict.GetAt(name);
                    acRasterDef = tr1.GetObject(acImgDefId, OpenMode.ForWrite) as RasterImageDef;
                }
                else
                {
                    // Create a raster image definition
                    RasterImageDef acRasterDefNew = new RasterImageDef();

                    // Set the source for the image file
                    acRasterDefNew.SourceFileName = qrCodeFilePath;

                    // Load the image into memory
                    acRasterDefNew.Load();

                    // Add the image definition to the dictionary
                    acImgDict.UpgradeOpen();
                    acImgDefId = acImgDict.SetAt(name, acRasterDefNew);
                    tr1.AddNewlyCreatedDBObject(acRasterDefNew, true);
                    acRasterDef = acRasterDefNew;
                    //bRasterDefCreated = true;
                }

                // Open the Block table for read and the Block table record Model space for write
                BlockTable acBlkTbl = tr1.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord acBlkTblRec = tr1.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // Create the new image and assign it the image definition
                RasterImage acRaster = new RasterImage();
                acRaster.SetDatabaseDefaults(db);
                acRaster.ImageDefId = acImgDefId;

                // Define the width and height of the image
                double sizePicture = blockLen * 45 / 375; //QR du dinh dat tai cot cuoi cung cua khung
                var width = new Vector3d(sizePicture, 0, 0);
                var height = new Vector3d(0,sizePicture, 0);

                // Define the position for the image 
                Point3d insPt = basePoint;
                CoordinateSystem3d coordinateSystem = new CoordinateSystem3d(insPt, width, height);

                //// Define and assign a coordinate system for the image's orientation
                //CoordinateSystem3d coordinateSystem = new CoordinateSystem3d(insPt, width * 2, height * 2);
                acRaster.Orientation = coordinateSystem;
                //acRaster.Rotation = 0;

                // Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acRaster);
                tr1.AddNewlyCreatedDBObject(acRaster, true);

                tr1.Commit();
            }
        }
        #endregion
    }
}
