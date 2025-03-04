using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Teigha.ApplicationServices;
using Teigha.DatabaseServices;
using Teigha.Runtime;

namespace QRCodeInDrawing.TaoMaQR
{
    public class DeleteQRCodeInDrawing
    {
        [CommandMethod("DeleteQRCode")]
        public void DeleteQRCode()
        { 
            Document doc=Application.DocumentManager.MdiActiveDocument;
            Database dat=doc.Database;

            using (Transaction tr = dat.TransactionManager.StartTransaction())
            { 
                BlockTable blTb=tr.GetObject(dat.BlockTableId,OpenMode.ForRead)as BlockTable;
                BlockTableRecord blTbRec = tr.GetObject(blTb[BlockTableRecord.ModelSpace],OpenMode.ForWrite)as BlockTableRecord;

                foreach (ObjectId id in blTbRec)
                { 
                    Entity entity=tr.GetObject(id,OpenMode.ForRead)as Entity;

                    if (entity is RasterImage)
                    { 
                        RasterImage img=entity as RasterImage;
                        RasterImageDef imgDef=tr.GetObject(img.ImageDefId,OpenMode.ForRead)as RasterImageDef;

                        string imagePath=imgDef.SourceFileName;
                        string extension=Path.GetExtension(imagePath).ToLower();

                        if (extension == ".png")
                        {
                            img.UpgradeOpen();
                            img.Erase();
                        }
                    }
                }
                tr.Commit();
            }
            doc.Editor.WriteMessage("\n全てのQRコードは図面から消されました!!!!");
        }
    }
}
