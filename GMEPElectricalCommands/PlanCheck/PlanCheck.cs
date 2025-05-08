using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using MySql.Data.MySqlClient;

namespace ElectricalCommands.PlanCheck
{
  public abstract class PlanCheck
  {
    public string Name;
    public string Description;
    public string Query;
    public string ProjectId;
    public abstract string Check(List<ObjectId> blockList, MySqlConnection db);
  }

  public class CheckCaliforniaStamp : PlanCheck
  {
    public CheckCaliforniaStamp(string projectId)
    {
      Name = "Stamp";
      Description = "Checks if valid stamp is on plan";
      Query =
        @"
        SELECT
        projects.directory,
        project_titleblock_sizes.size
        FROM projects
        LEFT JOIN project_titleblock_sizes
        ON project_titleblock_sizes.id = projects.titleblock_size_id
        WHERE projects.id = @projectId AND projects.state = 'CA'
      ";
      ProjectId = projectId;
    }

    public override string Check(List<ObjectId> blockList, MySqlConnection connection)
    { // HERE set a directory in the database for the project anc test this
      MySqlCommand command = new MySqlCommand(Query, connection);
      command.Parameters.AddWithValue("projectId", ProjectId);
      MySqlDataReader reader = command.ExecuteReader();
      if (reader.Read())
      {
        string directory = reader.GetString(0);
        string titleblockSize = reader.GetString(1);
        string titleBlockPath = directory + $"\\XREF\\TBLK {titleblockSize}.dwg";
        reader.Close();
        if (!File.Exists(titleBlockPath))
        {
          titleBlockPath = directory + $"\\XREF\\TBLK.dwg";
          if (!File.Exists(titleBlockPath))
          {
            return "FAILED - Title block not found or not in XREF folder";
          }
        }
        Document doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Application
          .DocumentManager
          .MdiActiveDocument;
        Database db = doc.Database;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          List<ObjectId> idList = CADObjectCommands.GetObjectIdsFromBlockName(
            tr,
            blockList,
            "TBLK"
          );
          if (idList.Count > 0)
          {
            BlockReference br = tr.GetObject(idList[0], OpenMode.ForRead) as BlockReference;
            if (br != null)
            {
              BlockTableRecord btr = (BlockTableRecord)
                tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
              if (btr.IsFromExternalReference)
              {
                var xrefDb = btr.GetXrefDatabase(false);
                if (xrefDb != null)
                {
                  using (Transaction xrefTr = xrefDb.TransactionManager.StartTransaction())
                  {
                    BlockTable xrefBt = (BlockTable)
                      xrefTr.GetObject(xrefDb.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord xrefModelSpace = (BlockTableRecord)
                      xrefTr.GetObject(xrefBt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                    List<ObjectId> xrefModelSpaceObjectIds = new List<ObjectId>();
                    foreach (ObjectId id in xrefModelSpace)
                    {
                      xrefModelSpaceObjectIds.Add(id);
                      //try
                      //{
                      //  BlockReference brr = (BlockReference)xrefTr.GetObject(id, OpenMode.ForRead);
                      //  BlockTableRecord btrStandard = (BlockTableRecord)
                      //    xrefTr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                      //  Console.WriteLine("blocky " + btrStandard.Name);
                      //}
                      //catch { }
                    }
                    List<ObjectId> xrefIdList = CADObjectCommands.GetObjectIdsFromBlockName(
                      xrefTr,
                      xrefModelSpaceObjectIds,
                      "CA_E-STAMP"
                    );
                    if (xrefIdList.Count > 0)
                    {
                      BlockReference xrefBr =
                        tr.GetObject(xrefIdList[0], OpenMode.ForRead) as BlockReference;
                      if (xrefBr != null)
                      {
                        BlockTableRecord xrefBtr = (BlockTableRecord)
                          xrefTr.GetObject(xrefBr.BlockTableRecord, OpenMode.ForRead);
                        foreach (ObjectId id in xrefBtr)
                        {
                          Entity entity = (Entity)id.GetObject(OpenMode.ForRead);
                          if (entity != null)
                          {
                            if (entity.GetType() == typeof(DBText))
                            {
                              DBText text = (DBText)xrefTr.GetObject(id, OpenMode.ForRead);
                              Regex r = new Regex(
                                @"(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)[0-9]{2}"
                              );
                              Match m = r.Match(text.TextString);
                              if (m.Success)
                              {
                                DateTime dt = DateTime.ParseExact(
                                  text.TextString,
                                  "MM/dd/yyyy",
                                  new CultureInfo("en-Us"),
                                  DateTimeStyles.None
                                );
                                if (dt > DateTime.Today)
                                {
                                  return "PASSED";
                                }
                                else
                                {
                                  return "FAILED - CA stamp is out of date";
                                }
                              }
                              else
                              {
                                return "FAILED - date not found on CA stamp";
                              }
                            }
                          }
                        }
                        return "FAILED - stamp could not be opened as block reference";
                      }
                      else
                      {
                        return "FAILED - CA stamp was flattened or is invalid";
                      }
                    }
                    else
                    {
                      return "FAILED - CA stamp not found in XREF";
                    }
                  }
                }
                else
                {
                  return "FAILED - Titleblock XREF is empty";
                }
              }
              else
              {
                return "FAILED - Titleblock not an XREF";
              }
            }
            else
            {
              return "FAILED - Titleblock could not be opened as block reference";
            }
          }
          else
          {
            return "FAILED - Titleblock not found in block collection";
          }
        }
      }
      reader.Close();
      return "N/A - Project not in California";
    }
  }

  public class CheckKitchenNotes : PlanCheck
  {
    public CheckKitchenNotes(string projectId)
    {
      Name = "Kitchen Notes";
      Description =
        "Checks if kitchen notes are on plan if kitchen equipment exists for this project";
      Query =
        $"SELECT * FROM electrical_equipment WHERE project_id = @projectId AND category_id = 4";
      ProjectId = projectId;
    }

    public override string Check(List<ObjectId> blockList, MySqlConnection connection)
    {
      MySqlCommand command = new MySqlCommand(Query, connection);
      command.Parameters.AddWithValue("projectId", ProjectId);
      MySqlDataReader reader = command.ExecuteReader();
      if (reader.Read())
      {
        Document doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Application
          .DocumentManager
          .MdiActiveDocument;
        Database db = doc.Database;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          List<ObjectId> idList = CADObjectCommands.GetObjectIdsFromBlockName(
            tr,
            blockList,
            "GMEP KITCHEN NOTES"
          );
          if (idList.Count == 0)
          {
            return "FAILED - 'GMEP KITCHEN NOTES' block not found in file";
          }
          foreach (ObjectId id in idList)
          {
            Point3d position = CADObjectCommands.GetObjectPosition(id);
            if (position.X > 0 && position.X < 50)
            {
              if (position.Y > 12.4153 && position.Y < 36)
              {
                return "PASSED";
              }
            }
          }
          return "FAILED - Kitchen notes not placed on plan";
        }
      }
      else
      {
        return "N/A - No kitchen equipment in scope";
      }
    }
  }
}
