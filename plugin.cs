using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;

[assembly: CommandMethod("EXTRACT_BOM")]

namespace BoilerCadCopilot
{
    public class CadTextItem
    {
        public string raw_text { get; set; }
        public double x { get; set; }
        public double y { get; set; }
    }

    public class PluginPayload
    {
        public List<CadTextItem> items { get; set; }
    }

    public class Commands
    {
        private static readonly HttpClient client = new HttpClient();

        [CommandMethod("EXTRACT_BOM")]
        public async void ExtractBomFromCad()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\n[Agent] Bắt đầu quét các đối tượng Text trong bản vẽ...");

            List<CadTextItem> textList = new List<CadTextItem>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in btr)
                {
                    Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);

                    // 1. Quét DBText thường
                    if (ent is DBText dbText)
                    {
                        string val = dbText.TextString.Trim();
                        if (IsValidAnnotation(val))
                        {
                            textList.Add(new CadTextItem {
                                raw_text = val,
                                x = dbText.Position.X,
                                y = dbText.Position.Y
                            });
                        }
                    }
                    // 2. Quét MText
                    else if (ent is MText mText)
                    {
                        string val = mText.Contents.Trim();
                        if (IsValidAnnotation(val))
                        {
                            textList.Add(new CadTextItem {
                                raw_text = val,
                                x = mText.Location.X,
                                y = mText.Location.Y
                            });
                        }
                    }
                }
                tr.Commit();
            }

            if (textList.Count == 0)
            {
                ed.WriteMessage("\n[Agent] Không tìm thấy text ghi chú phù hợp (1-P, 2-P, Ống, V30...).");
                return;
            }

            ed.WriteMessage($"\n[Agent] Đã lọc được {textList.Count} ghi chú. Gửi đến Local Backend...");

            try
            {
                var payload = new PluginPayload { items = textList };
                string jsonBody = JsonConvert.SerializeObject(payload);
                var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync("http://localhost:8000/api/cad-plugin-extract", content);
                if (response.IsSuccessStatusCode)
                {
                    string resString = await response.Content.ReadAsStringAsync();
                    dynamic resData = JsonConvert.DeserializeObject(resString);
                    string excelPath = resData.excel_path;

                    ed.WriteMessage($"\n[Agent] Bóc tách hoàn tất! Mở file Excel: {excelPath}");
                    
                    // Tự động mở file Excel vừa tạo
                    Process.Start(new ProcessStartInfo(excelPath) { UseShellExecute = true });
                }
                else
                {
                    ed.WriteMessage($"\n[Agent Lỗi API]: {response.StatusCode}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Agent Lỗi Kết Nối]: {ex.Message}");
            }
        }

        private bool IsValidAnnotation(string text)
        {
            return text.Contains("1-P") || text.Contains("2-P") || 
                   text.Contains("1-O") || text.Contains("2-O") || 
                   text.Contains("V30") || text.Contains("SS400") || 
                   text.Contains("Láp") || text.Contains("Tôn");
        }
    }
}