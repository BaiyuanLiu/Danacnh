using NHSE.Core;
using NHSE.Sprites;
using NHSE.WinForms.Subforms.Map;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NHSE.WinForms
{
    public sealed partial class FieldItemEditor : Form, IItemLayerEditor
    {
        private readonly MainSave SAV;

        private readonly MapManager Map;
        private readonly MapViewer View;

        private bool Loading;
        private int SelectedBuildingIndex;

        private int HoverX;
        private int HoverY;
        private int DragX = -1;
        private int DragY = -1;
        private bool Dragging;

        public ItemEditor ItemProvider => ItemEdit;
        public ItemLayer SpawnLayer => Map.CurrentLayer;

        private TerrainBrushEditor? tbeForm;

        public FieldItemEditor(MainSave sav)
        {
            InitializeComponent();
            this.TranslateInterface(GameInfo.CurrentLanguage);

            var scale = (PB_Acre.Width - 2) / 32;
            SAV = sav;
            Map = new MapManager(sav);
            View = new MapViewer(Map, scale);

            Loading = true;

            LoadComboBoxes();
            LoadBuildings(sav);
            ReloadMapBackground();
            LoadEditors();
            LB_Items.SelectedIndex = 0;
            CB_Acre.SelectedIndex = 0;
            CB_MapAcre.SelectedIndex = 0;
            Loading = false;
            LoadItemGridAcre();
        }

        private void LoadComboBoxes()
        {
            foreach (var acre in MapGrid.Acres)
                CB_Acre.Items.Add(acre.Name);

            var exterior = AcreCoordinate.GetGridWithExterior(9, 8);
            foreach (var acre in exterior)
                CB_MapAcre.Items.Add(acre.Name);

            CB_MapAcreSelect.DisplayMember = nameof(ComboItem.Text);
            CB_MapAcreSelect.ValueMember = nameof(ComboItem.Value);
            CB_MapAcreSelect.DataSource = ComboItemUtil.GetArray<ushort>(typeof(OutsideAcre));

            NUD_MapAcreTemplateOutside.Value = SAV.OutsideFieldTemplateUniqueId;
            NUD_MapAcreTemplateField.Value = SAV.MainFieldParamUniqueID;
        }

        private void LoadBuildings(MainSave sav)
        {
            NUD_PlazaX.Value = sav.EventPlazaLeftUpX;
            NUD_PlazaY.Value = sav.EventPlazaLeftUpZ;

            foreach (var obj in Map.Buildings)
                LB_Items.Items.Add(obj.ToString());
        }

        private void LoadEditors()
        {
            var data = GameInfo.Strings.ItemDataSource.ToList();
            var field = FieldItemList.Items.Select(z => z.Value).ToList();
            data.Add(field, GameInfo.Strings.InternalNameTranslation);
            ItemEdit.Initialize(data, true);
            PG_TerrainTile.SelectedObject = new TerrainTile();
        }

        private int AcreIndex => CB_Acre.SelectedIndex;

        private void ChangeAcre(object sender, EventArgs e)
        {
            ChangeViewToAcre(AcreIndex);
            CB_MapAcre.Text = CB_Acre.Text;
        }

        private void ChangeViewToAcre(int acre)
        {
            View.SetViewToAcre(acre);
            LoadItemGridAcre();
        }

        private void LoadItemGridAcre()
        {
            ReloadItems();
            ReloadAcreBackground();
            UpdateArrowVisibility();
        }

        private int GetItemTransparency() => ((int)(0xFF * TR_Transparency.Value / 100d) << 24) | 0x00FF_FFFF;

        private void ReloadMapBackground()
        {
            PB_Map.BackgroundImage = View.GetBackgroundTerrain(SelectedBuildingIndex);
            PB_Map.Invalidate(); // background image reassigning to same img doesn't redraw; force it
        }

        private void ReloadAcreBackground()
        {
            var tbuild = (byte)TR_BuildingTransparency.Value;
            var tterrain = (byte)TR_Terrain.Value;
            PB_Acre.BackgroundImage = View.GetBackgroundAcre(L_Coordinates.Font, tbuild, tterrain, SelectedBuildingIndex);
            PB_Acre.Invalidate(); // background image reassigning to same img doesn't redraw; force it
        }

        private void ReloadMapItemGrid() => PB_Map.Image = View.GetMapWithReticle(GetItemTransparency());

        private void ReloadAcreItemGrid() => PB_Acre.Image = View.GetLayerAcre(GetItemTransparency());

        public void ReloadItems()
        {
            ReloadAcreItemGrid();
            ReloadMapItemGrid();
        }

        private void ReloadBuildingsTerrain()
        {
            ReloadAcreBackground();
            ReloadMapBackground();
        }

        private void UpdateArrowVisibility()
        {
            B_Up.Enabled = View.CanUp;
            B_Down.Enabled = View.CanDown;
            B_Left.Enabled = View.CanLeft;
            B_Right.Enabled = View.CanRight;
        }

        private void PB_Acre_MouseClick(object sender, MouseEventArgs e)
        {
            if (Dragging)
            {
                ResetDrag();
                return;
            }

            if (RB_Item.Checked)
                OmniTile(e);
            else if (RB_Terrain.Checked)
                OmniTileTerrain(e);
        }

        private void ResetDrag()
        {
            DragX = -1;
            DragY = -1;
            Dragging = false;
        }

        private void OmniTile(MouseEventArgs e)
        {
            var tile = GetTile(Map.CurrentLayer, e, out var x, out var y);
            OmniTile(tile, x, y);
        }

        private void OmniTileTerrain(MouseEventArgs e)
        {
            SetHoveredItem(e);
            var x = View.X + HoverX;
            var y = View.Y + HoverY;
            var tile = Map.Terrain.GetTile(x / 2, y / 2);
            if (tbeForm?.brushSelected != true)
            {
                OmniTileTerrain(tile);
                return;
            }

            if (tbeForm.Slider_thickness.Value <= 1)
            {
                SetTile(tile);
                return;
            }

            List<TerrainTile> selectedTiles = new();
            int radius = tbeForm.Slider_thickness.Value;
            int threshold = (radius * radius) / 2;
            for (int i = -radius; i < radius; i++)
            {
                for (int j = -radius; j < radius; j++)
                {
                    if ((i * i) + (j * j) < threshold)
                        selectedTiles.Add(Map.Terrain.GetTile((x / 2) + i, (y / 2) + j));
                }
            }

            SetTiles(selectedTiles);
        }

        private void OmniTile(Item tile, int x, int y)
        {
            switch (ModifierKeys)
            {
                default:
                    ViewTile(tile, x, y);
                    return;

                case Keys.Alt | Keys.Control:
                case Keys.Alt | Keys.Control | Keys.Shift:
                    ReplaceTile(tile, x, y);
                    return;

                case Keys.Shift:
                    SetTile(tile, x, y);
                    return;

                case Keys.Alt:
                    DeleteTile(tile, x, y);
                    return;
            }
        }

        private void OmniTileTerrain(TerrainTile tile)
        {
            switch (ModifierKeys)
            {
                default:
                    ViewTile(tile);
                    return;

                case Keys.Shift | Keys.Control:
                    RotateTile(tile);
                    return;

                case Keys.Shift:
                    SetTile(tile);
                    return;

                case Keys.Alt:
                    DeleteTile(tile);
                    return;
            }
        }

        private Item GetTile(FieldItemLayer layer, MouseEventArgs e, out int x, out int y)
        {
            SetHoveredItem(e);
            return layer.GetTile(x = View.X + HoverX, y = View.Y + HoverY);
        }

        private void SetHoveredItem(MouseEventArgs e)
        {
            GetAcreCoordinates(e, out HoverX, out HoverY);

            // Mouse event may fire with a slightly too large x/y; clamp just in case.
            HoverX &= 0x1F;
            HoverY &= 0x1F;
        }

        private void GetAcreCoordinates(MouseEventArgs e, out int x, out int y)
        {
            x = e.X / View.AcreScale;
            y = e.Y / View.AcreScale;
        }

        private void PB_Acre_MouseDown(object sender, MouseEventArgs e) => ResetDrag();

        private void PB_Acre_MouseMove(object sender, MouseEventArgs e)
        {
            var l = Map.CurrentLayer;
            if (e.Button == MouseButtons.Left && CHK_MoveOnDrag.Checked)
            {
                MoveDrag(e);
                return;
            }
            if (e.Button == MouseButtons.Left && tbeForm?.brushSelected == true)
            {
                OmniTileTerrain(e);
            }

            var oldTile = l.GetTile(View.X + HoverX, View.Y + HoverY);
            var tile = GetTile(l, e, out var x, out var y);
            if (ReferenceEquals(tile, oldTile))
                return;
            var str = GameInfo.Strings;
            var name = str.GetItemName(tile);
            bool active = Map.Items.GetIsActive(NUD_Layer.Value == 0, x, y);
            if (active)
                name = $"{name} [Active]";
            TT_Hover.SetToolTip(PB_Acre, name);
            SetCoordinateText(x, y);
        }

        private void MoveDrag(MouseEventArgs e)
        {
            GetAcreCoordinates(e, out var nhX, out var nhY);

            if (DragX == -1)
            {
                DragX = nhX;
                DragY = nhY;
                return;
            }

            var dX = DragX - nhX;
            var dY = DragY - nhY;

            if (ModifierKeys == Keys.Control)
            {
                dX *= 2;
                dY *= 2;
            }

            if ((dX & 1) == 1)
                dX ^= 1;
            if ((dY & 1) == 1)
                dY ^= 1;

            var aX = Math.Abs(dX);
            var aY = Math.Abs(dY);
            if (aX < 2 && aY < 2)
                return;

            DragX = nhX;
            DragY = nhY;
            if (!View.SetViewTo(View.X + dX, View.Y + dY))
                return;

            Dragging = true;
            LoadItemGridAcre();
        }

        private void ViewTile(Item tile, int x, int y)
        {
            if (CHK_RedirectExtensionLoad.Checked && tile.IsExtension)
            {
                var l = Map.CurrentLayer;
                var rx = Math.Max(0, Math.Min(l.MaxWidth - 1, x - tile.ExtensionX));
                var ry = Math.Max(0, Math.Min(l.MaxHeight - 1, y - tile.ExtensionY));
                var redir = l.GetTile(rx, ry);
                if (redir.IsRoot && redir.ItemId == tile.ExtensionItemId)
                    tile = redir;
            }

            ViewTile(tile);
        }

        private void ViewTile(Item tile)
        {
            ItemEdit.LoadItem(tile);
            TC_Editor.SelectedTab = Tab_Item;
        }

        private void ViewTile(TerrainTile tile)
        {
            var pgt = (TerrainTile)PG_TerrainTile.SelectedObject;
            pgt.CopyFrom(tile);
            PG_TerrainTile.SelectedObject = pgt;
            TC_Editor.SelectedTab = Tab_Terrain;
        }

        private void SetTile(Item tile, int x, int y)
        {
            var l = Map.CurrentLayer;
            var pgt = new Item();
            ItemEdit.SetItem(pgt);

            if (pgt.IsFieldItem && CHK_FieldItemSnap.Checked)
            {
                // coordinates must be even (not odd-half)
                x &= 0xFFFE;
                y &= 0xFFFE;
                tile = l.GetTile(x, y);
            }

            var permission = l.IsOccupied(pgt, x, y);
            switch (permission)
            {
                case PlacedItemPermission.OutOfBounds:
                case PlacedItemPermission.Collision when CHK_NoOverwrite.Checked:
                    System.Media.SystemSounds.Asterisk.Play();
                    return;
            }

            // Clean up original placed data
            if (tile.IsRoot && CHK_AutoExtension.Checked)
                l.DeleteExtensionTiles(tile, x, y);

            // Set new placed data
            if (pgt.IsRoot && CHK_AutoExtension.Checked)
                l.SetExtensionTiles(pgt, x, y);
            tile.CopyFrom(pgt);

            ReloadItems();
        }

        private void ReplaceTile(Item tile, int x, int y)
        {
            var l = Map.CurrentLayer;
            var pgt = new Item();
            ItemEdit.SetItem(pgt);

            if (pgt.IsFieldItem && CHK_FieldItemSnap.Checked)
            {
                // coordinates must be even (not odd-half)
                x &= 0xFFFE;
                y &= 0xFFFE;
                tile = l.GetTile(x, y);
            }

            var permission = l.IsOccupied(pgt, x, y);
            switch (permission)
            {
                case PlacedItemPermission.OutOfBounds:
                    System.Media.SystemSounds.Asterisk.Play();
                    return;
            }

            bool wholeMap = (ModifierKeys & Keys.Shift) != 0;
            var copy = new Item(tile.RawValue);
            var count = View.ReplaceFieldItems(copy, pgt, wholeMap);
            if (count == 0)
            {
                WinFormsUtil.Alert(MessageStrings.MsgFieldItemModifyNone);
                return;
            }
            LoadItemGridAcre();
            WinFormsUtil.Alert(string.Format(MessageStrings.MsgFieldItemModifyCount, count));
        }

        private void RotateTile(TerrainTile tile)
        {
            bool rotated = tile.Rotate();
            if (!rotated)
            {
                System.Media.SystemSounds.Asterisk.Play();
                return;
            }
            ReloadBuildingsTerrain();
        }

        private void SetTile(TerrainTile tile)
        {
            var pgt = (TerrainTile)PG_TerrainTile.SelectedObject;
            if (tbeForm?.randomizeVariation == true)
            {
                switch (pgt.UnitModel)
                {
                    case TerrainUnitModel.Cliff5B:
                    case TerrainUnitModel.River5B:
                        Random rand = new();
                        pgt.Variation = (ushort)rand.Next(4);
                        break;
                }
            }

            tile.CopyFrom(pgt);

            ReloadBuildingsTerrain();
        }

        private void SetTiles(IEnumerable<TerrainTile> tiles)
        {
            var pgt = (TerrainTile)PG_TerrainTile.SelectedObject;
            foreach (TerrainTile tile in tiles)
            {
                tile.CopyFrom(pgt);
            }

            ReloadBuildingsTerrain();
        }

        private void DeleteTile(Item tile, int x, int y)
        {
            if (CHK_AutoExtension.Checked)
            {
                if (!tile.IsRoot)
                {
                    x -= tile.ExtensionX;
                    y -= tile.ExtensionY;
                    tile = Map.CurrentLayer.GetTile(x, y);
                }
                Map.CurrentLayer.DeleteExtensionTiles(tile, x, y);
            }

            tile.Delete();
            ReloadItems();
        }

        private void DeleteTile(TerrainTile tile)
        {
            tile.Clear();
            ReloadBuildingsTerrain();
        }

        private void B_Cancel_Click(object sender, EventArgs e) => Close();

        private void B_Save_Click(object sender, EventArgs e)
        {
            var unsupported = Map.Items.GetUnsupportedTiles();
            if (unsupported.Count != 0)
            {
                var err = MessageStrings.MsgFieldItemUnsupportedLayer2Tile;
                var ask = MessageStrings.MsgAskContinue;
                var prompt = WinFormsUtil.Prompt(MessageBoxButtons.YesNo, err, ask);
                if (prompt != DialogResult.Yes)
                    return;
            }

            Map.Items.Save();
            SAV.SetTerrainTiles(Map.Terrain.Tiles);

            SAV.SetAcreBytes(Map.Terrain.BaseAcres);
            SAV.OutsideFieldTemplateUniqueId = (ushort)NUD_MapAcreTemplateOutside.Value;
            SAV.MainFieldParamUniqueID = (ushort)NUD_MapAcreTemplateField.Value;

            SAV.Buildings = Map.Buildings;
            SAV.EventPlazaLeftUpX = Map.PlazaX;
            SAV.EventPlazaLeftUpZ = Map.PlazaY;
            Close();
        }

        private void Menu_View_Click(object sender, EventArgs e)
        {
            var x = View.X + HoverX;
            var y = View.Y + HoverY;

            if (RB_Item.Checked)
            {
                var tile = Map.CurrentLayer.GetTile(x, y);
                ViewTile(tile, x, y);
            }
            else if (RB_Terrain.Checked)
            {
                var tile = Map.Terrain.GetTile(x / 2, y / 2);
                ViewTile(tile);
            }
        }

        private void Menu_Set_Click(object sender, EventArgs e)
        {
            var x = View.X + HoverX;
            var y = View.Y + HoverY;

            if (RB_Item.Checked)
            {
                var tile = Map.CurrentLayer.GetTile(x, y);
                SetTile(tile, x, y);
            }
            else if (RB_Terrain.Checked)
            {
                var tile = Map.Terrain.GetTile(x / 2, y / 2);
                SetTile(tile);
            }
        }

        private void Menu_Reset_Click(object sender, EventArgs e)
        {
            var x = View.X + HoverX;
            var y = View.Y + HoverY;

            if (RB_Item.Checked)
            {
                var tile = Map.CurrentLayer.GetTile(x, y);
                DeleteTile(tile, x, y);
            }
            else if (RB_Terrain.Checked)
            {
                var tile = Map.Terrain.GetTile(x / 2, y / 2);
                DeleteTile(tile);
            }
        }

        private bool hasActivate = true;

        private void CM_Click_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!RB_Item.Checked)
            {
                if (hasActivate)
                    CM_Click.Items.Remove(Menu_Activate);
                hasActivate = false;
                return;
            }

            var isBase = NUD_Layer.Value == 0;
            var x = View.X + HoverX;
            var y = View.Y + HoverY;
            Menu_Activate.Text = Map.Items.GetIsActive(isBase, x, y) ? "Inactivate" : "Activate";
            CM_Click.Items.Add(Menu_Activate);
            hasActivate = true;
        }

        private void Menu_Activate_Click(object sender, EventArgs e)
        {
            var x = View.X + HoverX;
            var y = View.Y + HoverY;
            var isBase = NUD_Layer.Value == 0;
            Map.Items.SetIsActive(isBase, x, y, !Map.Items.GetIsActive(isBase, x, y));
        }

        private void B_Up_Click(object sender, EventArgs e)
        {
            if (ModifierKeys == Keys.Shift)
                CB_Acre.SelectedIndex = Math.Max(0, CB_Acre.SelectedIndex - MapGrid.AcreWidth);
            else if (View.ArrowUp())
                LoadItemGridAcre();
        }

        private void B_Left_Click(object sender, EventArgs e)
        {
            if (ModifierKeys == Keys.Shift)
                CB_Acre.SelectedIndex = Math.Max(0, CB_Acre.SelectedIndex - 1);
            else if (View.ArrowLeft())
                LoadItemGridAcre();
        }

        private void B_Right_Click(object sender, EventArgs e)
        {
            if (ModifierKeys == Keys.Shift)
                CB_Acre.SelectedIndex = Math.Min(CB_Acre.SelectedIndex + 1, CB_Acre.Items.Count - 1);
            else if (View.ArrowRight())
                LoadItemGridAcre();
        }

        private void B_Down_Click(object sender, EventArgs e)
        {
            if (ModifierKeys == Keys.Shift)
                CB_Acre.SelectedIndex = Math.Min(CB_Acre.SelectedIndex + MapGrid.AcreWidth, CB_Acre.Items.Count - 1);
            else if (View.ArrowDown())
                LoadItemGridAcre();
        }

        private void B_DumpAcre_Click(object sender, EventArgs e) => MapDumpHelper.DumpLayerAcreSingle(Map.CurrentLayer, AcreIndex, CB_Acre.Text, (int)NUD_Layer.Value);
        // 调整参数：将 pixelSize 改为 scaleFactor（缩放倍数），更符合语义
        private void ExportMapToImage(string savePath)
        {
            // 1. 初始化日志（保留原有日志逻辑）
            string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "导出地图日志.txt");
            File.WriteAllText(logPath, $"===== 导出开始：{DateTime.Now} =====\r\n");
            Action<string> log = (msg) => {
                File.AppendAllText(logPath, $"{DateTime.Now:HH:mm:ss.fff} | {msg}\r\n");
                Console.WriteLine(msg);
            };

            try
            {
                log("开始校验基础参数");
                if (string.IsNullOrEmpty(savePath))
                    throw new ArgumentNullException(nameof(savePath), "保存路径不能为空");

                if (Map.CurrentLayer is not FieldItemLayer fieldLayer)
                    throw new InvalidOperationException("当前图层不是物品层，无法导出图片！");

                if (fieldLayer.MaxWidth <= 0 || fieldLayer.MaxHeight <= 0)
                    throw new ArgumentException("图层尺寸无效（宽度/高度不能≤0）", nameof(fieldLayer));

                // 核心配置：单个物品占2×2个原Tile（64×64×2=128×128）
                const int originalTileSize = 16;       // 原单个Tile尺寸
                const int expandRatio = 2;             // 扩大倍数（2×2格）
                const int expandedTileSize = originalTileSize * expandRatio; // 128×128

                log($"原Tile尺寸：{originalTileSize}×{originalTileSize}，扩大后：{expandedTileSize}×{expandedTileSize}（占2×2格）");

                // 画布尺寸：仍按原Tile数×原尺寸（总大小不变）
                int canvasWidth = fieldLayer.MaxWidth * originalTileSize;
                int canvasHeight = fieldLayer.MaxHeight * originalTileSize;
                log($"计算画布尺寸：{canvasWidth}×{canvasHeight}");

                // 校验画布尺寸
                if (canvasWidth <= 0 || canvasHeight <= 0 || canvasWidth > 32767 || canvasHeight > 32767)
                    throw new InvalidOperationException($"画布尺寸无效：{canvasWidth}×{canvasHeight}（需在1~32767之间）");

                // 确保保存目录存在
                string saveDir = Path.GetDirectoryName(savePath);
                log($"保存目录：{saveDir}");
                if (!Directory.Exists(saveDir))
                {
                    Directory.CreateDirectory(saveDir);
                    log("创建保存目录成功");
                }

                // 创建Bitmap画布
                log("开始创建Bitmap画布");
                Bitmap mapImage = new Bitmap(canvasWidth, canvasHeight);
                log("Bitmap创建成功");

                try
                {
                    log("开始创建Graphics对象");
                    using (Graphics g = Graphics.FromImage(mapImage))
                    {
                        g.Clear(Color.White);
                        // 优化缩放清晰度（像素风格保留）
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None;
                        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.None;
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                        log("画布清空完成，开始遍历Tile（步长2）");

                        // 遍历Tile：步长改为2（避免重复绘制）
                        int totalTile = (fieldLayer.MaxWidth / expandRatio) * (fieldLayer.MaxHeight / expandRatio);
                        int drawnTile = 0;
                        // x/y每次+2，只处理偶数位置的Tile（保证占2×2格不重叠）
                        for (int x = 0; x < fieldLayer.MaxWidth; x += expandRatio)
                        {
                            for (int y = 0; y < fieldLayer.MaxHeight; y += expandRatio)
                            {
                                drawnTile++;
                                if (drawnTile % 100 == 0) log($"已绘制{drawnTile}/{totalTile}个Tile");

                                Item tile = fieldLayer.GetTile(x, y);
                                if (tile.ItemId == Item.NONE) continue;

                                // 计算绘制坐标（基于原Tile尺寸）
                                int drawX = x * originalTileSize;
                                int drawY = y * originalTileSize;

                                // 关键：校验扩大后的图片是否越界
                                if (drawX + expandedTileSize > canvasWidth || drawY + expandedTileSize > canvasHeight)
                                {
                                    log($"跳过越界Tile：[{x},{y}]，扩大后坐标：{drawX},{drawY}，尺寸：{expandedTileSize}×{expandedTileSize}");
                                    continue;
                                }

                                // 获取并绘制物品图片（扩大到128×128）
                                log($"获取Tile[{x},{y}]物品[{tile.ItemId}]图片");
                                Image? itemImage = ItemSprite.GetItemSprite(tile);
                                if (itemImage != null)
                                {
                                    log($"Tile[{x},{y}]物品图片不为空，尺寸：{itemImage.Width}×{itemImage.Height}");
                                    try
                                    {
                                        using (Image cloneImage = new Bitmap(itemImage))
                                        {
                                            log($"开始绘制Tile[{x},{y}]物品图片（扩大到{expandedTileSize}×{expandedTileSize}）");
                                            g.DrawImage(
                                                cloneImage,
                                                new Rectangle(drawX, drawY, expandedTileSize, expandedTileSize), // 目标尺寸：128×128
                                                new Rectangle(0, 0, cloneImage.Width, cloneImage.Height),       // 原图区域
                                                GraphicsUnit.Pixel
                                            );
                                            log($"Tile[{x},{y}]物品图片绘制成功");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        log($"绘制Tile[{x},{y}]失败：{ex.Message}，改用占位符");
                                        DrawDefaultTile(g, x, y, originalTileSize, expandRatio, tile.ItemId);
                                    }
                                }
                                else
                                {
                                    log($"Tile[{x},{y}]物品图片为空，绘制占位符（扩大到{expandedTileSize}×{expandedTileSize}）");
                                    DrawDefaultTile(g, x, y, originalTileSize, expandRatio, tile.ItemId);
                                }
                            }
                        }
                        log("所有Tile遍历完成");
                    }

                    // 保存图片
                    log("开始保存图片到指定路径");
                    mapImage.Save(savePath, ImageFormat.Png);
                    log("图片保存成功");
                }
                finally
                {
                    mapImage.Dispose();
                    log("Bitmap资源已释放");
                }

                log("===== 导出完成 =====");
            }
            catch (Exception ex)
            {
                log($"===== 导出异常：{ex.Message} =====");
                log($"异常类型：{ex.GetType().Name}");
                log($"异常堆栈：{ex.StackTrace}");
                log($"内部异常：{ex.InnerException?.Message ?? "无"}");
                throw;
            }
        }

        // 辅助方法：适配扩大后的占位符（占2×2格）
        private void DrawDefaultTile(Graphics g, int x, int y, int originalTileSize, int expandRatio, ushort itemId)
        {
            int drawX = x * originalTileSize;
            int drawY = y * originalTileSize;
            int expandedTileSize = originalTileSize * expandRatio; // 128×128

            if (expandedTileSize <= 2) return;

            // 绘制蓝色方块（占2×2格）
            using (Brush brush = new SolidBrush(Color.LightBlue))
            using (Pen pen = new Pen(Color.Black, 2))
            {
                g.FillRectangle(brush, drawX, drawY, expandedTileSize - 2, expandedTileSize - 2);
                g.DrawRectangle(pen, drawX, drawY, expandedTileSize - 2, expandedTileSize - 2);
            }

            // 绘制物品ID（字体大小适配128×128）
            using (Font font = new Font("Arial", 24, FontStyle.Regular, GraphicsUnit.Pixel)) // 24号字体适配128×128
            using (Brush textBrush = new SolidBrush(Color.Black))
            {
                SizeF textSize = g.MeasureString(itemId.ToString(), font);
                // 文字居中显示
                float textX = drawX + (expandedTileSize - textSize.Width) / 2;
                float textY = drawY + (expandedTileSize - textSize.Height) / 2;
                if (textX >= 0 && textY >= 0 && textX + textSize.Width < drawX + expandedTileSize)
                {
                    g.DrawString(itemId.ToString(), font, textBrush, textX, textY);
                }
            }
        }

        // 调用处修改（去掉pixelSize参数）
        private void B_DumpAllAcres_Click(object sender, EventArgs e)
        {
            MapDumpHelper.DumpLayerAcreAll(Map.CurrentLayer);

            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "PNG图片 (*.png)|*.png|所有文件 (*.*)|*.*";
                    sfd.FileName = $"地图导出_{DateTime.Now:yyyyMMddHHmmss}.png";
                    sfd.Title = "保存地图图片";

                    if (sfd.ShowDialog() != DialogResult.OK)
                        return;

                    // 直接调用（无需pixelSize）
                    ExportMapToImage(sfd.FileName);

                    MessageBox.Show(
                        $"地图图片导出成功！\n路径：{sfd.FileName}",
                        "导出成功",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"地图图片导出失败：{ex.Message}",
                    "导出失败",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void B_ImportAcre_Click(object sender, EventArgs e)
        {
            var layer = Map.CurrentLayer;
            if (!MapDumpHelper.ImportToLayerAcreSingle(layer, AcreIndex, CB_Acre.Text, (int)NUD_Layer.Value))
                return;
            ChangeViewToAcre(AcreIndex);
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void B_ImportAllAcres_Click(object sender, EventArgs e)
        {
            if (!MapDumpHelper.ImportToLayerAcreAll(Map.CurrentLayer))
                return;
            ChangeViewToAcre(AcreIndex);
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void B_DumpBuildings_Click(object sender, EventArgs e) => MapDumpHelper.DumpBuildings(Map.Buildings);

        private void B_ImportBuildings_Click(object sender, EventArgs e)
        {
            if (!MapDumpHelper.ImportBuildings(Map.Buildings))
                return;

            for (int i = 0; i < Map.Buildings.Count; i++)
                LB_Items.Items[i] = Map.Buildings[i].ToString();
            LB_Items.SelectedIndex = 0;
            System.Media.SystemSounds.Asterisk.Play();
            ReloadBuildingsTerrain();
        }

        private void B_DumpTerrainAcre_Click(object sender, EventArgs e) => MapDumpHelper.DumpTerrainAcre(Map.Terrain, AcreIndex, CB_Acre.Text);

        private void B_DumpTerrainAll_Click(object sender, EventArgs e) => MapDumpHelper.DumpTerrainAll(Map.Terrain);

        private void B_ImportTerrainAcre_Click(object sender, EventArgs e)
        {
            if (!MapDumpHelper.ImportTerrainAcre(Map.Terrain, AcreIndex, CB_Acre.Text))
                return;
            ChangeViewToAcre(AcreIndex);
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void B_ImportTerrainAll_Click(object sender, EventArgs e)
        {
            if (!MapDumpHelper.ImportTerrainAll(Map.Terrain))
                return;
            ChangeViewToAcre(AcreIndex);
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void Menu_SavePNG_Click(object sender, EventArgs e)
        {
            var pb = WinFormsUtil.GetUnderlyingControl<PictureBox>(sender);
            if (pb?.Image == null)
            {
                WinFormsUtil.Alert(MessageStrings.MsgNoPictureLoaded);
                return;
            }

            CM_Picture.Close(ToolStripDropDownCloseReason.CloseCalled);

            const string name = "map";
            using var sfd = new SaveFileDialog
            {
                Filter = "png file (*.png)|*.png|All files (*.*)|*.*",
                FileName = $"{name}.png",
            };
            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            if (!Menu_SavePNGTerrain.Checked)
            {
                PB_Map.Image.Save(sfd.FileName, ImageFormat.Png);
            }
            else if (!Menu_SavePNGItems.Checked)
            {
                PB_Map.BackgroundImage!.Save(sfd.FileName, ImageFormat.Png);
            }
            else
            {
                var img = (Bitmap)PB_Map.BackgroundImage!.Clone();
                using var gfx = Graphics.FromImage(img);
                gfx.DrawImage(PB_Map.Image, new Point(0, 0));
                img.Save(sfd.FileName, ImageFormat.Png);
            }
        }

        private void CM_Picture_Closing(object sender, ToolStripDropDownClosingEventArgs e)
        {
            if (e.CloseReason == ToolStripDropDownCloseReason.ItemClicked && sender != Menu_SavePNG)
                e.Cancel = true;
        }

        private void PB_Map_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;
            ClickMapAt(e, true);
        }

        private void ClickMapAt(MouseEventArgs e, bool skipLagCheck)
        {
            var layer = Map.Items.Layer1;
            int mX = e.X;
            int mY = e.Y;
            bool centerReticle = CHK_SnapToAcre.Checked;
            View.GetViewAnchorCoordinates(mX, mY, out var x, out var y, centerReticle);
            x &= 0xFFFE;
            y &= 0xFFFE;

            var acre = layer.GetAcre(x, y);
            bool sameAcre = AcreIndex == acre;
            if (!skipLagCheck)
            {
                if (CHK_SnapToAcre.Checked)
                {
                    if (sameAcre)
                        return;
                }
                else
                {
                    const int delta = 0; // disabled = 0
                    var dx = Math.Abs(View.X - x);
                    var dy = Math.Abs(View.Y - y);
                    if (dx <= delta && dy <= delta && !sameAcre)
                        return;
                }
            }

            if (!CHK_SnapToAcre.Checked)
            {
                if (View.SetViewTo(x, y))
                    LoadItemGridAcre();
                return;
            }

            if (!sameAcre)
                CB_Acre.SelectedIndex = acre;
        }

        private void PB_Map_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ClickMapAt(e, false);
            }
            else if (e.Button == MouseButtons.None)
            {
                View.GetCursorCoordinates(e.X, e.Y, out var x, out var y);
                SetCoordinateText(x, y);
            }
        }

        private void SetCoordinateText(int x, int y) => L_Coordinates.Text = $"({x:000},{y:000}) = (0x{x:X2},0x{y:X2})";

        private void NUD_Layer_ValueChanged(object sender, EventArgs e)
        {
            Map.MapLayer = (int)NUD_Layer.Value - 1;
            LoadItemGridAcre();
        }

        private void Remove(ToolStripItem sender, Func<int, int, int, int, int> removal)
        {
            bool wholeMap = (ModifierKeys & Keys.Shift) != 0;

            string q = string.Format(MessageStrings.MsgFieldItemRemoveAsk, sender.Text);
            var question = WinFormsUtil.Prompt(MessageBoxButtons.YesNo, q);
            if (question != DialogResult.Yes)
                return;

            int count = View.ModifyFieldItems(removal, wholeMap);

            if (count == 0)
            {
                WinFormsUtil.Alert(MessageStrings.MsgFieldItemRemoveNone);
                return;
            }
            LoadItemGridAcre();
            WinFormsUtil.Alert(string.Format(MessageStrings.MsgFieldItemRemoveCount, count));
        }

        private void Modify(ToolStripItem sender, Func<int, int, int, int, int> action)
        {
            bool wholeMap = (ModifierKeys & Keys.Shift) != 0;

            string q = string.Format(MessageStrings.MsgFieldItemModifyAsk, sender.Text);
            var question = WinFormsUtil.Prompt(MessageBoxButtons.YesNo, q);
            if (question != DialogResult.Yes)
                return;

            int count = View.ModifyFieldItems(action, wholeMap);

            if (count == 0)
            {
                WinFormsUtil.Alert(MessageStrings.MsgFieldItemModifyNone);
                return;
            }
            LoadItemGridAcre();
            WinFormsUtil.Alert(string.Format(MessageStrings.MsgFieldItemModifyCount, count));
        }

        private void B_RemoveEditor_Click(object sender, EventArgs e) => Remove(B_RemoveEditor, (min, max, x, y)
            => Map.CurrentLayer.RemoveAllLike(min, max, x, y, ItemEdit.SetItem(new Item())));

        private void B_RemoveAllWeeds_Click(object sender, EventArgs e) => Remove(B_RemoveAllWeeds, Map.CurrentLayer.RemoveAllWeeds);

        private void B_RemoveAllTrees_Click(object sender, EventArgs e) => Remove(B_RemoveAllTrees, Map.CurrentLayer.RemoveAllTrees);

        private void B_FillHoles_Click(object sender, EventArgs e) => Remove(B_FillHoles, Map.CurrentLayer.RemoveAllHoles);

        private void B_RemovePlants_Click(object sender, EventArgs e) => Remove(B_RemovePlants, Map.CurrentLayer.RemoveAllPlants);

        private void B_RemoveFences_Click(object sender, EventArgs e) => Remove(B_RemoveFences, Map.CurrentLayer.RemoveAllFences);

        private void B_RemoveObjects_Click(object sender, EventArgs e) => Remove(B_RemoveObjects, Map.CurrentLayer.RemoveAllObjects);

        private void B_RemoveAll_Click(object sender, EventArgs e) => Remove(B_RemoveAll, Map.CurrentLayer.RemoveAll);

        private void B_RemovePlacedItems_Click(object sender, EventArgs e) => Remove(B_RemovePlacedItems, Map.CurrentLayer.RemoveAllPlacedItems);

        private void B_RemoveShells_Click(object sender, EventArgs e) => Remove(B_RemoveShells, Map.CurrentLayer.RemoveAllShells);

        private void B_RemoveBranches_Click(object sender, EventArgs e) => Remove(B_RemoveBranches, Map.CurrentLayer.RemoveAllBranches);

        private void B_RemoveFlowers_Click(object sender, EventArgs e) => Remove(B_RemoveFlowers, Map.CurrentLayer.RemoveAllFlowers);

        private void B_RemoveBushes_Click(object sender, EventArgs e) => Remove(B_RemoveBushes, Map.CurrentLayer.RemoveAllBushes);

        private void B_WaterFlowers_Click(object sender, EventArgs e) => Modify(B_WaterFlowers, (xmin, ymin, width, height)
            => Map.CurrentLayer.WaterAllFlowers(xmin, ymin, width, height, (ModifierKeys & Keys.Control) != 0));

        private static void ShowContextMenuBelow(ToolStripDropDown c, Control n) => c.Show(n.PointToScreen(new Point(0, n.Height)));

        private void B_RemoveItemDropDown_Click(object sender, EventArgs e) => ShowContextMenuBelow(CM_Remove, B_RemoveItemDropDown);

        private void B_DumpLoadField_Click(object sender, EventArgs e) => ShowContextMenuBelow(CM_DLField, B_DumpLoadField);

        private void B_DumpLoadTerrain_Click(object sender, EventArgs e) => ShowContextMenuBelow(CM_DLTerrain, B_DumpLoadTerrain);

        private void B_DumpLoadBuildings_Click(object sender, EventArgs e) => ShowContextMenuBelow(CM_DLBuilding, B_DumpLoadBuildings);

        private void B_ModifyAllTerrain_Click(object sender, EventArgs e) => ShowContextMenuBelow(CM_Terrain, B_ModifyAllTerrain);

        private void B_DumpLoadAcres_Click(object sender, EventArgs e) => ShowContextMenuBelow(CM_DLMapAcres, B_DumpLoadAcres);

        private void TR_Transparency_Scroll(object sender, EventArgs e) => ReloadItems();

        private void TR_BuildingTransparency_Scroll(object sender, EventArgs e) => ReloadBuildingsTerrain();

        private void TR_Terrain_Scroll(object sender, EventArgs e) => ReloadBuildingsTerrain();

        #region Buildings

        private void B_Help_Click(object sender, EventArgs e)
        {
            using var form = new BuildingHelp();
            form.ShowDialog();
        }

        private void NUD_PlazaX_ValueChanged(object sender, EventArgs e)
        {
            if (Loading)
                return;
            Map.PlazaX = (uint)NUD_PlazaX.Value;
            ReloadBuildingsTerrain();
        }

        private void NUD_PlazaY_ValueChanged(object sender, EventArgs e)
        {
            if (Loading)
                return;
            Map.PlazaY = (uint)NUD_PlazaY.Value;
            ReloadBuildingsTerrain();
        }

        private void LB_Items_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (LB_Items.SelectedIndex < 0)
                return;
            LoadIndex(LB_Items.SelectedIndex);

            // View location snap has changed the view. Reload everything
            LoadItemGridAcre();
            ReloadMapBackground();
        }

        private void LoadIndex(int index)
        {
            Loading = true;
            SelectedBuildingIndex = index;
            var b = Map.Buildings[index];
            NUD_BuildingType.Value = (int)b.BuildingType;
            NUD_X.Value = b.X;
            NUD_Y.Value = b.Y;
            NUD_Angle.Value = b.Angle;
            NUD_Bit.Value = b.Bit;
            NUD_Type.Value = b.Type;
            NUD_TypeArg.Value = b.TypeArg;
            NUD_UniqueID.Value = b.UniqueID;
            Loading = false;

            // -32 for relative offset on map (buildings can be placed on the exterior ocean acres)
            // -16 to put it in the center of the view
            const int shift = 48;
            var x = (b.X - shift) & 0xFFFE;
            var y = (b.Y - shift) & 0xFFFE;
            View.SetViewTo(x, y);
        }

        private void NUD_BuildingType_ValueChanged(object sender, EventArgs e)
        {
            if (Loading || sender is not NumericUpDown n)
                return;

            var b = Map.Buildings[SelectedBuildingIndex];
            if (sender == NUD_BuildingType)
                b.BuildingType = (BuildingType)n.Value;
            else if (sender == NUD_X)
                b.X = (ushort)n.Value;
            else if (sender == NUD_Y)
                b.Y = (ushort)n.Value;
            else if (sender == NUD_Angle)
                b.Angle = (byte)n.Value;
            else if (sender == NUD_Bit)
                b.Bit = (sbyte)n.Value;
            else if (sender == NUD_Type)
                b.Type = (ushort)n.Value;
            else if (sender == NUD_TypeArg)
                b.TypeArg = (byte)n.Value;
            else if (sender == NUD_UniqueID)
                b.UniqueID = (ushort)n.Value;

            LB_Items.Items[SelectedBuildingIndex] = Map.Buildings[SelectedBuildingIndex].ToString();
            ReloadBuildingsTerrain();
        }

        #endregion Buildings

        #region Acres

        private void CB_MapAcre_SelectedIndexChanged(object sender, EventArgs e)
        {
            var acre = Map.Terrain.BaseAcres[CB_MapAcre.SelectedIndex * 2];
            CB_MapAcreSelect.SelectedValue = (int)acre;

            // Jump view if available
            if (CB_Acre.Items.OfType<string>().Any(z => z == CB_MapAcre.Text))
                CB_Acre.Text = CB_MapAcre.Text;
        }

        private void CB_MapAcreSelect_SelectedValueChanged(object sender, EventArgs e)
        {
            if (Loading)
                return;

            var index = CB_MapAcre.SelectedIndex;
            var value = WinFormsUtil.GetIndex(CB_MapAcreSelect);

            var oldValue = Map.Terrain.BaseAcres[index * 2];
            if (value == oldValue)
                return;
            byte[] ValueBytes = BitConverter.GetBytes(value);
            var a = index * 2;
            Map.Terrain.BaseAcres[a] = ValueBytes[0];
            Map.Terrain.BaseAcres[a + 1] = ValueBytes[1];
            ReloadBuildingsTerrain();
        }

        private void B_DumpMapAcres_Click(object sender, EventArgs e)
        {
            if (!MapDumpHelper.DumpMapAcresAll(Map.Terrain.BaseAcres))
                return;
            ReloadBuildingsTerrain();
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void B_ImportMapAcres_Click(object sender, EventArgs e)
        {
            if (!MapDumpHelper.ImportMapAcresAll(Map.Terrain.BaseAcres))
                return;
            ReloadBuildingsTerrain();
            System.Media.SystemSounds.Asterisk.Play();
        }

        #endregion Acres

        private void B_ZeroElevation_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes != WinFormsUtil.Prompt(MessageBoxButtons.YesNo, MessageStrings.MsgTerrainSetElevation0))
                return;
            foreach (var t in Map.Terrain.Tiles)
                t.Elevation = 0;
            ReloadBuildingsTerrain();
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void B_SetAllTerrain_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes != WinFormsUtil.Prompt(MessageBoxButtons.YesNo, MessageStrings.MsgTerrainSetAll))
                return;

            var pgt = (TerrainTile)PG_TerrainTile.SelectedObject;
            bool interiorOnly = DialogResult.Yes == WinFormsUtil.Prompt(MessageBoxButtons.YesNo, MessageStrings.MsgTerrainSetAllSkipExterior);
            Map.Terrain.SetAll(pgt, interiorOnly);

            ReloadBuildingsTerrain();
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void B_SetAllRoadTiles_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes != WinFormsUtil.Prompt(MessageBoxButtons.YesNo, MessageStrings.MsgTerrainSetAll))
                return;

            var pgt = (TerrainTile)PG_TerrainTile.SelectedObject;
            bool interiorOnly = DialogResult.Yes == WinFormsUtil.Prompt(MessageBoxButtons.YesNo, MessageStrings.MsgTerrainSetAllSkipExterior);
            Map.Terrain.SetAllRoad(pgt, interiorOnly);

            ReloadBuildingsTerrain();
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void B_ClearPlacedDesigns_Click(object sender, EventArgs e)
        {
            MapManager.ClearDesignTiles(SAV);
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void B_ExportPlacedDesigns_Click(object sender, EventArgs e)
        {
            using var sfd = new SaveFileDialog
            {
                Filter = "nhmd file (*.nhmd)|*.nhmd",
                FileName = "Island MyDesignMap.nhmd",
            };
            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            string path = sfd.FileName;
            var tiles = MapManager.ExportDesignTiles(SAV);
            File.WriteAllBytes(path, tiles);
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void B_ImportPlacedDesigns_Click(object sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "nhmd file (*.nhmd)|*.nhmd",
                FileName = "Island MyDesignMap.nhmd",
            };
            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            string path = ofd.FileName;
            var tiles = File.ReadAllBytes(path);
            MapManager.ImportDesignTiles(SAV, tiles);
            System.Media.SystemSounds.Asterisk.Play();
        }

        private void Menu_Spawn_Click(object sender, EventArgs e) => new BulkSpawn(this, View.X, View.Y).ShowDialog();

        private void Menu_Bulk_Click(object sender, EventArgs e)
        {
            var editor = new BatchEditor(SpawnLayer.Tiles, ItemEdit.SetItem(new Item()));
            editor.ShowDialog();
            SpawnLayer.ClearDanglingExtensions(0, 0, SpawnLayer.MaxWidth, SpawnLayer.MaxHeight);
            LoadItemGridAcre();
        }

        private void B_TerrainBrush_Click(object sender, EventArgs e)
        {
            tbeForm = new TerrainBrushEditor(PG_TerrainTile, this);
            tbeForm.Show();
        }

        private void FieldItemEditor_FormClosed(object sender, FormClosedEventArgs e)
        {
            tbeForm?.Close();
        }
    }

    public interface IItemLayerEditor
    {
        void ReloadItems();

        ItemEditor ItemProvider { get; }
        ItemLayer SpawnLayer { get; }
    }
}