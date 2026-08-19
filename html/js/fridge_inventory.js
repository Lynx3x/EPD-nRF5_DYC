/**
 * 冰箱库存 (FridgeInventory) 模块
 *
 * 将用户从 Excel 粘贴/网页编辑的"区→分类→条目"数据，
 * 用缝合像素字体渲染成 Excel 表格样式（带合并单元格效果）的位图，
 * 发送到墨水屏。
 *
 * 数据格式（从 Excel 粘贴，TSV）：
 *  - `#区名` 开头：新区（区名），同行后面是分类名表头
 *  - 后续行：第一列空，各列数据分别归入对应分类
 *  - 括号()或（）内为备注，命中状态词则标红
 *
 * 渲染特性：
 *  - Excel 表格样式：区名 | 分类名 | 内容 三列，细线边框
 *  - 合并单元格效果：区名垂直跨区所有行居中
 *  - 列数自适应：不同区分类列数不同，自动分配列宽
 *  - 内容自动换行，条目空格分隔
 *  - 正文12px，括号备注10px，弃用8px
 */

class FridgeInventory {
  constructor() {
    // 状态词（标红）：命中括号内这些词就标红
    // 所有括号备注一律小字标红（不做关键词识别，由用户用括号自行控制）
    this.BODY_PX = 12;   // 正文
    this.NOTE_PX = 10;   // 括号备注小字
    this.zones = [];
    this.title = '冰箱库存';
  }

  // ====== Excel 粘贴解析 ======

  /**
   * 解析 TSV 文本
   * 规则：
   *  - 行首 `#` → 新区（后面是区名），同行剩余列为分类名表头
   *  - 其余行：各列数据归入当前区对应分类
   * @param {string} tsvText
   * @returns {Array<{name:string, categories:Array<{name:string, items:Array}>}>}
   */
  parseTSV(tsvText) {
    const lines = tsvText.split(/\r?\n/).filter((l) => l.trim());
    const zones = [];
    let curZone = null;

    for (const raw of lines) {
      const cells = raw.replace(/\t+$/, '').split('\t').map((c) => c.trim());
      const col0 = cells[0] || '';

      // 区名行：以 # 开头
      if (col0.startsWith('#')) {
        const zoneName = col0.replace(/^#+/, '').trim();
        curZone = { name: zoneName || '区', categories: [] };
        zones.push(curZone);
        // 同行 col1.. 是分类名表头
        for (let i = 1; i < cells.length; i++) {
          if (cells[i]) curZone.categories.push({ name: cells[i], items: [] });
        }
        continue;
      }

      // 数据行
      if (!curZone) continue;
      for (let i = 1; i < cells.length; i++) {
        const cell = cells[i];
        if (!cell) continue;
        // 若该列还没有对应分类（列数比表头多），补一个分类
        if (!curZone.categories[i - 1]) {
          curZone.categories.push({ name: `列${i}`, items: [] });
        }
        const cat = curZone.categories[i - 1];
        if (cat) cat.items.push(this.parseItem(cell));
      }
    }

    // 清理空分类
    zones.forEach((z) => {
      z.categories = z.categories.filter((c) => c.items.length > 0 || c.name);
    });
    this.zones = zones.filter((z) => z.categories.length > 0);
    return this.zones;
  }

  /**
   * 把单个单元格解析为 Item，拆分括号备注
   * 规则：只要有中英文括号，括号内一律作为备注（小字标红）
   * @param {string} text
   * @returns {{text: string, note: string|null, noteRed: boolean}}
   */
  parseItem(text) {
    // 匹配中英文括号内的备注
    const m = text.match(/^(.*?)[(（]([^)）]*)[)）]$/);
    if (m && m[1]) {
      return { text: m[1].trim(), note: m[2].trim(), noteRed: true };
    }
    return { text: text.trim(), note: null, noteRed: false };
  }

  // ====== 渲染 ======

  /**
   * 渲染库存表格到 canvas
   * 布局：三列 [区名 | 分类名 | 内容]，分类是行
   *  - 区名列：跨该区所有行（合并单元格效果，垂直居中）
   *  - 分类名列：每个分类占一行区域（合并单元格效果，垂直居中）
   *  - 内容列：条目横排 + 自动换行
   * @param {HTMLCanvasElement} canvas
   * @param {CanvasRenderingContext2D} ctx
   */
  render(canvas, ctx) {
    const W = canvas.width;
    const H = canvas.height;
    const BODY = this.BODY_PX;
    const NOTE = this.NOTE_PX;
    const fontBody = 'FusionPixel12';
    const fontNote = 'FusionPixel10';
    const black = '#000000';
    const red = '#ff0000';

    ctx.imageSmoothingEnabled = false;
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);

    // 布局：内容占满全部画面
    const margin = 8;
    const top = 8;
    const bottom = H - 8;
    const left = margin;
    const right = W - margin;
    const totalW = right - left; // 384

    // 右下角二维码（60px），"嵌"在表格右下角，只避让该区域内容
    const QR_SIZE = 60;
    const qrX = W - QR_SIZE - 8; // 二维码左缘
    const qrY = H - QR_SIZE - 8; // 二维码上缘

    // 测量字体用于动态列宽
    ctx.font = `${BODY}px "${fontBody}"`;

    // 区名列宽：取所有区名最大宽度 + 边距
    let maxZoneW = 0;
    this.zones.forEach((z) => {
      maxZoneW = Math.max(maxZoneW, ctx.measureText(z.name).width);
    });
    const zoneColW = Math.round(maxZoneW) + 12;

    // 分类名列宽：取所有分类名最大宽度 + 边距（保证"半成品(即食)"不重叠）
    let maxCatW = 0;
    this.zones.forEach((z) => {
      z.categories.forEach((c) => {
        maxCatW = Math.max(maxCatW, ctx.measureText(c.name).width);
      });
    });
    const catColW = Math.round(maxCatW) + 12;

    const catRightX = left + zoneColW + catColW; // 内容列左缘
    const itemColW = right - catRightX;          // 内容列宽（到右缘）

    // 行高
    const rowH = Math.round(BODY * 1.3); // 16

    // 计算布局（二维码感知迭代：底部行避让二维码）
    const layout = this.computeLayout(this.zones, catRightX, right, rowH, qrX, qrY,
      BODY, NOTE, fontBody, fontNote, ctx);
    const totalRows = layout.totalRows;
    const effRowH = layout.effRowH;

    // ---- 绘制（统一用绝对行号：top + 全局行号 * effRowH）----
    for (const zone of layout.zones) {
      const zoneTop = top + zone.startRow * effRowH;
      const zoneBottom = zoneTop + zone.rowCount * effRowH;
      const zoneMidY = Math.round((zoneTop + zoneBottom) / 2);

      for (const cat of zone.cats) {
        const catGlobalRow = zone.startRow + cat.startRow;
        const catTop = top + catGlobalRow * effRowH;
        const catBottom = catTop + cat.rowCount * effRowH;
        const catMidY = Math.round((catTop + catBottom) / 2);

        // 分类名（分类名列，垂直+水平居中跨其内容行）
        this.drawTextCenteredV(ctx, cat.name, catColW, BODY, fontBody, black, left + zoneColW, catMidY, black);

        // 内容行（横排 + 换行，底部进入二维码区域的行右边界收窄）
        for (let r = 0; r < cat.rows.length; r++) {
          const rowTop = top + (catGlobalRow + r) * effRowH;
          // 该行是否落入二维码垂直范围
          const inQR = (rowTop + effRowH) > qrY;
          // 该行可用右边界（进二维码则收窄到二维码左缘）
          const rowRight = inQR ? (qrX - 4) : right;
          const rowItemW = rowRight - catRightX;
          this.drawItemRow(ctx, cat.rows[r], catRightX, rowTop, rowItemW, effRowH,
            BODY, NOTE, fontBody, fontNote, black, red);
        }
      }

      // 区名（区名列，垂直+水平居中跨全区）
      this.drawTextCenteredV(ctx, zone.name, zoneColW, BODY, fontBody, black, left, zoneMidY, black);
    }

    // ---- 画表格线 ----
    this.drawGrid(ctx, layout, left, right, top, bottom, effRowH, zoneColW, catColW, black);

    // ---- 右下角二维码（嵌在表格右下角）----
    this.drawQRCode(canvas, ctx, W, H, QR_SIZE);

    return { totalRows, totalHeight: totalRows * effRowH, effRowH };
  }

  /**
   * 在右下角绘制二维码（墨水屏控制台地址），白边白底
   * 用 div 容器方式让 qrcodejs 生成（canvas 方式不显示），
   * 取内部 canvas 后 drawImage 缩放到目标大小。
   * @param {number} qrSize 二维码边长（目标绘制大小）
   * @param {number} genSize 二维码内部生成大小（建议整倍数，保证清晰）
   */
  drawQRCode(canvas, ctx, W, H, qrSize, genSize) {
    const QR_SIZE = qrSize || 72;
    const GEN_SIZE = genSize || (QR_SIZE * 2); // 内部生成大一点，缩小保证清晰
    const QR_URL = 'https://lynx3x.github.io/EPD-nRF5_DYC/html/index.html';
    const qrX = W - QR_SIZE - 8;
    const qrY = H - QR_SIZE - 8;

    try {
      // div 容器方式生成（qrcodejs 内部创建 canvas 并 append 到 div）
      const container = document.createElement('div');
      container.style.cssText = 'position:absolute;left:-9999px;top:-9999px;width:' + GEN_SIZE + 'px;height:' + GEN_SIZE + 'px;';
      document.body.appendChild(container);
      new QRCode(container, {
        text: QR_URL,
        width: GEN_SIZE,
        height: GEN_SIZE,
        colorDark: '#000000',
        colorLight: '#ffffff',
        correctLevel: QRCode.CorrectLevel.M,
      });
      // 取 qrcodejs 内部创建的 canvas
      const inner = container.querySelector('canvas');
      if (inner) {
        ctx.drawImage(inner, qrX, qrY, QR_SIZE, QR_SIZE);
      } else {
        console.warn('QRCode: div内无canvas');
      }
      // 清理临时 div
      document.body.removeChild(container);
    } catch (e) {
      console.warn('QRCode render failed:', e);
    }
  }

  /**
   * 计算布局：每个分类的内容按内容列宽横排换行，得到每个分类占几行
   */
  /**
   * 计算布局：每个分类的内容横排换行，得到每个分类占几行
   * 二维码感知迭代：底部进入二维码区域的行，右边界收窄到二维码左缘。
   * 行高受总行数影响，而二维码避让又受行高影响，用固定点迭代收敛。
   * @param {Array} zones 数据
   * @param {number} itemLeft 内容列左缘
   * @param {number} itemRightFull 内容列完整右缘
   * @param {number} rowH 初始行高
   * @param {number} qrX 二维码左缘
   * @param {number} qrY 二维码上缘
   */
  computeLayout(zones, itemLeft, itemRightFull, rowH, qrX, qrY, BODY, NOTE, fontBody, fontNote, ctx) {
    ctx.font = `${BODY}px "${fontBody}"`;

    let effRowH = rowH;
    let prevTotalRows = -1;
    let layout = null;
    let totalRows = 0;

    for (let iter = 0; iter < 5; iter++) {
      const zoneLayouts = [];
      let startRow = 0;

      for (const zone of zones) {
        const cats = [];
        let catStartRow = 0;

        for (const cat of zone.categories) {
          const rows = [];
          let curLine = [];
          let curWidth = 0;
          let lineCount = 0; // 当前分类已产行数

          const segs = cat.items.map((item) => {
            const textW = ctx.measureText(item.text).width;
            let noteSeg = null;
            if (item.note) {
              ctx.font = `${NOTE}px "${fontNote}"`;
              noteSeg = { text: `(${item.note})`, width: ctx.measureText(`(${item.note})`).width };
              ctx.font = `${BODY}px "${fontBody}"`;
            }
            return { item, textW, noteSeg };
          });

          for (const seg of segs) {
            // 当前行全局行号 → 该行 y 是否进入二维码垂直范围
            const globalRow = startRow + catStartRow + lineCount;
            const rowTop = globalRow * effRowH;
            const inQR = (rowTop + effRowH) > qrY;
            const rowRight = inQR ? (qrX - 4) : itemRightFull;
            const maxWidth = (rowRight - itemLeft) - 8;

            const segW = seg.textW + (seg.noteSeg ? seg.noteSeg.width : 0);
            const sp = curLine.length ? 8 : 0;
            if (curWidth + sp + segW > maxWidth && curLine.length) {
              rows.push(curLine);
              curLine = [];
              curWidth = 0;
              lineCount++;
            }
            const gap = curLine.length ? 8 : 0;
            curWidth += gap + segW;
            curLine.push(seg);
          }
          if (curLine.length) rows.push(curLine);
          if (!rows.length) rows.push([]);

          cats.push({ name: cat.name, rows, rowCount: rows.length, startRow: catStartRow });
          catStartRow += rows.length;
        }

        zoneLayouts.push({ name: zone.name, cats, startRow, rowCount: catStartRow });
        startRow += catStartRow;
      }

      totalRows = startRow;
      layout = { zones: zoneLayouts, totalRows };
      if (totalRows === prevTotalRows) break; // 收敛
      prevTotalRows = totalRows;

      // 更新行高（总高超可用则压缩）
      const availH = (300 - 8) - 8; // H-8 - top
      effRowH = (totalRows * rowH > availH) ? Math.max(BODY + 2, Math.floor(availH / totalRows)) : rowH;
    }

    return { zones: layout.zones, totalRows: layout.totalRows, effRowH };
  }

  /**
   * 绘制一行内容（正文12px + 括号备注10px标红 + 空格分隔）
   */
  drawItemRow(ctx, rowSegs, x, y, itemColW, rowH, BODY, NOTE, fontBody, fontNote, black, red) {
    const baseY = y + Math.round((rowH - BODY) / 2);
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = 'left';
    let cx = x + 4;

    // 测量正文字形，用于垂直居中（基线对齐到行中心）
    ctx.font = `${BODY}px "${fontBody}"`;
    const bodyM = ctx.measureText('永');
    const bodyAscent = bodyM.actualBoundingBoxAscent || BODY;
    const bodyDescent = bodyM.actualBoundingBoxDescent || 0;
    // 让正文字形中心落在行中心：baseline = 行中心 + (ascent-descent)/2
    const rowCenterY = y + rowH / 2;
    const bodyBaselineY = rowCenterY + (bodyAscent - bodyDescent) / 2;

    for (let i = 0; i < rowSegs.length; i++) {
      const seg = rowSegs[i];
      if (i > 0) {
        ctx.font = `${BODY}px "${fontBody}"`;
        ctx.fillStyle = black;
        ctx.fillText(' ', cx, bodyBaselineY);
        cx += 8;
      }
      // 正文
      ctx.font = `${BODY}px "${fontBody}"`;
      ctx.fillStyle = black;
      ctx.fillText(seg.item.text, cx, bodyBaselineY);
      cx += ctx.measureText(seg.item.text).width;
      // 备注小字（比正文小，基线下移对齐）
      if (seg.item.note) {
        const noteStr = `(${seg.item.note})`;
        ctx.font = `${NOTE}px "${fontNote}"`;
        ctx.fillStyle = seg.item.noteRed ? red : black;
        // 备注字形中心也对齐到行中心（比正文小的字，基线下移一点）
        const noteM = ctx.measureText(noteStr);
        const noteAscent = noteM.actualBoundingBoxAscent || NOTE;
        const noteDescent = noteM.actualBoundingBoxDescent || 0;
        const noteBaselineY = rowCenterY + (noteAscent - noteDescent) / 2;
        ctx.fillText(noteStr, cx, noteBaselineY);
        cx += seg.noteSeg.width;
      }
    }
  }

  /**
   * 垂直居中绘制文本（用于合并单元格的区名/分类名）
   * 正确公式：baseline 使字形中心精确落在 midY。
   * 默认水平居中（在列内居中），也可左对齐。
   * @param {number} x 所在列左缘
   * @param {number} colW 所在列宽（用于水平居中）
   */
  drawTextCenteredV(ctx, text, colW, fontPx, fontFamily, color, x, midY, drawColor, hAlignCenter = true) {
    ctx.font = `${fontPx}px "${fontFamily}"`;
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = hAlignCenter ? 'center' : 'left';
    ctx.fillStyle = drawColor;
    const m = ctx.measureText(text);
    const ascent = m.actualBoundingBoxAscent || fontPx;
    const descent = m.actualBoundingBoxDescent || 0;
    const baselineY = midY + (ascent - descent) / 2;
    // 水平居中：中心 = x + colW/2；左对齐：x + 6
    const textX = hAlignCenter ? (x + colW / 2) : (x + 6);
    ctx.fillText(text, textX, baselineY);
  }

  /**
   * 画表格网格线（三列：区名 | 分类名 | 内容）
   */
  drawGrid(ctx, layout, left, right, top, bottom, rowH, zoneColW, catColW, black) {
    ctx.strokeStyle = black;
    ctx.lineWidth = 1;

    const zoneColX = left + zoneColW;      // 区名|分类名 分隔
    const catColX = zoneColX + catColW;    // 分类名|内容 分隔
    const totalRows = layout.totalRows;
    const totalH = totalRows * rowH;

    // 外框（整表）
    ctx.strokeRect(left, top, right - left, totalH);

    // 区名|分类名 竖线（贯穿全区）
    ctx.beginPath();
    ctx.moveTo(zoneColX, top); ctx.lineTo(zoneColX, top + totalH);
    ctx.stroke();

    // 分类名|内容 竖线（贯穿全区）
    ctx.beginPath();
    ctx.moveTo(catColX, top); ctx.lineTo(catColX, top + totalH);
    ctx.stroke();

    // 区之间横线（外框内）
    for (const z of layout.zones) {
      const zBottom = top + (z.startRow + z.rowCount) * rowH;
      if (zBottom < top + totalH) {
        ctx.beginPath();
        ctx.moveTo(left, zBottom); ctx.lineTo(right, zBottom);
        ctx.stroke();
      }
    }

    // 分类之间的横线（分类名|内容 区域内）
    for (const z of layout.zones) {
      for (const c of z.cats) {
        const cBottom = top + (z.startRow + c.startRow + c.rowCount) * rowH;
        if (cBottom < top + (z.startRow + z.rowCount) * rowH) { // 分类非本区最后一行
          ctx.beginPath();
          ctx.moveTo(zoneColX, cBottom); ctx.lineTo(right, cBottom);
          ctx.stroke();
        }
        // 分类内部内容换行线
        for (let r = 1; r < c.rowCount; r++) {
          const y = top + (z.startRow + c.startRow + r) * rowH;
          ctx.beginPath();
          ctx.moveTo(catColX, y); ctx.lineTo(right, y);
          ctx.stroke();
        }
      }
    }
  }
  /**
   * 初始化事件绑定
   */
  init() {
    const self = this;

    // 生成预览
    document.getElementById('fridge-parse-btn').addEventListener('click', () => {
      self.refreshPreview();
    });

    // 发送到屏幕
    document.getElementById('fridge-send-btn').addEventListener('click', () => {
      self.sendToScreen();
    });

    // 初始预览（空白状态）
    this.refreshPreview();

    // 预加载像素字体
    const fontLoadPromises = [];
    ['FusionPixel10', 'FusionPixel12'].forEach((fam) => {
      [10, 12].forEach((px) => {
        fontLoadPromises.push(document.fonts.load(`${px}px "${fam}"`, '冰箱库存'));
      });
    });
    Promise.allSettled(fontLoadPromises).then(() => {
      self.refreshPreview();
    });
  }

  /**
   * 从输入框读取文本并解析
   */
  loadFromInput() {
    const text = document.getElementById('fridge-input').value;
    if (!text.trim()) return null;
    this.parseTSV(text);
    return this.zones;
  }

  /**
   * 刷新网页预览
   */
  refreshPreview() {
    const canvas = document.getElementById('fridge-preview');
    const ctx = canvas.getContext('2d');
    canvas.width = 400;
    canvas.height = 300;
    const zones = this.loadFromInput();
    if (!zones || zones.length === 0) {
      // 无数据：显示提示
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, 400, 300);
      ctx.fillStyle = '#999999';
      ctx.font = '12px "FusionPixel12"';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText('请粘贴冰箱数据', 200, 150);
      this.updateStats(null);
      return;
    }
    this.render(canvas, ctx);
    this.updateStats(zones);
  }

  /**
   * 更新统计信息
   */
  updateStats(zones) {
    const el = document.getElementById('fridge-stats');
    if (!el) return;
    if (!zones) {
      el.textContent = '';
      return;
    }
    let items = 0;
    let cats = 0;
    zones.forEach((z) => {
      cats += z.categories.length;
      z.categories.forEach((c) => (items += c.items.length));
    });
    el.textContent = `${zones.length}区 ${cats}分类 ${items}条目`;
  }

  /**
   * 发送到墨水屏
   */
  async sendToScreen() {
    if (!epdCharacteristic) {
      addLog("未连接设备，请先连接蓝牙");
      showToast("请先连接设备");
      return;
    }

    const zones = this.loadFromInput();
    if (!zones || zones.length === 0) {
      addLog("请先粘贴冰箱数据");
      showToast("请先粘贴冰箱数据");
      return;
    }

    // 确保主画布尺寸
    const canvasSize = document.getElementById('canvasSize').value;
    if (canvasSize !== '4.2_400_300') {
      const sizeSelect = document.getElementById('canvasSize');
      sizeSelect.value = '4.2_400_300';
      updateCanvasSize();
    }
    document.getElementById('ditherMode').value = 'threeColor';

    // 渲染到主画布
    this.render(canvas, ctx);

    startTime = new Date().getTime();
    const status = document.getElementById("status");
    status.parentElement.style.display = "block";

    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const processedData = processImageData(imageData, 'threeColor');

    updateButtonStatus(true);

    let useCRC = (appVersion >= 0x20) && (typeof BleTransfer !== 'undefined');
    const transferFn = useCRC ? writeImageCRC : writeImage;
    if (useCRC) addLog("使用CRC校验传输模式");

    const halfLength = Math.floor(processedData.length / 2);
    const blackWhiteData = processedData.slice(0, halfLength);
    const redWhiteData = processedData.slice(halfLength);

    await transferFn(blackWhiteData, 'bw');
    await transferFn(redWhiteData, 'red');

    const refreshOk = await write(EpdCmd.REFRESH);
    updateButtonStatus();

    if (refreshOk) {
      const sendTime = (new Date().getTime() - startTime) / 1000.0;
      addLog(`发送完成！耗时: ${sendTime}s`);
      setStatus(`发送完成！耗时: ${sendTime}s`);
      addLog("屏幕刷新完成前请不要操作。");
    } else {
      addLog('刷新指令发送失败，请重试。');
    }
    status.parentElement.style.display = "none";
  }
}

// 全局实例
let fridgeInventory = null;
