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
    // 所有括号备注一律小字标红（不做关键词识别，由用户用括号自行控制）
    this.BODY_PX = 12;   // 正文
    this.NOTE_PX = 10;   // 括号备注小字
    this.zones = [];
    this.title = '冰箱库存';
    this.style = 'panel'; // 渲染风格: 'table'(表格) / 'panel'(信息面板) / 'card'(卡片)
    this.orientation = 'landscape'; // 方向: 'landscape'(横屏400x300) / 'portrait'(竖屏300x400)
    this.screen = '4.2'; // 屏幕尺寸: '4.2'(400x300) / '7.5'(800x480)
    // 正文字体/字号（可独立选择，localStorage保存）
    this.body_font = 'FusionPixel12';
    this.body_size = 12;
    // 备注字体/字号
    this.note_font = 'FusionPixel12';
    this.note_size = 10;
    // 字号满幅测试状态
    this.testSize = 10; // 当前测试字号 (8/10/12)
    // 满幅测试用的文本（去掉标点空白后逐字填屏）
    this.testText = '我在漫宿高处的房间居住了一小段时间，而后由三尖之门返回，我的躯体在灰烬中跃跃欲出。我已毛发全无，如大理石般千古不朽，铸炉之火仍在我内里燃烧。我具有〖足以塑形的气力〗。我再不会变老。也许我将反叛。也许有朝一日，我将升得更高。狂暴之门是一道旧伤口，不会轻易开启，但它还是再度开启了。连石源之神都使用过此门，那么在祂们前有多少神用过呢？辉光是一个疑问，我们所有人，作为燃料供给我们的沐光明者的见证者与我，都以肯定作答。地撒银光，晚风凉爽；夜晚的寂静引来言语，如暗色丝绒引来手指抚弄。蛾在黑暗中寻见光，我亦如是。一种似有似无的惆怅染上了我。每当灯之准则放光时，我便可以多记起一些残篇知识。脑海中的灰尘一扫而空，至少暂时如此。我在半夜惊醒。我的手掌，我的脸上都渗出汗珠。慢慢地，我松开紧攥的拳头。我又滑落回睡梦中。但某种东西留下了，好像房间角落的蛛网。沙皇时代我曾在泰加林狩猎，那里的深处有一个虎口形状的湖。也许我的大敌不会去那里找我，如果他去了，那就是时候奋起反抗了。我们交谈时，万查妮一寸一寸靠近。她握住我手时，她的手指传来机械般的力道，她的皮肤有着余烬般的热力。“我不是光，”她警告我说，“我们应当对彼此宽容。”埃尔里奇已基本能管住自己在谈话中不用刀剔牙了，但他还不能把刀一直放在靴子里不拔出来。埃尔里奇正在耐心等待再次捅人刀子的时机。芮妮拉在三个不相邻的辖区里被判三次无罪释放，她很可能永远不会因为任何事获罪。悲伤时，她便闭口不言，但她时常悲伤。我如今知道了该做出什么姿势，手该恰到好处地放在何处，慢慢领她那可爱的颈脖靠上我肩。我们的生活是幸福的，我们看上去是一对般配的恋人；当然了，她的哥哥时而来访——但那是因为我出门旅行时她寂寞，需要陪伴。我的脑子装满了的知识，快要撑裂颅骨。“只欲求而不行动者于世有害。”“与其哺育只欲求而不行动的婴孩，不如将其尽早扼杀在摇篮中。”“我们指引前路，我们照明驱暗，我们无有怜悯之心。”秘密温柔，夜晚更加温柔。大海会讲述，但是聆听不是何时都是明智的。灰暗的海水每天早晨都冲击着岩礁，又整夜整夜地对我发出低吼。每一天的日出都来得比前一日更早。风也变得更加温柔。春天快要来了。';
  }

  // ====== Excel 粘贴解析 ======

  /**
   * 解析 TSV 文本
   * 规则：
   *  - 行首 `#` → 新区（后面是区名）
   *  - 其余行：第一列 = 分类名，后面所有非空列 = 该分类的内容（横排）
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
        continue;
      }

      // 分类行：第一列=分类名，后面=内容
      if (!curZone) continue;
      const catName = col0;
      if (!catName) continue;
      const cat = { name: catName, items: [] };
      curZone.categories.push(cat);
      for (let i = 1; i < cells.length; i++) {
        const cell = cells[i];
        if (!cell) continue;
        cat.items.push(this.parseItem(cell));
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
    // 中文括号统一转英文括号（更小不占地方）
    const normalized = text.replace(/（/g, '(').replace(/）/g, ')');
    const trimmed = normalized.trim();

    // 末尾备注: xxx(备注)
    let m = trimmed.match(/^(.*?)[(]([^)]*)[)]$/);
    if (m && m[1].trim()) {
      return { text: m[1].trim(), note: m[2].trim(), noteRed: true };
    }
    // 前置备注: (备注)xxx
    m = trimmed.match(/^[(]([^)]*)[)](.*)$/);
    if (m && m[2].trim()) {
      return { text: m[2].trim(), note: m[1].trim(), noteRed: true };
    }
    return { text: trimmed, note: null, noteRed: false };
  }

  // ====== 渲染 ======

  /**
   * 渲染入口：根据 this.style 选择风格
   *  - 'table': 原表格风格（区名|分类名|内容 三列，带表格线）
   *  - 'panel': 信息面板风格（区名24px大字，无表格线，留白分区）
   *  - 'card': 印刷品/分栏卡片风格（每区一个带边框卡片，区名在卡片顶部）
   */
  render(canvas, ctx) {
    if (this.style === 'table') {
      return this.renderStyleTable(canvas, ctx);
    }
    if (this.style === 'card') {
      return this.renderStyleCard(canvas, ctx);
    }
    return this.renderStylePanel(canvas, ctx);
  }

  /**
   * 方向 A：信息面板风格
   *  - 区名 24px 大字（清晰整数倍）
   *  - 无表格线，靠留白分区
   *  - 分类名 + 条目横排，层级缩进
   *  - 红色仅备注
   * @param {HTMLCanvasElement} canvas
   * @param {CanvasRenderingContext2D} ctx
   */
  renderStylePanel(canvas, ctx) {
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

    // 布局
    const margin = 8;
    const top = 4;        // 顶部留白压缩（给内容省空间）
    const bottom = H - 6; // 底部留白
    const left = margin;
    const right = W - margin;

    // 右下角二维码（60px）
    const QR_SIZE = 60;
    const qrX = W - QR_SIZE - 8;
    const qrY = H - QR_SIZE - 8;

    // 区名 20px（10 字体放大 2 倍，清晰），正文 12px
    const ZONE_PX = 20;
    const zoneFont = 'FusionPixel10';
    const rowH = Math.round(BODY * 1.3); // 16
    const zoneRowH = ZONE_PX + 4; // 24（20px字 + 少量行距，压缩顶部区名行）

    // 分类名列宽 = 最长分类名(含括号)宽 + 极小间距
    ctx.font = `${BODY}px "${fontBody}"`;
    let maxCatW = 0;
    this.zones.forEach((z) => {
      z.categories.forEach((c) => {
        // 分类名含〔〕括号，计算完整宽度
        const label = `〔${c.name}〕`;
        maxCatW = Math.max(maxCatW, ctx.measureText(label).width);
      });
    });
    const catBoxW = Math.round(maxCatW) + 4; // 分类名列宽（极小间距）
    const catBoxX = left; // 分类名左缘（与区名对齐，左侧）
    const itemX = catBoxX + catBoxW; // 内容列起点（分类名右侧）

    // 布局：每区 = 区名行 + 内容行
    let totalHeight = 0;
    const zoneLayouts = [];
    for (const zone of this.zones) {
      const cats = [];
      let catStartRow = 0;
      for (const cat of zone.categories) {
        // 内容列宽 = right - itemX（二维码避让在此列内处理）
        const itemColFullW = right - itemX;
        const rows = [];
        let curLine = [];
        let curWidth = 0;
        const maxWidth = itemColFullW - 8;
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
          const segW = seg.textW + (seg.noteSeg ? seg.noteSeg.width : 0);
          const sp = curLine.length ? 8 : 0;
          // 超宽就换行（即使当前行空也换，避免单个超长条目累积超框）
          if (curLine.length && curWidth + sp + segW > maxWidth) {
            rows.push(curLine);
            curLine = [];
            curWidth = 0;
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
      zoneLayouts.push({ name: zone.name, cats, contentRows: catStartRow });
      totalHeight += zoneRowH + catStartRow * rowH;
    }

    // 超屏处理：行高固定 16px（备注清晰），压缩顶部/区留白，仍放不下则省略底部内容
    const availH = bottom - top;
    const ZONE_GAP = 3; // 区之间留白（压缩）
    const fixedCost = zoneLayouts.length * (zoneRowH + ZONE_GAP);
    const contentRowsTotal = zoneLayouts.reduce((a, z) => a + z.contentRows, 0);
    // 行高固定 16px，不再压缩（否则备注发糊）
    let effRowH = rowH; // 16
    // 如果 16px 行高放不下，内容会在绘制时被省略（break），保证不溢出
    // 同时压缩顶部留白已在 top 定义处处理

    // ---- 绘制 ----
    let y = top;
    for (const zone of zoneLayouts) {
      // 区名（20px，左对齐）
      ctx.font = `${ZONE_PX}px "${zoneFont}"`;
      ctx.textBaseline = 'alphabetic';
      ctx.textAlign = 'left';
      ctx.fillStyle = black;
      const zm = ctx.measureText(zone.name);
      const zAscent = zm.actualBoundingBoxAscent || ZONE_PX;
      const zDescent = zm.actualBoundingBoxDescent || 0;
      // 底部保护：区名行也需在可用范围内
      if (y + zoneRowH <= bottom) {
        ctx.fillText(zone.name, left, y + zoneRowH / 2 + (zAscent - zDescent) / 2);
      }
      y += zoneRowH;

      for (const cat of zone.cats) {
        const catTop = y;
        // 分类名用〔〕括号包裹（如 〔饮料〕），右对齐到内容列起点，紧贴内容
        ctx.font = `${BODY}px "${fontBody}"`;
        ctx.textBaseline = 'alphabetic';
        ctx.textAlign = 'right';
        ctx.fillStyle = black;
        const catLabel = `〔${cat.name}〕`;
        const cm = ctx.measureText(catLabel);
        const cAscent = cm.actualBoundingBoxAscent || BODY;
        const cDescent = cm.actualBoundingBoxDescent || 0;
        // 分类名画在第一行，垂直居中，右对齐到 itemX-2（紧贴内容）
        const catFirstRowCenter = catTop + effRowH / 2;
        ctx.fillText(catLabel, itemX - 2, catFirstRowCenter + (cAscent - cDescent) / 2);
        ctx.textAlign = 'left';

        // 内容行：从分类名右侧（固定列宽 catBoxW）开始，不侵占分类名
        for (let r = 0; r < cat.rows.length; r++) {
          const rowTop = catTop + r * effRowH;
          // 底部保护：超出可用高度则不绘制（避免内容跑到屏幕外/被切割）
          if (rowTop + effRowH > bottom) break;
          const inQR = (rowTop + effRowH) > qrY;
          const rowRight = inQR ? (qrX - 4) : right;
          const rowItemW = rowRight - itemX;
          this.drawItemRow(ctx, cat.rows[r], itemX, rowTop, rowItemW, effRowH,
            BODY, NOTE, fontBody, fontNote, black, red);
        }
        y = catTop + cat.rowCount * effRowH;
      }

      // 区之间留白
      y += ZONE_GAP;
    }

    // ---- 右下角二维码 ----
    this.drawQRCode(canvas, ctx, W, H, QR_SIZE);

    return { totalHeight: y - top, effRowH };
  }

  /**
   * 方向 B：印刷品/分栏卡片风格
   *  - 每个区一个带边框的卡片块
   *  - 区名在卡片顶部（大字号 + 底部装饰线）
   *  - 分类名用〔〕括号，内容横排
   *  - 红色仅备注
   */
  renderStyleCard(canvas, ctx) {
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

    // 布局：顶部设备名行（占一整行），卡片内容从下方开始
    const margin = 6;
    // 缩放倍数 = 正文字号 / 基准12px，所有布局/行高/间距随字号联动，避免字重叠
    const BODY_PX = this.body_size || BODY;     // 正文字号
    const SCALE = BODY_PX / 12;
    const DEVICE_NAME_ROW_H = 20 * SCALE;   // 设备名行高
    const top = margin + DEVICE_NAME_ROW_H; // 卡片内容起始y（设备名行之下）
    const bottom = H - 6 * SCALE;
    const left = margin;
    const right = W - margin;

    // 右下角二维码（按缩放：4.2寸60px，7.5寸120px）
    const QR_SCALE = SCALE;
    const QR_SIZE = 60 * QR_SCALE;
    const qrX = W - QR_SIZE - 8 * QR_SCALE;
    const qrY = H - QR_SIZE - 8 * QR_SCALE;
    // 正文/备注字体与字号：用户独立选择（setFont 设置，localStorage保存）
    const bodyFontFam = this.body_font || fontBody; // 正文字体族
    const NOTE_PX = this.note_size || NOTE;     // 备注字号
    const noteFontFam = this.note_font || fontNote; // 备注字体族
    const ZONE_PX = BODY_PX;        // 区名与正文同字号
    const zoneFont = bodyFontFam;
    const rowH = Math.round(16 * SCALE);  // 内容行高随字号缩放（避免字重叠）
    const effRowH = rowH;

    // 分类标记符号
    const CAT_MARK = '■';

    // 卡片内缩进
    const cardPad = 4 * SCALE;
    const ZONE_GAP = 6 * SCALE;
    const zoneRowH = BODY_PX;

    // 分类名起点（每个分类内容紧贴各自分类名右侧，不统一按最长算）
    ctx.font = `${BODY_PX}px "${bodyFontFam}"`;
    const catBoxX = left + cardPad; // 分类名起点

    // 布局：每区卡片 = 区名行 + (分类名与内容同行)*
    const zoneLayouts = [];
    let totalHeight = 0;
    let heightSoFar = 0; // 前面所有卡片的累计高度（用于二维码避让y计算）
    for (const zone of this.zones) {
      const cats = [];
      let catStartRow = 0;
      // 本区内容第一行的全局y基准 = 累计高度 + 区名行高
      const zoneRowGlobal = heightSoFar + zoneRowH;
      for (const cat of zone.categories) {
        // 该分类自己的内容起点 = 分类名(■名称)右缘 + 少量
        const catLabel = `${CAT_MARK} ${cat.name}`;
        const catLabelW = ctx.measureText(catLabel).width;
        const itemX_cat = catBoxX + catLabelW + 6;
        // 内容列宽 = 卡片右边 - 该分类内容起点
        const maxWidth = (right - cardPad) - itemX_cat - 4;
        const rows = [];
        let curLine = [];
        let curWidth = 0;
        // 当前行全局y（绝对像素，用于二维码避让判断）
        // = 本区内容基准y(zoneRowGlobal 已是像素) + 本分类之前分类的行数(catStartRow)*effRowH
        let rowY = zoneRowGlobal + catStartRow * effRowH; // 该分类第一行内容的绝对y
        const segs = cat.items.map((item) => {
          const textW = ctx.measureText(item.text).width;
          let noteSeg = null;
          if (item.note) {
            ctx.font = `${NOTE_PX}px "${noteFontFam}"`;
            noteSeg = { text: `(${item.note})`, width: ctx.measureText(`(${item.note})`).width };
            ctx.font = `${BODY_PX}px "${bodyFontFam}"`;
          }
          return { item, textW, noteSeg };
        });
        for (const seg of segs) {
          const segW = seg.textW + (seg.noteSeg ? seg.noteSeg.width : 0);
          const sp = curLine.length ? 8 : 0;
          // 当前行的可用宽度：该行y是否进入二维码垂直范围
          const inQR = (rowY + effRowH) > qrY;
          const rowMaxW = inQR ? ((qrX - 4) - itemX_cat) : maxWidth;
          if (curLine.length && curWidth + sp + segW > rowMaxW) {
            rows.push(curLine);
            curLine = [];
            curWidth = 0;
            rowY += effRowH;
          }
          const gap = curLine.length ? 8 : 0;
          curWidth += gap + segW;
          curLine.push(seg);
        }
        if (curLine.length) rows.push(curLine);
        if (!rows.length) rows.push([]);
        // 分类名与内容同行：占 rows.length 行，记录该分类内容起点 itemX_cat
        cats.push({ name: cat.name, rows, rowCount: rows.length, itemX: itemX_cat, startRow: catStartRow });
        catStartRow += rows.length;
      }
      zoneLayouts.push({ name: zone.name, cats, contentRows: catStartRow });
      totalHeight += zoneRowH + catStartRow * effRowH + cardPad;
      // 更新累计高度（供下一区二维码避让y计算）
      heightSoFar += zoneRowH + catStartRow * effRowH + cardPad + ZONE_GAP;
    }
    totalHeight += zoneLayouts.length * ZONE_GAP;

    // ---- 绘制 ----
    // 顶部设备名行（占一整行）
    this.drawDeviceNameRow(ctx, W, left, right, margin, DEVICE_NAME_ROW_H, SCALE);
    let y = top;
    for (const zone of zoneLayouts) {
      const cardTop = y;
      const cardHeight = zoneRowH + zone.contentRows * effRowH + cardPad;

      // 区名（12px，骑在框线上）
      ctx.font = `${ZONE_PX}px "${zoneFont}"`;
      const zoneLabelW = ctx.measureText(zone.name).width;
      const zm2 = ctx.measureText(zone.name);
      const zAscent2 = zm2.actualBoundingBoxAscent || ZONE_PX;
      const zDescent2 = zm2.actualBoundingBoxDescent || 0;
      const lineBreakStart = left + 10;
      const zoneLabelX = lineBreakStart + 2;
      const zoneLabelBaseline = cardTop + (zAscent2 - zDescent2) / 2;
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(zoneLabelX - 2, cardTop - zAscent2, zoneLabelW + 4, zAscent2 + zDescent2);
      ctx.textBaseline = 'alphabetic';
      ctx.textAlign = 'left';
      ctx.fillStyle = black;
      ctx.fillText(zone.name, zoneLabelX, zoneLabelBaseline);

      // 卡片边框（单线，上边线在区名处断开）
      const zoneLeftX = zoneLabelX + zoneLabelW + 4;
      ctx.strokeStyle = black;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(left, cardTop); ctx.lineTo(lineBreakStart, cardTop);
      ctx.moveTo(zoneLeftX, cardTop); ctx.lineTo(right, cardTop);
      ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(left, cardTop); ctx.lineTo(left, cardTop + cardHeight);
      ctx.lineTo(right, cardTop + cardHeight);
      ctx.lineTo(right, cardTop);
      ctx.stroke();

      // 分类名与内容同行
      let yContent = cardTop + zoneRowH - 2;
      for (const cat of zone.cats) {
        const catTop = yContent;
        // 分类名（第一行左侧，固定列宽）
        ctx.font = `${BODY_PX}px "${bodyFontFam}"`;
        ctx.textBaseline = 'alphabetic';
        ctx.textAlign = 'left';
        ctx.fillStyle = black;
        const catLabel = `${CAT_MARK} ${cat.name}`;
        const cm = ctx.measureText(catLabel);
        const cAscent = cm.actualBoundingBoxAscent || BODY_PX;
        const cDescent = cm.actualBoundingBoxDescent || 0;
        const catFirstRowCenter = catTop + effRowH / 2;
        ctx.fillText(catLabel, catBoxX, catFirstRowCenter + (cAscent - cDescent) / 2);

        // 内容行（从该分类名右侧同行开始，换行也不侵占分类名）
        for (let r = 0; r < cat.rows.length; r++) {
          const rowTop = catTop + r * effRowH;
          if (rowTop + effRowH > bottom) break;
          const inQR = (rowTop + effRowH) > qrY;
          const rowRight = inQR ? (qrX - 4) : (right - cardPad);
          const rowItemW = rowRight - cat.itemX;
          this.drawItemRow(ctx, cat.rows[r], cat.itemX, rowTop, rowItemW, effRowH,
            BODY_PX, NOTE_PX, bodyFontFam, noteFontFam, black, red);
        }
        yContent = catTop + cat.rowCount * effRowH;
      }
      y = cardTop + cardHeight + ZONE_GAP;
    }

    // ---- 底部字体备注（设备上可见当前字体设置）----
    this.drawFontInfo(canvas, ctx, W, H, SCALE);

    // ---- 右下角二维码 ----
    this.drawQRCode(canvas, ctx, W, H, QR_SIZE);

    return { totalHeight: y - top, effRowH };
  }

  /**
   * 在底部绘制当前字体设置备注（设备上可见）
   * 位置：左下角，避免与右下角二维码冲突
   * 用观致8px字体 + 黑色（墨水屏只能显示黑白红，浅灰会失效）
   */
  drawFontInfo(canvas, ctx, W, H, scale) {
    const bodyF = this.body_font || 'FusionPixel12';
    const noteF = this.note_font || 'FusionPixel12';
    const bodyS = this.body_size || 12;
    const noteS = this.note_size || 10;
    const names = {
      FusionPixel12: 'Fusion12', FusionPixel10: 'Fusion10', FusionPixel: 'Fusion8',
      WQY12: '文泉驿12', WQY16: '文泉驿16', Cubic11: '立方体11', Unifont: 'Unifont',
    };
    // 字体名-字号 用 - 连接，避免 Fusion1212 这种歧义；观致8px 标注为提示字
    const bodyStr = `${names[bodyF] || bodyF}-${bodyS}`;
    const noteStr = `${names[noteF] || noteF}-${noteS}`;
    const text = `正文 ${bodyStr} / 备注 ${noteStr} / 系统字 观致-8`;
    // 观致8px 必须用 8 的整数倍字号（8/16/24...），否则非整数倍会发糊
    const guanScale = Math.max(1, Math.round(scale));
    const px = 8 * guanScale;   // 8px 或 16px
    const fontFam = 'GuanZhi8'; // 观致8px系统提示小字
    ctx.font = `${px}px "${fontFam}"`;
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = 'left';
    ctx.fillStyle = '#000000'; // 黑色（墨水屏可显示）
    // 左下角，距边 8px
    ctx.fillText(text, 8, H - 8);
  }

  renderStyleTable(canvas, ctx) {
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

    // 右下角二维码（60px），"嵌"在表格右下角
    const QR_SIZE = 60;
    const qrX = W - QR_SIZE - 8;
    const qrY = H - QR_SIZE - 8;

    // 测量字体用于动态列宽
    ctx.font = `${BODY}px "${fontBody}"`;

    // 区名列宽：取所有区名最大宽度 + 边距
    let maxZoneW = 0;
    this.zones.forEach((z) => {
      maxZoneW = Math.max(maxZoneW, ctx.measureText(z.name).width);
    });
    const zoneColW = Math.round(maxZoneW) + 12;

    // 分类名列宽：取所有分类名最大宽度 + 边距
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
          const inQR = (rowTop + effRowH) > qrY;
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

    // ---- 右下角二维码 ----
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
   * 绘制像素风装饰边框（双线 + 四角L形角花）
   * 像素放置游戏UI风格，适合墨水屏黑白+少量红
   * @param {CanvasRenderingContext2D} ctx
   * @param {number} x 外框左上x
   * @param {number} y 外框左上y
   * @param {number} w 外框宽
   * @param {number} h 外框高
   * @param {string} color 主色（默认黑）
   * @param {string} accentColor 角花强调色（默认同主色，可传红）
   */
  drawPixelBorder(ctx, x, y, w, h, color = '#000000', accentColor = null) {
    const c = color;
    const a = accentColor || color;
    const gap = 2; // 双线间距

    ctx.fillStyle = c;
    // 外框线（4条边，1px）
    ctx.fillRect(x, y, w, 1);            // 上
    ctx.fillRect(x, y + h - 1, w, 1);    // 下
    ctx.fillRect(x, y, 1, h);            // 左
    ctx.fillRect(x + w - 1, y, 1, h);    // 右

    // 内框线（间距 gap）
    ctx.fillRect(x + gap, y + gap, w - gap * 2, 1);            // 上
    ctx.fillRect(x + gap, y + h - gap - 1, w - gap * 2, 1);    // 下
    ctx.fillRect(x + gap, y + gap, 1, h - gap * 2);            // 左
    ctx.fillRect(x + w - gap - 1, y + gap, 1, h - gap * 2);    // 右

    // 四角 L 形角花（用强调色，内外框在角上相连）
    ctx.fillStyle = a;
    // 左上角
    ctx.fillRect(x, y, 3, 1);
    ctx.fillRect(x, y, 1, 3);
    ctx.fillRect(x + gap, y + gap, 2, 1);
    ctx.fillRect(x + gap, y + gap, 1, 2);
    // 右上角
    ctx.fillRect(x + w - 3, y, 3, 1);
    ctx.fillRect(x + w - 1, y, 1, 3);
    ctx.fillRect(x + w - gap - 2, y + gap, 2, 1);
    ctx.fillRect(x + w - gap - 1, y + gap, 1, 2);
    // 左下角
    ctx.fillRect(x, y + h - 1, 3, 1);
    ctx.fillRect(x, y + h - 3, 1, 3);
    ctx.fillRect(x + gap, y + h - gap - 1, 2, 1);
    ctx.fillRect(x + gap, y + h - gap - 2, 1, 2);
    // 右下角
    ctx.fillRect(x + w - 3, y + h - 1, 3, 1);
    ctx.fillRect(x + w - 1, y + h - 3, 1, 3);
    ctx.fillRect(x + w - gap - 2, y + h - gap - 1, 2, 1);
    ctx.fillRect(x + w - gap - 1, y + h - gap - 2, 1, 2);
  }

  /**
   * 绘制顶部设备名行（占一整行，右对齐设备名 + 底部分隔线）
   * 设备名从全局 bleDevice.name 获取，未连接则显示空行（仍占位）
   */
  drawDeviceNameRow(ctx, W, left, right, margin, rowH, scale) {
    let name = '';
    try {
      if (typeof bleDevice !== 'undefined' && bleDevice && bleDevice.name) {
        name = bleDevice.name;
      }
    } catch (e) { /* 忽略 */ }
    const px = 12 * (scale || 1); // 设备名字号（按屏幕缩放）
    ctx.font = `${px}px "FusionPixel12"`;
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = 'right';
    ctx.fillStyle = '#000000';
    // 设备名垂直居中在该行
    const m = ctx.measureText(name || '设备');
    const ascent = m.actualBoundingBoxAscent || px;
    const descent = m.actualBoundingBoxDescent || 0;
    const centerY = margin + rowH / 2;
    if (name) {
      ctx.fillText(name, right, centerY + (ascent - descent) / 2);
    }
    ctx.textAlign = 'left';
    // 底部分隔线（设备名行和内容区分开）
    ctx.fillStyle = '#000000';
    ctx.fillRect(left, margin + rowH, right - left, 1);
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
   * 切换渲染风格并刷新预览
   * @param {string} style 'panel' 或 'table'
   */
  setStyle(style) {
    if (style === 'table' || style === 'card') {
      this.style = style;
    } else {
      this.style = 'panel';
    }
    this.refreshPreview();
  }

  /**
   * 切换横屏/竖屏并刷新预览
   * 横屏 400x300，竖屏 300x400
   */
  setOrientation(orientation) {
    this.orientation = (orientation === 'portrait') ? 'portrait' : 'landscape';
    this.refreshPreview();
  }

  /**
   * 切换屏幕尺寸并刷新预览
   * '4.2' → 400x300，'7.5' → 800x480（内容放大）
   */
  setScreen(screen) {
    this.screen = (screen === '7.5') ? '7.5' : '4.2';
    this.refreshPreview();
  }

  /**
   * 设置正文/备注的字体或字号，刷新预览并保存
   * @param {string} which 'body' / 'note'
   * @param {string} field 'font' / 'size'
   * @param {string} value 字体族名 或 字号
   */
  setFont(which, field, value) {
    const key = which + '_' + field; // body_font / body_size / note_font / note_size
    if (field === 'size') {
      const v = parseFloat(value);
      this[key] = (isFinite(v) && v > 0) ? v : (which === 'note' ? 10 : 12);
    } else {
      this[key] = value;
    }
    try {
      localStorage.setItem('fridge_' + key, String(this[key]));
    } catch (e) { /* 忽略 */ }
    this.refreshPreview();
    this.updateFontInfo();
  }

  /**
   * 更新页面下方的当前字体摘要
   */
  updateFontInfo() {
    const el = document.getElementById('fridge-fontinfo');
    if (!el) return;
    const fmt = (fam, px) => {
      // 字体族友好名
      const names = {
        FusionPixel12: 'FusionPixel12', FusionPixel10: 'FusionPixel10', FusionPixel: 'FusionPixel8',
        WQY12: '文泉驿12', WQY16: '文泉驿16', Cubic11: '立方体11', Unifont: 'Unifont',
      };
      return (names[fam] || fam) + ' ' + px + 'px';
    };
    el.textContent = '当前字体: 正文 ' + fmt(this.body_font, this.body_size) + ' · 备注 ' + fmt(this.note_font, this.note_size);
  }

  /**
   * 初始化事件绑定
   */
  init() {
    const self = this;

    // 恢复上次输入的冰箱数据（localStorage 缓存）
    const inputEl = document.getElementById('fridge-input');
    try {
      const saved = localStorage.getItem('fridge_input_data');
      if (saved && inputEl) {
        inputEl.value = saved;
      }
    } catch (e) { /* localStorage 不可用则忽略 */ }

    // 输入时自动保存（缓存），并自动刷新预览
    if (inputEl) {
      inputEl.addEventListener('input', () => {
        try {
          localStorage.setItem('fridge_input_data', inputEl.value);
        } catch (e) { /* 忽略 */ }
        self.refreshPreview();
      });
    }

    // 恢复用户保存的字号缩放
    // 恢复用户保存的正文/备注字体字号
    try {
      const cfg = [
        ['body_font', 'fridge-body-font'],
        ['body_size', 'fridge-body-size'],
        ['note_font', 'fridge-note-font'],
        ['note_size', 'fridge-note-size'],
      ];
      cfg.forEach(([key, elId]) => {
        const saved = localStorage.getItem('fridge_' + key);
        if (saved) {
          if (key.endsWith('_size')) {
            const v = parseFloat(saved);
            if (isFinite(v) && v > 0) this[key] = v;
          } else {
            this[key] = saved;
          }
          const el = document.getElementById(elId);
          if (el) el.value = String(this[key]);
        }
      });
    } catch (e) { /* 忽略 */ }

    // 显示当前字体摘要
    this.updateFontInfo();

    // 生成预览
    document.getElementById('fridge-parse-btn').addEventListener('click', () => {
      self.refreshPreview();
    });

    // 发送到屏幕
    document.getElementById('fridge-send-btn').addEventListener('click', () => {
      self.sendToScreen();
    });

    // 初始预览（恢复缓存后渲染，不再空白）
    this.refreshPreview();

    // 预加载像素字体
    const fontLoadPromises = [];
    ['FusionPixel', 'FusionPixel10', 'FusionPixel12'].forEach((fam) => {
      [8, 10, 12, 16, 20, 24].forEach((px) => {
        fontLoadPromises.push(document.fonts.load(`${px}px "${fam}"`, '冷藏冰箱库存'));
      });
    });
    // 文泉驿点阵字体（16px档用）
    fontLoadPromises.push(document.fonts.load('16px "WQY16"', '冷藏冰箱库存'));
    fontLoadPromises.push(document.fonts.load('12px "WQY12"', '冷藏冰箱库存'));
    // 立方体11 + Unifont
    fontLoadPromises.push(document.fonts.load('16px "Cubic11"', '冷藏冰箱库存'));
    fontLoadPromises.push(document.fonts.load('12px "Cubic11"', '冷藏冰箱库存'));
    fontLoadPromises.push(document.fonts.load('16px "Unifont"', '冷藏冰箱库存'));
    fontLoadPromises.push(document.fonts.load('12px "Unifont"', '冷藏冰箱库存'));
    // 观致8px（系统提示小字）
    fontLoadPromises.push(document.fonts.load('8px "GuanZhi8"', '字体'));
    Promise.allSettled(fontLoadPromises).then(() => {
      self.refreshPreview();
      self.refreshFillTest();
    });

    // 字号满幅测试
    const testSizeBtns = document.querySelectorAll('.test-size-btn');
    testSizeBtns.forEach((btn) => {
      btn.addEventListener('click', () => {
        self.testSize = parseInt(btn.dataset.size, 10);
        testSizeBtns.forEach((b) => b.classList.toggle('active', b === btn));
        self.refreshFillTest();
      });
    });
    const testSendBtn = document.getElementById('fridge-test-send-btn');
    if (testSendBtn) testSendBtn.addEventListener('click', () => self.sendTestToScreen());
    const fontsizeSendBtn = document.getElementById('fridge-fontsize-send-btn');
    if (fontsizeSendBtn) fontsizeSendBtn.addEventListener('click', () => self.sendFontSizeTestToScreen());

    // 初始满幅测试预览
    this.refreshFillTest();
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
    // 按屏幕尺寸 + 方向设置画布尺寸
    // 4.2寸: 横400x300 竖300x400；7.5寸: 横800x480 竖480x800（2倍）
    const is75 = (this.screen === '7.5');
    if (this.orientation === 'portrait') {
      canvas.width = is75 ? 480 : 300;
      canvas.height = is75 ? 800 : 400;
    } else {
      canvas.width = is75 ? 800 : 400;
      canvas.height = is75 ? 480 : 300;
    }
    const zones = this.loadFromInput();
    if (!zones || zones.length === 0) {
      // 无数据：显示提示（适配当前方向尺寸）
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.fillStyle = '#999999';
      ctx.font = '12px "FusionPixel12"';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText('请粘贴冰箱数据', canvas.width / 2, canvas.height / 2);
      this.updateStats(null);
      return;
    }
    // 渲染前确保字体已加载（12px 字体文件大，手机端网络慢可能未就绪）
    // 用 try-catch 暴露手机端具体错误
    try {
      const result = this.render(canvas, ctx);
      this.updateStats(zones, result);
    } catch (e) {
      console.error('冰箱库存渲染失败:', e);
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, 400, 300);
      ctx.fillStyle = '#cc0000';
      ctx.font = '10px "FusionPixel10"';
      ctx.textAlign = 'left';
      ctx.textBaseline = 'top';
      ctx.fillText('渲染错误: ' + (e && e.message ? e.message : e), 8, 8);
      if (typeof addLog === 'function') addLog('冰箱库存渲染错误: ' + (e && e.message ? e.message : e));
    }
  }

  /**
   * 更新统计信息
   */
  updateStats(zones, result) {
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
    let extra = '';
    if (result && result.effRowH) {
      extra = ` 行高${result.effRowH}px${result.totalHeight ? ` 总高${result.totalHeight}` : ''}`;
    }
    el.textContent = `${zones.length}区 ${cats}分类 ${items}条目${extra}`;
  }

  /**
   * 根据基准字号返回对应字体族名
   */
  fontFamily(px) {
    // 8px -> FusionPixel, 10px -> FusionPixel10, 12px -> FusionPixel12
    if (px <= 8) return 'FusionPixel';
    if (px <= 10) return 'FusionPixel10';
    return 'FusionPixel12';
  }

  /**
   * 渲染"字号满幅测试"：用指定字号把整屏填满文字
   * 使用给定的文本，逐字排布成满屏，展示该字号的信息量
   */
  renderFillTest(canvas, ctx, sizePx) {
    const W = canvas.width;
    const H = canvas.height;
    const px = sizePx;
    const fontFamily = this.fontFamily(px);
    ctx.font = `${px}px "${fontFamily}"`;
    ctx.textBaseline = 'top';
    ctx.textAlign = 'left';
    ctx.imageSmoothingEnabled = false;

    // 白底
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);
    ctx.fillStyle = '#000000';

    // 行高 = 字号 * 1.3（像素字体行距）
    const lineH = Math.round(px * 1.3);
    // 每行能容纳的字数：400 宽 / 每字宽（中文字宽=字号）
    const charsPerLine = Math.floor(W / px);
    // 能容纳的行数
    const lines = Math.floor((H - 4) / lineH);

    // 测试文本：去掉标点与空白，逐字填充
    const rawText = this.testText;
    const cleanText = rawText.replace(/[\s　、。，；：""''《》〈〉（）【】〖〗—…·!?！？]/g, '');
    const chars = Array.from(cleanText);
    const totalNeeded = charsPerLine * lines;

    // 逐行填入字符（超出部分循环复用，保证填满）
    let y = 2;
    let ci = 0;
    for (let l = 0; l < lines; l++) {
      let line = '';
      for (let c = 0; c < charsPerLine; c++) {
        line += chars.length ? chars[ci % chars.length] : ' ';
        ci++;
      }
      ctx.fillText(line, 2, y);
      y += lineH;
    }

    // 返回统计信息，供 info 显示
    return {
      size: px,
      charsPerLine: charsPerLine,
      lines: lines,
      totalChars: charsPerLine * lines,
      sourceChars: chars.length,
    };
  }

  /**
   * 刷新字号测试预览
   */
  refreshFillTest() {
    const canvas = document.getElementById('fridge-test-preview');
    const ctx = canvas.getContext('2d');
    canvas.width = 400;
    canvas.height = 300;
    const info = this.renderFillTest(canvas, ctx, this.testSize);
    const infoEl = document.getElementById('fridge-test-info');
    if (infoEl) {
      infoEl.textContent =
        `${info.size}px：每行 ${info.charsPerLine} 字 × ${info.lines} 行，` +
        `满屏可容纳 ${info.totalChars} 字（源文本 ${info.sourceChars} 字${info.sourceChars >= info.totalChars ? '' : '，已循环复用'}）`;
    }
  }

  /**
   * 把测试图渲染到主画布（供发送）
   */
  renderTestToMain() {
    this.renderFillTest(canvas, ctx, this.testSize);
  }

  /**
   * 发送测试图到墨水屏
   */
  async sendTestToScreen() {
    if (!epdCharacteristic) {
      addLog("未连接设备，请先连接蓝牙");
      showToast("请先连接设备");
      return;
    }
    // 先确保主画布尺寸
    const canvasSize = document.getElementById('canvasSize').value;
    if (canvasSize !== '4.2_400_300') {
      const sizeSelect = document.getElementById('canvasSize');
      sizeSelect.value = '4.2_400_300';
      updateCanvasSize();
    }
    document.getElementById('ditherMode').value = 'threeColor';

    // 渲染测试图到主画布
    this.renderTestToMain();

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

  /**
   * 渲染"字号对比测试"：把多个字号并排显示，用于发送到设备看哪种最清晰
   * 关键：整数倍放大（8→16, 10→20, 12→24）应比非整数倍更锐利
   */
  renderFontSizeTest(canvas, ctx) {
    const W = canvas.width;
    const H = canvas.height;
    // 每档：{ 字体族, 渲染字号, 标注 }
    const tests = [
      { font: 'FusionPixel12', px: 12, label: '12px(12字体)' },
      { font: 'FusionPixel',   px: 16, label: '16px(8字体x2)' },
      { font: 'FusionPixel10', px: 20, label: '20px(10字体x2)' },
      { font: 'FusionPixel12', px: 24, label: '24px(12字体x2)' },
    ];
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);
    ctx.fillStyle = '#000000';
    let y = 14;
    for (const t of tests) {
      ctx.font = t.px + 'px "' + t.font + '"';
      ctx.textBaseline = 'alphabetic';
      ctx.fillText(t.label + ' 冷藏', 10, y + t.px * 0.8);
      y += t.px * 1.6;
    }
    ctx.font = '12px "FusionPixel12"';
    ctx.fillStyle = '#888888';
    ctx.fillText('（发送到设备看清晰度，选区名用字号）', 10, H - 10);
  }

  /**
   * 把字号对比测试图渲染到主画布（供发送）
   */
  renderFontSizeTestToMain() {
    this.renderFontSizeTest(canvas, ctx);
  }

  /**
   * 发送字号对比测试图到墨水屏
   */
  async sendFontSizeTestToScreen() {
    if (!epdCharacteristic) {
      addLog("未连接设备，请先连接蓝牙");
      showToast("请先连接设备");
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

    // 渲染字号对比测试图到主画布
    this.renderFontSizeTestToMain();

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

    // 发送前准备：主画布设为驱动尺寸，渲染并旋转到位
    document.getElementById('ditherMode').value = 'threeColor';
    // 根据屏幕尺寸确定驱动分辨率：4.2寸 400x300，7.5寸 800x480
    const is75 = (this.screen === '7.5');
    const DRV_W = is75 ? 800 : 400;
    const DRV_H = is75 ? 480 : 300;
    // 竖屏渲染尺寸（横竖互换）
    const PRT_W = is75 ? 480 : 300;
    const PRT_H = is75 ? 800 : 400;
    // 主画布设为驱动尺寸
    if (canvas.width !== DRV_W || canvas.height !== DRV_H) {
      canvas.width = DRV_W;
      canvas.height = DRV_H;
    }
    ctx.setTransform(1, 0, 0, 1, 0, 0); // 重置变换
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, DRV_W, DRV_H);

    if (this.orientation === 'portrait') {
      // 竖屏：先渲染 竖屏尺寸，旋转90°后画到主画布(驱动尺寸)
      const tmp = document.createElement('canvas');
      tmp.width = PRT_W;
      tmp.height = PRT_H;
      const tmpCtx = tmp.getContext('2d', { willReadFrequently: true });
      this.render(tmp, tmpCtx);
      // 旋转到驱动尺寸（translate 用 tmp 宽度=驱动高度）
      ctx.translate(0, DRV_H);
      ctx.rotate(-Math.PI / 2);
      ctx.drawImage(tmp, 0, 0);
      ctx.setTransform(1, 0, 0, 1, 0, 0); // 重置，确保 getImageData 读到最终像素
    } else {
      // 横屏：直接渲染主画布
      this.render(canvas, ctx);
    }

    // 发送用主画布（与 sendimg 完全一致）
    const sendCanvas = canvas;
    const sendCtx = ctx;

    startTime = new Date().getTime();
    const status = document.getElementById("status");
    status.parentElement.style.display = "block";

    const imageData = sendCtx.getImageData(0, 0, sendCanvas.width, sendCanvas.height);
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
