/**
 * 冰箱备忘 (FridgeNote) 模块
 *
 * 将用户填写的"项目名/备注"表格，用缝合像素字体渲染成一张位图，
 * 直接画到全局画布上，再复用现有 BLE 传输流程发送到墨水屏。
 *
 * 特性：
 *  1. 像素字体（8/10/12px 等宽），整数像素渲染，保证锐利
 *  2. 三色支持：黑色正文、白色背景、红色标题/分隔线（三色屏）
 *  3. 表格自动居中、六行均分，项目名左对齐、备注右对齐
 *  4. 超长内容自动截断/换行，避免溢出
 *  5. 网页内预览 + 一键发送
 */
class FridgeNote {
  constructor() {
    this.ROWS = 6;
    this.COLS = 2;
    this.ROW_LABELS = ['项目名', '备注'];
    this.testSize = 10; // 当前测试字号 (8/10/12)
    this.testRunning = false; // 是否正在测试填充模式
    // 满幅测试用的文本（去掉标点空白后逐字填屏）
    this.testText = '我在漫宿高处的房间居住了一小段时间，而后由三尖之门返回，我的躯体在灰烬中跃跃欲出。我已毛发全无，如大理石般千古不朽，铸炉之火仍在我内里燃烧。我具有〖足以塑形的气力〗。我再不会变老。也许我将反叛。也许有朝一日，我将升得更高。我在漫宿高处的房间居住了一小段时间，而后由三尖之门返回，我撕去了我旧肉体那湿黏的条条褴褛。我的新躯体外表光滑，内里像甜美的果实般鲜红。我的四肢牢固如缆绳。我的五感好似刀子。我不会变老。狂暴之门是一道旧伤口，不会轻易开启，但它还是再度开启了。连石源之神都使用过此门，那么在祂们前有多少神用过呢？辉光是一个疑问，我们所有人，作为燃料供给我们的沐光明者的见证者与我，都以肯定作答。地撒银光，晚风凉爽；夜晚的寂静引来言语，如暗色丝绒引来手指抚弄。蛾在黑暗中寻见光，我亦如是。一种似有似无的惆怅染上了我。每当灯之准则放光时，我便可以多记起一些残篇知识。脑海中的灰尘一扫而空，至少暂时如此。我在半夜惊醒。我的手掌，我的脸上都渗出汗珠。慢慢地，我松开紧攥的拳头。我又滑落回睡梦中。但某种东西留下了，好像房间角落的蛛网。沙皇时代我曾在泰加林狩猎，那里的深处有一个虎口形状的湖。也许我的大敌不会去那里找我，如果他去了，那就是时候奋起反抗了。我们交谈时，万查妮一寸一寸靠近。她握住我手时，她的手指传来机械般的力道，她的皮肤有着余烬般的热力。“我不是光，”她警告我说，“我们应当对彼此宽容。”埃尔里奇已基本能管住自己在谈话中不用刀剔牙了。但他还不能把刀一直放在靴子里不拔出来。埃尔里奇正在耐心等待再次捅人刀子的时机。芮妮拉在三个不相邻的辖区里被判三次无罪释放。她很可能永远不会因为任何事获罪。悲伤时，她便闭口不言。但她时常悲伤。我如今知道了该做出什么姿势，手该恰到好处地放在何处，慢慢领她那可爱的颈脖靠上我肩。我们的生活是幸福的。我们看上去是一对般配的恋人。当然了，她的哥哥时而来访——但那是因为我出门旅行时她寂寞，需要陪伴。我的脑子装满了的知识，快要撑裂颅骨。“只欲求而不行动者于世有害。”“与其哺育只欲求而不行动的婴孩，不如将其尽早扼杀在摇篮中。”“我们指引前路，我们照明驱暗，我们无有怜悯之心。”秘密温柔，夜晚更加温柔。大海会讲述，但是聆听不是何时都是明智的。灰暗的海水每天早晨都冲击着岩礁，又整夜整夜地对我发出低吼。每一天的日出都来得比前一日更早。风也变得更加温柔。春天快要来了。';
  }

  /**
   * 从表格 DOM 读取内容
   * @returns {{title: string, fontSize: string, colorMode: string, rows: string[][]}}
   */
  getData() {
    const title = document.getElementById('fridge-title').value.trim();
    const fontSize = document.getElementById('fridge-font-size').value;
    const colorMode = document.getElementById('fridge-color-mode').value;

    const rows = [];
    const tbody = document.querySelector('#fridge-table-body');
    const trs = tbody.querySelectorAll('tr');
    trs.forEach((tr) => {
      const inputs = tr.querySelectorAll('input');
      const row = [inputs[0].value.trim(), inputs[1].value.trim()];
      rows.push(row);
    });
    // 确保六行
    while (rows.length < this.ROWS) rows.push(['', '']);
    return { title, fontSize, colorMode, rows };
  }

  /**
   * 渲染备忘卡片到目标 canvas 的 2D context
   * 使用像素字体，整数坐标，关闭抗锯齿，保证 8/10/12px 锐利
   */
  render(canvas, ctx, data) {
    const W = canvas.width;
    const H = canvas.height;
    const fontSize = parseInt(data.fontSize, 10) || 12;
    // 像素字体必须是整数倍显示才锐利；这里直接取用户选的基准字号
    const px = fontSize;
    const fontFamily = this.fontFamily(px);
    // 字体名带引号，避免空格等特殊字符
    ctx.font = `${px}px "${fontFamily}"`;
    ctx.textBaseline = 'top';
    ctx.textAlign = 'left';

    // 关掉图像平滑，保证像素字锐利（对文本也生效）
    ctx.imageSmoothingEnabled = false;

    // 白底
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);

    // 颜色工具
    const black = '#000000';
    const red = '#ff0000';
    const useRed = data.colorMode !== 'blackWhite';

    // 中文像素字体的行高因子（fusion-pixel 实际约占字号的 ~1.3 倍）
    const lineH = Math.round(px * 1.3);
    // 垂直方向让文字在行内稍微下移一点，视觉居中（像素字体偏上）
    const textDy = Math.round(px * 0.1);

    // ---- 标题 ----
    const title = data.title || '冰箱备忘';
    let titleY = 0;
    if (title) {
      ctx.fillStyle = useRed ? red : black;
      ctx.textAlign = 'center';
      // 标题字号 = 基准字号 + 4px（保持像素整数）
      const titlePx = px + 4;
      ctx.font = `${titlePx}px "${fontFamily}"`;
      const titleWidth = ctx.measureText(title).width;
      // 标题居上居中
      titleY = Math.round(px * 0.5);
      ctx.fillText(title, Math.round(W / 2), titleY);
      // 标题下加一条红色分隔线
      ctx.fillStyle = useRed ? red : black;
      ctx.fillRect(
        Math.round((W - Math.min(titleWidth, W - 20)) / 2),
        titleY + titlePx + 6,
        Math.round(Math.min(titleWidth, W - 20)),
        2
      );
      titleY = titleY + titlePx + 6 + 2;
    }

    // ---- 表格区域 ----
    const topPad = 4;
    const bottomPad = 6;
    const tableTop = titleY + topPad;
    const tableBottom = H - bottomPad;
    // 表头高度固定为一行
    const headerHeight = Math.round(px * 1.4);
    // 6 行数据均分剩余空间
    const dataAreaTop = tableTop + headerHeight;
    const dataAreaHeight = tableBottom - dataAreaTop;
    const rowHeight = Math.floor(dataAreaHeight / this.ROWS);
    const rowGap = dataAreaHeight - rowHeight * this.ROWS; // 余数，均摊到各行
    // 表格左右留边
    const margin = Math.max(8, Math.round(W * 0.04));
    const tableLeft = margin;
    const tableRight = W - margin;
    const tableWidth = tableRight - tableLeft;
    // 第一列宽度：按最长项目名自适应（最多 8 字），其余给备注
    ctx.font = `${px}px "${fontFamily}"`;
    let maxItemChars = 4;
    data.rows.forEach((row) => {
      if (row[0]) maxItemChars = Math.max(maxItemChars, row[0].length);
    });
    // 限制项目名最多显示 8 字，备注至少留 10 字宽
    maxItemChars = Math.min(8, maxItemChars);
    const col1Width = Math.max(px * 3, Math.round(px * maxItemChars) + px);
    const col2Left = tableLeft + col1Width;

    // 列头
    ctx.fillStyle = '#000000';
    ctx.textAlign = 'left';
    let y = tableTop;
    // 表头文字垂直居中
    ctx.fillText(this.ROW_LABELS[0], tableLeft + 4, tableTop + Math.round(headerHeight / 2) - Math.round(px * 0.4));
    ctx.fillText(this.ROW_LABELS[1], col2Left + 4, tableTop + Math.round(headerHeight / 2) - Math.round(px * 0.4));
    // 表头下划线（红色）
    ctx.fillStyle = useRed ? red : black;
    ctx.fillRect(tableLeft, tableTop + headerHeight - 2, tableWidth, 2);
    // 数据行从表头下方开始
    y = dataAreaTop;

    // 数据行
    ctx.fillStyle = '#000000';
    for (let i = 0; i < this.ROWS; i++) {
      const rowTop = y;
      // 余数均摊到前几行，最后一行正好落在 tableBottom
      const rowBottom = y + rowHeight + (i < rowGap ? 1 : 0);
      const item = (data.rows[i] && data.rows[i][0]) || '';
      const note = (data.rows[i] && data.rows[i][1]) || '';

      // 行分隔线（浅灰，用 0.3 透明度模拟浅色；墨水屏上显示为浅灰）
      ctx.fillStyle = 'rgba(0,0,0,0.3)';
      ctx.fillRect(tableLeft, rowBottom - 1, tableWidth, 1);
      ctx.fillStyle = '#000000';

      // 项目名：左对齐，单行，超长截断
      ctx.font = `${px}px "${fontFamily}"`;
      ctx.textAlign = 'left';
      let itemText = item;
      const maxItemWidth = col1Width - 8;
      while (itemText && ctx.measureText(itemText).width > maxItemWidth) {
        itemText = itemText.slice(0, -1);
      }
      ctx.fillText(itemText, tableLeft + 4, rowTop + textDy);

      // 备注：右对齐，单行，超长截断
      ctx.textAlign = 'right';
      let noteText = note;
      const maxNoteWidth = (tableRight - col2Left) - 8;
      while (noteText && ctx.measureText(noteText).width > maxNoteWidth) {
        noteText = noteText.slice(0, -1);
      }
      ctx.fillText(noteText, tableRight - 4, rowTop + textDy);

      y = rowBottom;
    }
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
   * 根据基准字号返回对应字体族名
   */
  fontFamily(px) {
    // 8px -> FusionPixel, 10px -> FusionPixel10, 12px -> FusionPixel12
    if (px <= 8) return 'FusionPixel';
    if (px <= 10) return 'FusionPixel10';
    return 'FusionPixel12';
  }

  /**
   * 刷新网页预览
   */
  refreshPreview() {
    const canvas = document.getElementById('fridge-preview');
    const ctx = canvas.getContext('2d');
    canvas.width = 400;
    canvas.height = 300;
    const data = this.getData();
    this.render(canvas, ctx, data);
  }

  /**
   * 渲染到主画布（供发送）
   */
  renderToMain() {
    const data = this.getData();
    this.render(canvas, ctx, data);
  }

  /**
   * 发送到墨水屏：确保主画布为 400x300，渲染，复用现有传输流程
   */
  async sendToScreen() {
    if (!epdCharacteristic) {
      addLog("未连接设备，请先连接蓝牙");
      showToast("请先连接设备");
      return;
    }

    // 先确保主画布尺寸为 4.2 寸 400x300（切换尺寸会触发 updateCanvasSize 清空画布，
    // 所以必须在渲染之前做）
    const canvasSize = document.getElementById('canvasSize').value;
    if (canvasSize !== '4.2_400_300') {
      const sizeSelect = document.getElementById('canvasSize');
      sizeSelect.value = '4.2_400_300';
      updateCanvasSize();
    }
    const epdDriverSelect = document.getElementById('epddriver');
    // 三色屏需要 ditherMode 为 threeColor，否则 processImageData 输出格式不对
    document.getElementById('ditherMode').value = 'threeColor';

    // 渲染备忘卡片到主画布（此时画布已是 400x300）
    this.renderToMain();

    // 复用现有流程：读取画布 -> 三色转换 -> CRC 传输 -> 刷新
    startTime = new Date().getTime();
    const status = document.getElementById("status");
    status.parentElement.style.display = "block";

    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const processedData = processImageData(imageData, 'threeColor');

    updateButtonStatus(true);

    // 使用 CRC 传输
    let useCRC = (appVersion >= 0x20) && (typeof BleTransfer !== 'undefined');
    const transferFn = useCRC ? writeImageCRC : writeImage;

    if (useCRC) {
      addLog("使用CRC校验传输模式");
    }

    const halfLength = Math.floor(processedData.length / 2);
    const blackWhiteData = processedData.slice(0, halfLength);
    const redWhiteData = processedData.slice(halfLength);

    // 4.2 寸三色屏：黑白平面 + 红白平面
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
    const hideStatusBar = () => {
      status.parentElement.style.display = "none";
    };
    hideStatusBar();
  }

  /**
   * 初始化事件绑定
   */
  init() {
    const self = this;

    // 表格输入实时预览
    const tbody = document.querySelector('#fridge-table-body');
    tbody.addEventListener('input', () => self.refreshPreview());

    document.getElementById('fridge-title').addEventListener('input', () => self.refreshPreview());
    document.getElementById('fridge-font-size').addEventListener('change', () => self.refreshPreview());
    document.getElementById('fridge-color-mode').addEventListener('change', () => self.refreshPreview());

    document.getElementById('fridge-send-btn').addEventListener('click', () => self.sendToScreen());
    document.getElementById('fridge-preview-btn').addEventListener('click', () => {
      self.renderToMain();
      showToast('已渲染到主画布，可预览或发送');
      addLog('备忘卡片已渲染到主画布');
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
    document.getElementById('fridge-test-send-btn').addEventListener('click', () => self.sendTestToScreen());

    // 初始预览
    self.refreshPreview();
    self.refreshFillTest();

    // 预加载像素字体：确保字体就绪后再渲染，避免首屏用默认字体导致"残缺"
    // （12px 字体文件 6.6MB，网络慢时可能加载慢；font-display:block 已避免残缺闪屏，
    //   这里再主动等字体就绪后重绘一次，彻底消除首屏问题）
    const fontLoadPromises = [];
    ['FusionPixel', 'FusionPixel10', 'FusionPixel12'].forEach((fam) => {
      [8, 10, 12, 14, 16].forEach((px) => {
        fontLoadPromises.push(document.fonts.load(`${px}px "${fam}"`, '冰箱备忘'));
      });
    });
    Promise.allSettled(fontLoadPromises).then(() => {
      // 字体全部就绪（或失败），重绘预览
      self.refreshPreview();
      self.refreshFillTest();
      addLog('像素字体加载完成');
    });
  }
}

// 全局实例
let fridgeNote = null;
