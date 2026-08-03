// 官方 document-skills pptx skill：用 pptxgenjs 从零建 deck
// 设计指南：Midnight Executive 色板（1E2761 藏青 / CADCFC 冰蓝 / 白）
// 视觉母题：圆角"芯片"（AI 芯片主题），三明治结构（暗封面+结尾/亮内容）
// 安全字体：Cambria(标题) + Calibri(正文)；颜色一律不带 # 前缀
const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9'; // 10in x 5.625in
pptx.author = 'MelonAI';
pptx.title = '今日三大AI要闻：打工人视角';

// ---- 调色板 ----
const NAVY = '1E2761';      // 主色（暗）
const MIDNIGHT = '141A3C';  // 深墨蓝（次暗）
const ICE = 'CADCFC';       // 冰蓝（亮色点缀）
const WHITE = 'FFFFFF';
const INK = '20265A';       // 正文深蓝
const MUTE = '5A6090';      // 次要文字
const ICE_SOFT = 'EEF3FF';  // 冰蓝浅底（卡片）
const ACCENT = 'F5B041';    // 琥珀色（仅行动卡，克制使用）

const SERIF = 'Cambria';
const SANS = 'Calibri';

const M = 0.5; // margin

// ---- 工具 ----
function header(slide, kicker, title) {
    slide.addText(kicker, { x: M, y: 0.42, w: 8, h: 0.3, fontSize: 12, bold: true, color: NAVY, fontFace: SANS, charSpacing: 2 });
    slide.addText(title, { x: M, y: 0.72, w: 9, h: 0.6, fontSize: 30, bold: true, color: MIDNIGHT, fontFace: SERIF });
    // 母题：标题左下角小芯片
    slide.addShape(pptx.ShapeType.roundRect, { x: M, y: 1.42, w: 0.5, h: 0.12, fill: { color: ICE }, line: { color: ICE }, rectRadius: 0.5 });
}

function chip(slide, x, y, size, text, fill, color) {
    // 视觉母题：圆角芯片（数字/对勾容器）
    slide.addShape(pptx.ShapeType.roundRect, {
        x, y, w: size, h: size,
        fill: { color: fill }, line: { color: fill }, rectRadius: 0.3
    });
    slide.addText(text, {
        x, y, w: size, h: size, align: 'center', valign: 'middle',
        fontSize: size * 0.42, bold: true, color, fontFace: SANS
    });
}

// ============================================================
// Slide 1 · 封面（暗）
// ============================================================
let s = pptx.addSlide();
s.background = { color: NAVY };
// 右上角一组芯片（母题开场）
chip(s, 8.9, 0.55, 0.6, 'AI', ICE, NAVY);
chip(s, 9.35, 0.55, 0.6, '3', WHITE, NAVY);
s.addText('MELONAI · 职场AI日报 · 2026-08-02', { x: M, y: 1.0, w: 8, h: 0.3, fontSize: 13, bold: true, color: ICE, fontFace: SANS, charSpacing: 2 });
s.addText('今日三大AI要闻', { x: M, y: 1.5, w: 9, h: 1.2, fontSize: 48, bold: true, color: WHITE, fontFace: SERIF });
s.addText('打工人视角解读 —— 5分钟看懂今天该知道的 AI 大事', { x: M, y: 2.6, w: 9, h: 0.5, fontSize: 17, color: ICE, fontFace: SANS });
// 底部三张预览芯片卡
const topics = [
    { n: '01', t: '国产免费反超', d: 'DeepSeek-V4-Flash 正式版，干活能力接近顶级' },
    { n: '02', t: 'AI 干长活立功', d: 'Astra 花 $2000 连破 10 道数学世纪难题' },
    { n: '03', t: '开源登顶·资本大战', d: 'Kimi K3 全球最大开源模型，估值 350 亿' },
];
const cardW = 2.8, cardH = 1.55, gap = 0.2;
topics.forEach((tp, i) => {
    const x = M + i * (cardW + gap);
    const y = 3.35;
    s.addShape(pptx.ShapeType.roundRect, { x, y, w: cardW, h: cardH, fill: { color: MIDNIGHT }, line: { color: MIDNIGHT }, rectRadius: 0.08 });
    s.addText(tp.n, { x: x + 0.22, y: y + 0.16, w: 1, h: 0.5, fontSize: 20, bold: true, color: ICE, fontFace: SERIF });
    s.addText(tp.t, { x: x + 0.22, y: y + 0.62, w: cardW - 0.4, h: 0.35, fontSize: 13, bold: true, color: WHITE, fontFace: SANS });
    s.addText(tp.d, { x: x + 0.22, y: y + 0.98, w: cardW - 0.4, h: 0.5, fontSize: 9, color: ICE, fontFace: SANS });
});
s.addText('来源：官方博客 / 官方 GitHub / 官方产品站 ｜ 融资信息均标注「据媒体报道」', { x: M, y: 5.15, w: 9, h: 0.3, fontSize: 8, color: ICE, fontFace: SANS });

// ============================================================
// Slide 2 · 目录（亮）
// ============================================================
s = pptx.addSlide();
s.background = { color: WHITE };
header(s, 'AGENDA · 今日目录', '今天要聊的三件大事');
const agenda = [
    { n: '01', t: 'DeepSeek-V4-Flash 正式版发布', d: '国产免费模型反超：顶级干活能力，几乎免费', sub: '总参 284B / 激活 13B · 1M 上下文 · 单Token算力前代27%' },
    { n: '02', t: 'OpenAI Astra 数学突破', d: 'AI 花 $2000 连破 10 道数学世纪难题', sub: '249 页手稿 · Lean 4 形式化验证 · next major model' },
    { n: '03', t: 'Kimi K3 开源 + 融资', d: '全球最大开源模型免费给你用，背后公司要上市', sub: '2.8T 总参 · F轮35亿美元 · 估值350亿 · 港股传闻' },
];
agenda.forEach((a, i) => {
    const y = 1.75 + i * 1.25;
    s.addShape(pptx.ShapeType.roundRect, { x: M, y, w: 9, h: 1.08, fill: { color: ICE_SOFT }, line: { color: ICE_SOFT }, rectRadius: 0.12 });
    chip(s, M + 0.18, y + 0.22, 0.64, a.n, NAVY, WHITE);
    s.addText(a.t, { x: M + 1.05, y: y + 0.13, w: 5.4, h: 0.35, fontSize: 15, bold: true, color: INK, fontFace: SANS });
    s.addText(a.d, { x: M + 1.05, y: y + 0.48, w: 7.6, h: 0.3, fontSize: 11, color: INK, fontFace: SANS });
    s.addText(a.sub, { x: M + 1.05, y: y + 0.78, w: 7.6, h: 0.25, fontSize: 8.5, color: MUTE, fontFace: SANS });
});

// ============================================================
// facts + KPI 通用函数（亮）
// ============================================================
function factsSlide(kicker, title, facts, metrics) {
    const sl = pptx.addSlide();
    sl.background = { color: WHITE };
    header(sl, kicker, title);
    // 左侧事实列表
    facts.forEach((f, i) => {
        const y = 1.75 + i * 0.5;
        sl.addText([
            { text: '✓  ', options: { color: NAVY, bold: true } },
            { text: f, options: { color: INK } }
        ], { x: M, y, w: 5.2, h: 0.46, fontSize: 11.5, fontFace: SANS, valign: 'top', lineSpacingMultiple: 1.05 });
    });
    // 右侧 KPI 面板（暗色卡片组）
    const px = 6.0, py = 1.7, pw = 3.5, ph = 3.4;
    sl.addShape(pptx.ShapeType.roundRect, { x: px, y: py, w: pw, h: ph, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.06 });
    sl.addText('核心数据', { x: px + 0.25, y: py + 0.2, w: 3, h: 0.35, fontSize: 14, bold: true, color: WHITE, fontFace: SANS });
    metrics.forEach((mt, i) => {
        const col = i % 2, row = Math.floor(i / 2);
        const mx = px + 0.22 + col * 1.62;
        const my = py + 0.65 + row * 0.9;
        sl.addShape(pptx.ShapeType.roundRect, { x: mx, y: my, w: 1.5, h: 0.78, fill: { color: MIDNIGHT }, line: { color: MIDNIGHT }, rectRadius: 0.1 });
        sl.addText(mt[0], { x: mx, y: my + 0.06, w: 1.5, h: 0.4, align: 'center', fontSize: 16, bold: true, color: ICE, fontFace: SERIF });
        sl.addText(mt[1], { x: mx + 0.05, y: my + 0.5, w: 1.4, h: 0.24, align: 'center', fontSize: 7.5, color: WHITE, fontFace: SANS });
    });
    return sl;
}

// Slide 3 · 话题一 事实
factsSlide(
    'TOPIC 01 · DEEPSEEK-V4-FLASH', '国产免费模型，追到顶级 AI 脸上了',
    [
        '2026-07-31 正式版 API 上线公测，现有集成不改代码自动升级',
        '轻量化 MoE：总参数 284B / 激活 13B，1M 超长上下文',
        '官方 9 项 Agent 基准多项超过三个月前的 V4-Pro 预览版',
        '原生兼容 OpenAI Responses API，针对性适配 Codex',
        '支持思考 / 非思考模式切换',
        '百万 Token 场景单 Token 算力仅为前代 27%，价格极低',
    ],
    [
        ['284B', '总参数'], ['13B', '激活参数'], ['1M', '上下文'],
        ['27%', '单Token算力成本'], ['$0.0028', '缓存命中输入价'], ['免费', 'APP/WEB 可用'],
    ]
);

// Slide 4 · 话题一 视角+行动
function perspectiveSlide(kicker, title, viewText, actionText) {
    const sl = pptx.addSlide();
    sl.background = { color: WHITE };
    header(sl, kicker, title);
    // 视角卡（冰蓝浅底）
    sl.addShape(pptx.ShapeType.roundRect, { x: M, y: 1.75, w: 9, h: 1.9, fill: { color: ICE_SOFT }, line: { color: ICE_SOFT }, rectRadius: 0.08 });
    sl.addText('一句话翻译', { x: M + 0.3, y: 1.95, w: 3, h: 0.35, fontSize: 13, bold: true, color: NAVY, fontFace: SANS });
    sl.addText(viewText, { x: M + 0.3, y: 2.35, w: 8.4, h: 1.2, fontSize: 13, color: INK, fontFace: SANS, lineSpacingMultiple: 1.3, valign: 'top' });
    // 行动卡（藏青）
    sl.addShape(pptx.ShapeType.roundRect, { x: M, y: 3.85, w: 9, h: 1.35, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08 });
    sl.addText([
        { text: '今天就能带走 · 行动  ', options: { bold: true, color: ICE } },
        { text: actionText, options: { color: WHITE } }
    ], { x: M + 0.3, y: 4.02, w: 8.4, h: 1.05, fontSize: 12.5, fontFace: SANS, valign: 'top', lineSpacingMultiple: 1.25 });
    return sl;
}
perspectiveSlide(
    'TOPIC 01 · 打工人视角', '跟你有什么关系？',
    '顶级干活能力 + 极低成本 → 普通打工人也能用上顶尖模型。选工具不用迷信国外付费产品；拆掉「必须花钱买课 / 买会员才能用好 AI」的焦虑。',
    '打开 chat.deepseek.com 或 DeepSeek 开放平台，用 V4-Flash 干一件真实工作上的活（写周报 / 整理数据 / 写代码），和现在用的工具对比一次。'
);

// Slide 5 · 话题二 事实
factsSlide(
    'TOPIC 02 · OPENAI ASTRA', 'AI 花 $2000，连破 10 道数学世纪难题',
    [
        'OpenAI 官方博客首次公开下一代模型 Astra（next major model）内部版本',
        '在 10 个长期未决的数学 / 理论计算机科学难题上做出新结果',
        '领域：高维几何、编码理论、群论、量子复杂性、格密码学等',
        '每个论证用 Lean 4 形式化验证，证书公开在 GitHub',
        '发布 249 页手稿合集 + 每个问题的 AI 推理过程旁白',
        '官方称全部 token 成本约 2000 美元',
    ],
    [
        ['10', '解出的世纪难题'], ['249页', '公开手稿'], ['$2000', 'Token 成本'],
        ['Lean 4', '形式化验证'], ['10+', '覆盖数学领域'], ['首个', 'Astra 内部版公开'],
    ]
);

// Slide 6 · 话题二 视角+行动
perspectiveSlide(
    'TOPIC 02 · 打工人视角', 'AI 从「答一句」升级到「干一件长活」',
    '数学难题本身离打工人远，但内核是：AI 从「答一句」升级到「独立干一件长活」。以后你可以把整件工作交给 AI，而不是一句一句喂。接续上一期「AI 自己干活没人盯会闯祸」——这次是「AI 干长活立了功」。',
    '把一个能拆成「交代清楚就能跑」的小任务（如让它自己调研 + 整理成报告），整件委托给 AI Agent 试一次，观察它能自主跑多久、做到什么程度。'
);

// Slide 7 · 话题三 事实
factsSlide(
    'TOPIC 03 · KIMI K3 开源 + 融资', '全球最大开源 AI 模型，免费给你用',
    [
        '月之暗面发布 Kimi K3：全球首个开源的三万亿级别大模型',
        '7/27 全链条开源：权重 + 47 页技术报告 + 三套基础设施',
        'WebDev Arena 以 1678 Elo 登顶，首个登顶的开源模型',
        '7/29 完成超 35 亿美元 F 轮，投后估值 350 亿美元',
        'F 轮认购超 3 倍提前关闭，市场传闻最快 6 个月港股上市',
        '1M token 上下文，原生视觉 MoonViT-V2',
    ],
    [
        ['2.8T', '总参数(首个3万亿级)'], ['104B', '激活参数'], ['1M', '上下文'],
        ['$35亿', 'F轮融资'], ['$350亿', '投后估值'], ['1678', 'WebDev Arena Elo'],
    ]
);

// Slide 8 · 话题三 视角+行动
perspectiveSlide(
    'TOPIC 03 · 打工人视角', '看懂资本，才不会被「焦虑营销」收割',
    '① 顶级能力免费开源 → AI 只会越来越便宜，49 块的 AI 课真没必要买；② 你天天用的 Kimi，背后是几百亿美元资本大战——免费背后有人替你付费，看懂资本才不会被「焦虑营销」收割。',
    '打开 Kimi（kimi.com）用 K3 试一个你正犹豫要不要付费的功能，对比一下免费开源到底能不能打。'
);

// ============================================================
// Slide 9 · 总结清单（亮）
// ============================================================
s = pptx.addSlide();
s.background = { color: WHITE };
header(s, 'SUMMARY · 三句话带走', '今天下班就能做');
const takeaways = [
    { t: '工具认知更新', d: '国产免费模型已具备顶级干活能力，选 AI 工具不用迷信国外付费' },
    { t: '使用方式升级', d: 'AI 能独立干一整件长活了：从「一句一句喂」变成「整件委托」' },
    { t: '消费决策清醒', d: 'AI 只会越来越便宜，别为焦虑买单，看懂免费背后的资本逻辑' },
];
takeaways.forEach((tk, i) => {
    const y = 1.75 + i * 0.95;
    s.addShape(pptx.ShapeType.roundRect, { x: M, y, w: 9, h: 0.82, fill: { color: ICE_SOFT }, line: { color: ICE_SOFT }, rectRadius: 0.12 });
    chip(s, M + 0.18, y + 0.16, 0.5, '✓', NAVY, WHITE);
    s.addText(tk.t, { x: M + 0.9, y: y + 0.12, w: 2.3, h: 0.35, fontSize: 13, bold: true, color: NAVY, fontFace: SANS });
    s.addText(tk.d, { x: M + 3.3, y: y + 0.12, w: 5.6, h: 0.6, fontSize: 11.5, color: INK, fontFace: SANS, valign: 'top', lineSpacingMultiple: 1.15 });
});
// 行动清单
s.addText('三件小事，今天下班就能做', { x: M, y: 4.62, w: 6, h: 0.3, fontSize: 13, bold: true, color: NAVY, fontFace: SANS });
const actions = [
    '用 DeepSeek V4-Flash 干一件真实工作上的活，和现用工具对比',
    '把一个「交代清楚就能跑」的小任务整件委托给 AI Agent',
    '打开 Kimi 用 K3 试一个犹豫要不要付费的功能',
];
actions.forEach((a, i) => {
    s.addText([
        { text: '▢  ', options: { color: ACCENT, bold: true } },
        { text: a, options: { color: INK } }
    ], { x: M, y: 4.95 + i * 0.34, w: 8.8, h: 0.32, fontSize: 11, fontFace: SANS });
});

// ============================================================
// Slide 10 · 结尾（暗）
// ============================================================
s = pptx.addSlide();
s.background = { color: NAVY };
chip(s, 4.45, 1.0, 1.1, '职场AI', ICE, NAVY);
s.addText('每天 5 分钟，读懂 AI 大事', { x: M, y: 2.5, w: 9, h: 0.7, align: 'center', fontSize: 30, bold: true, color: WHITE, fontFace: SERIF });
s.addText('下期见', { x: M, y: 3.35, w: 9, h: 0.5, align: 'center', fontSize: 16, color: ICE, fontFace: SANS });
s.addText('数据来源：官方博客 / 官方 GitHub / 官方产品站 ｜ 融资类信息均标注「据媒体报道」', { x: M, y: 5.15, w: 9, h: 0.3, align: 'center', fontSize: 8, color: ICE, fontFace: SANS });

pptx.writeFile({ fileName: '今日三大AI要闻_official.pptx' }).then(() => {
    console.log('saved: 今日三大AI要闻_official.pptx');
}).catch((e) => { console.error(e); process.exit(1); });
