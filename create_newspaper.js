const pptxgen = require('pptxgenjs');

async function createNewspaper() {
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.title = '端午节活动报纸机';
    pptx.author = '华为上研分部';

    // 颜色配置 - 年轻活泼的配色
    const colors = {
        red: 'E63946',
        yellow: 'FFD166',
        green: '06D6A0',
        blue: '118AB2',
        dark: '073B4C',
        white: 'FFFFFF',
        lightGray: 'F8F9FA',
        orange: 'FF9F1C'
    };

    // ========== 第1页：封面大标题 ==========
    let slide1 = pptx.addSlide();
    slide1.background = { color: colors.red };

    // 顶部装饰条
    slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 0.3,
        fill: { color: colors.yellow }
    });

    // 主标题框
    slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 1, y: 1.2, w: 8, h: 2,
        fill: { color: colors.white },
        line: { color: colors.dark, width: 3 }
    });

    slide1.addText('端午安康', {
        x: 1, y: 1.3, w: 8, h: 1.2,
        fontSize: 60, fontFace: 'Arial', bold: true,
        color: colors.red, align: 'center'
    });

    slide1.addText('上海练秋湖 · 粽情端午', {
        x: 1, y: 2.4, w: 8, h: 0.7,
        fontSize: 28, fontFace: 'Arial',
        color: colors.dark, align: 'center'
    });

    // 副标题
    slide1.addText('服务与软件研发管理部上研分部', {
        x: 1, y: 3.5, w: 8, h: 0.5,
        fontSize: 18, fontFace: 'Arial',
        color: colors.white, align: 'center'
    });

    // 镂空提示 - 人形轮廓
    slide1.addShape(pptx.shapes.OVAL, {
        x: 4, y: 4.2, w: 2, h: 2.5,
        fill: { color: colors.red },
        line: { color: colors.yellow, width: 4, dashType: 'dash' }
    });

    slide1.addText('📸\n在此合影', {
        x: 3.5, y: 4.8, w: 3, h: 1.2,
        fontSize: 14, fontFace: 'Arial',
        color: colors.yellow, align: 'center', valign: 'middle'
    });

    // ========== 第2页：震惊体 ==========
    let slide2 = pptx.addSlide();
    slide2.background = { color: colors.white };

    // 红色头条条
    slide2.addShape(pptx.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 1.2,
        fill: { color: colors.red }
    });

    slide2.addText('🔥 头 版 头条 🔥', {
        x: 0, y: 0.1, w: 10, h: 0.5,
        fontSize: 16, fontFace: 'Arial', bold: true,
        color: colors.yellow, align: 'center'
    });

    slide2.addText('震 惊 ！', {
        x: 0, y: 0.5, w: 10, h: 0.7,
        fontSize: 44, fontFace: 'Arial', bold: true,
        color: colors.white, align: 'center', charSpacing: 12
    });

    // 内容区
    slide2.addShape(pptx.shapes.RECTANGLE, {
        x: 0.5, y: 1.5, w: 9, h: 2.2,
        fill: { color: colors.lightGray },
        line: { color: colors.dark, width: 2 }
    });

    slide2.addText('世界首富亮相上海练秋湖！', {
        x: 0.5, y: 1.6, w: 9, h: 0.8,
        fontSize: 36, fontFace: 'Arial', bold: true,
        color: colors.red, align: 'center'
    });

    slide2.addText('竟是来参加华为端午粽香宴？\n神秘嘉宾身份引发热议，现场图流出...', {
        x: 0.8, y: 2.4, w: 8.4, h: 1.1,
        fontSize: 22, fontFace: 'Arial',
        color: colors.dark, align: 'center'
    });

    // 镂空人形
    slide2.addShape(pptx.shapes.OVAL, {
        x: 3.5, y: 4, w: 1.5, h: 2,
        fill: { color: colors.white },
        line: { color: colors.red, width: 4, dashType: 'dash' }
    });

    slide2.addText('📸', {
        x: 3.5, y: 4.5, w: 1.5, h: 0.8,
        fontSize: 24, align: 'center'
    });

    // 小标签
    slide2.addShape(pptx.shapes.RECTANGLE, {
        x: 7.5, y: 4.2, w: 2, h: 0.5,
        fill: { color: colors.yellow }
    });
    slide2.addText('限时拍照', {
        x: 7.5, y: 4.2, w: 2, h: 0.5,
        fontSize: 12, fontFace: 'Arial', bold: true,
        color: colors.dark, align: 'center', valign: 'middle'
    });

    // ========== 第3页：大明星空降 ==========
    let slide3 = pptx.addSlide();
    slide3.background = { color: colors.dark };

    // 紫色渐变条
    slide3.addShape(pptx.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 1.5,
        fill: { color: '9B5DE5' }
    });

    slide3.addText('⭐ 娱 乐 版 ⭐', {
        x: 0, y: 0.2, w: 10, h: 0.5,
        fontSize: 14, fontFace: 'Arial',
        color: colors.white, align: 'center'
    });

    slide3.addText('大明星空降上海练秋湖！', {
        x: 0, y: 0.6, w: 10, h: 0.8,
        fontSize: 36, fontFace: 'Arial', bold: true,
        color: colors.white, align: 'center'
    });

    // 内容卡片
    slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: 0.8, y: 1.8, w: 5.5, h: 3,
        fill: { color: colors.white },
        rectRadius: 0.2
    });

    slide3.addText('端午限定福利来啦！\n\n神秘大明星闪现练秋湖\n为华为员工送端午祝福\n现场派发限量粽子礼盒\n更有惊喜互动环节...\n\n转发此消息，粽礼等你拿！', {
        x: 1, y: 2, w: 5.1, h: 2.6,
        fontSize: 18, fontFace: 'Arial',
        color: colors.dark, align: 'left', valign: 'top'
    });

    // 右侧镂空区
    slide3.addShape(pptx.shapes.OVAL, {
        x: 6.8, y: 2.2, w: 2.2, h: 2.8,
        fill: { color: colors.dark },
        line: { color: '9B5DE5', width: 4, dashType: 'dash' }
    });

    slide3.addText('📸\n在此留念', {
        x: 6.8, y: 3.2, w: 2.2, h: 1,
        fontSize: 16, fontFace: 'Arial',
        color: '9B5DE5', align: 'center'
    });

    // ========== 第4页：码农端午生存指南 ==========
    let slide4 = pptx.addSlide();
    slide4.background = { color: '1D3557' };

    // 标题栏
    slide4.addShape(pptx.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 1.3,
        fill: { color: '457B9D' }
    });

    slide4.addText('💻 技 术 版 💻', {
        x: 0, y: 0.1, w: 10, h: 0.4,
        fontSize: 12, fontFace: 'Arial',
        color: colors.white, align: 'center'
    });

    slide4.addText('码农端午生存指南 v2.0', {
        x: 0, y: 0.45, w: 10, h: 0.7,
        fontSize: 32, fontFace: 'Arial', bold: true,
        color: colors.white, align: 'center'
    });

    // 攻略卡片
    const tips = [
        { emoji: '🏃', text: '逃离代码forest', color: '06D6A0' },
        { emoji: '🍙', text: '狂炫粽子', color: 'FFD166' },
        { emoji: '🎮', text: '摸鱼一整天', color: 'FF9F1C' },
        { emoji: '📸', text: '拍照打卡', color: 'E63946' }
    ];

    tips.forEach((tip, i) => {
        const x = 0.6 + i * 2.35;
        slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: x, y: 1.6, w: 2.1, h: 2.4,
            fill: { color: tip.color },
            rectRadius: 0.15
        });

        slide4.addText(tip.emoji, {
            x: x, y: 1.8, w: 2.1, h: 0.9,
            fontSize: 40, align: 'center'
        });

        slide4.addText(tip.text, {
            x: x, y: 2.8, w: 2.1, h: 0.9,
            fontSize: 16, fontFace: 'Arial', bold: true,
            color: colors.dark, align: 'center', valign: 'middle'
        });
    });

    // 底部镂空
    slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: 3.5, y: 4.2, w: 3, h: 1.3,
        fill: { color: '1D3557' },
        line: { color: '06D6A0', width: 4, dashType: 'dash' },
        rectRadius: 0.2
    });

    slide4.addText('📸 合影区', {
        x: 3.5, y: 4.5, w: 3, h: 0.8,
        fontSize: 18, fontFace: 'Arial', bold: true,
        color: '06D6A0', align: 'center'
    });

    // ========== 第5页：情感鸡汤 ==========
    let slide5 = pptx.addSlide();
    slide5.background = { color: 'F8E9DD' };

    // 装饰条
    slide5.addShape(pptx.shapes.RECTANGLE, {
        x: 0, y: 0, w: 0.3, h: 5.6,
        fill: { color: 'E8A87C' }
    });

    slide5.addShape(pptx.shapes.RECTANGLE, {
        x: 9.7, y: 0, w: 0.3, h: 5.6,
        fill: { color: 'E8A87C' }
    });

    // 标题
    slide5.addText('🌿 生 活 版 🌿', {
        x: 0.5, y: 0.3, w: 9, h: 0.4,
        fontSize: 14, fontFace: 'Arial',
        color: 'C38D9E', align: 'center'
    });

    slide5.addText('端午限定', {
        x: 0.5, y: 0.7, w: 9, h: 0.8,
        fontSize: 44, fontFace: 'Georgia', bold: true,
        color: 'C38D9E', align: 'center'
    });

    slide5.addText('遇见更好的自己', {
        x: 0.5, y: 1.4, w: 9, h: 0.6,
        fontSize: 28, fontFace: 'Georgia', italic: true,
        color: '85C7B3', align: 'center'
    });

    // 引用框
    slide5.addShape(pptx.shapes.RECTANGLE, {
        x: 1.5, y: 2.2, w: 7, h: 1.8,
        fill: { color: colors.white },
        line: { color: 'C38D9E', width: 2 }
    });

    slide5.addText('"端午的每一颗粽子\n都是对生活的热爱"\n\n愿你我都能在这个端午\n收获满满的能量 🌟', {
        x: 1.5, y: 2.3, w: 7, h: 1.6,
        fontSize: 20, fontFace: 'Georgia', italic: true,
        color: colors.dark, align: 'center', valign: 'middle'
    });

    // 镂空 - 心形暗示
    slide5.addShape(pptx.shapes.OVAL, {
        x: 4, y: 4.1, w: 2, h: 2,
        fill: { color: 'F8E9DD' },
        line: { color: 'E8A87C', width: 4, dashType: 'dash' }
    });

    slide5.addText('📸\n留下美好', {
        x: 4, y: 4.6, w: 2, h: 1,
        fontSize: 14, fontFace: 'Arial',
        color: 'E8A87C', align: 'center'
    });

    // ========== 第6页：复古报纸风 ==========
    let slide6 = pptx.addSlide();
    slide6.background = { color: 'FFF8E7' };

    // 古典边框
    slide6.addShape(pptx.shapes.RECTANGLE, {
        x: 0.2, y: 0.2, w: 9.6, h: 5.2,
        fill: { color: 'FFF8E7' },
        line: { color: colors.dark, width: 3 }
    });

    slide6.addShape(pptx.shapes.RECTANGLE, {
        x: 0.35, y: 0.35, w: 9.3, h: 4.9,
        fill: { color: 'FFF8E7' },
        line: { color: colors.dark, width: 1 }
    });

    // 报名头
    slide6.addShape(pptx.shapes.RECTANGLE, {
        x: 0.5, y: 0.5, w: 9, h: 1,
        fill: { color: colors.dark }
    });

    slide6.addText('练 秋 湖 日 报', {
        x: 0.5, y: 0.55, w: 9, h: 0.6,
        fontSize: 36, fontFace: 'SimSun', bold: true,
        color: colors.white, align: 'center', charSpacing: 8
    });

    slide6.addText('第1期 · 端午特刊 · 2024年', {
        x: 0.5, y: 1.1, w: 9, h: 0.35,
        fontSize: 12, fontFace: 'SimSun',
        color: colors.white, align: 'center'
    });

    // 分隔线
    slide6.addShape(pptx.shapes.LINE, {
        x: 0.8, y: 1.6, w: 8.4, h: 0,
        line: { color: colors.dark, width: 2 }
    });

    // 主要内容
    slide6.addText('端午节的由来', {
        x: 0.8, y: 1.8, w: 4, h: 0.5,
        fontSize: 22, fontFace: 'SimSun', bold: true,
        color: colors.dark
    });

    slide6.addText('端午节源于纪念爱国诗人屈原。相传公元前278年，屈原于农历五月初五投汨罗江殉国，后人便以此日纪念他，包粽子、赛龙舟的习俗流传至今。', {
        x: 0.8, y: 2.3, w: 4.2, h: 1.5,
        fontSize: 14, fontFace: 'SimSun',
        color: colors.dark, align: 'left', valign: 'top'
    });

    // 右侧新闻
    slide6.addText('📢 今日要闻', {
        x: 5.3, y: 1.8, w: 4, h: 0.4,
        fontSize: 18, fontFace: 'SimSun', bold: true,
        color: colors.red
    });

    slide6.addText('华为上研分部端午活动\n将于练秋湖盛大举行\n现场粽子、互动游戏\n精美礼品等你来拿！', {
        x: 5.3, y: 2.2, w: 4.2, h: 1.3,
        fontSize: 14, fontFace: 'SimSun',
        color: colors.dark
    });

    // 底部镂空
    slide6.addShape(pptx.shapes.RECTANGLE, {
        x: 3.5, y: 4, w: 3, h: 1,
        fill: { color: 'FFF8E7' },
        line: { color: colors.dark, width: 3, dashType: 'dash' }
    });

    slide6.addText('📸 合影留念', {
        x: 3.5, y: 4.3, w: 3, h: 0.5,
        fontSize: 16, fontFace: 'SimSun', bold: true,
        color: colors.dark, align: 'center'
    });

    // 保存
    await pptx.writeFile({ fileName: '端午节活动报纸机.pptx' });
    console.log('✅ 端午节活动报纸机PPT创建成功！');
}

createNewspaper().catch(console.error);