const pptxgen = require('pptxgenjs');
const html2pptx = require('C:/Users/x00336913/AppData/Roaming/GTS-Cowork/sdk_config/.claude/skills/pptx/scripts/html2pptx');
const fs = require('fs');
const path = require('path');

const WORK_DIR = 'C:/Users/x00336913/Desktop/新建文件夹/dragonboat_ppt';

async function generateAllSlides() {
    console.log('开始生成PPT（修正版）...');

    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.title = '2026年端午节活动方案';
    pptx.author = '服务与软件研发管理部上研分部';

    // ========== Slide 1: 封面页 ==========
    const slide1Html = `<!DOCTYPE html>
<html>
<head>
<style>
html { background: #ffffff; }
body {
    width: 720pt; height: 405pt; margin: 0; padding: 0;
    background: #1C3A5F;
    font-family: Arial, sans-serif;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    position: relative;
    overflow: hidden;
}
.top-decoration {
    position: absolute;
    top: 20pt;
    right: 40pt;
    width: 150pt;
    height: 80pt;
    background: #2D5A27;
    border-radius: 0 0 0 50pt;
    opacity: 0.8;
}
.bottom-decoration {
    position: absolute;
    bottom: 0;
    left: 0;
    width: 300pt;
    height: 150pt;
    background: #4A7C59;
    border-radius: 0 100pt 0 0;
    opacity: 0.6;
}
.main-title {
    position: relative;
    z-index: 10;
    text-align: center;
    margin-bottom: 15pt;
}
.title-zong {
    font-size: 44pt;
    font-weight: bold;
    color: #F5B841;
    margin: 0;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
}
.subtitle-line {
    font-size: 22pt;
    color: #FFF8E7;
    margin: 5pt 0;
}
.info-box {
    background: rgba(255,248,231,0.15);
    border: 1pt solid #F5B841;
    border-radius: 8pt;
    padding: 12pt 30pt;
    margin-top: 20pt;
}
.dept-text {
    font-size: 14pt;
    color: #FFF8E7;
    margin: 3pt 0;
}
.date-text {
    font-size: 14pt;
    color: #FFF8E7;
    margin: 3pt 0;
}
</style>
</head>
<body>
<div class="top-decoration"></div>
<div class="bottom-decoration"></div>

<div class="main-title">
    <h1 class="title-zong">粽情端午，"码"上启航</h1>
    <p class="subtitle-line">——上研分部2026年端午节活动方案</p>
</div>

<div class="info-box">
    <p class="dept-text">服务与软件研发管理部 上研分部</p>
    <p class="date-text">2026年6月18日 (周五) 12:00 - 13:00</p>
</div>
</body>
</html>`;

    fs.writeFileSync(path.join(WORK_DIR, 'slide1.html'), slide1Html);
    console.log('Slide 1 HTML已创建');

    // ========== Slide 2: 活动议程 ==========
    const slide2Html = `<!DOCTYPE html>
<html>
<head>
<style>
html { background: #ffffff; }
body {
    width: 720pt; height: 405pt; margin: 0; padding: 0;
    background: #FFF8E7;
    font-family: Arial, sans-serif;
    display: flex;
    flex-direction: column;
    position: relative;
    overflow: hidden;
}
.header-bar {
    background: #2D5A27;
    height: 60pt;
    display: flex;
    align-items: center;
    padding: 0 40pt;
}
.title {
    font-size: 28pt;
    font-weight: bold;
    color: #FFF8E7;
    margin: 0;
}
.content-area {
    display: flex;
    flex: 1;
    padding: 30pt 40pt;
    gap: 25pt;
}
.agenda-item {
    flex: 1;
    background: white;
    border-radius: 12pt;
    padding: 25pt 20pt;
    display: flex;
    flex-direction: column;
    align-items: center;
    box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
    position: relative;
}
.num-badge {
    position: absolute;
    top: -15pt;
    width: 30pt;
    height: 30pt;
    background: #F5B841;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 14pt;
    font-weight: bold;
    color: white;
}
.icon-circle {
    width: 60pt;
    height: 60pt;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 15pt;
    margin-top: 10pt;
}
.icon-flag { background: #E8F5E9; }
.icon-clock { background: #E3F2FD; }
.icon-game { background: #FFF3E0; }
.icon-budget { background: #FCE4EC; }
.icon-emoji {
    font-size: 28pt;
}
.agenda-title {
    font-size: 16pt;
    font-weight: bold;
    color: #1C3A5F;
    margin: 0 0 8pt 0;
    text-align: center;
}
.agenda-desc {
    font-size: 11pt;
    color: #666;
    margin: 0;
    text-align: center;
    line-height: 1.4;
}
</style>
</head>
<body>
<div class="header-bar">
    <h1 class="title">活动议程一览</h1>
</div>

<div class="content-area">
    <div class="agenda-item">
        <div class="num-badge"><p>1</p></div>
        <div class="icon-circle icon-flag">
            <p class="icon-emoji">🎯</p>
        </div>
        <h3 class="agenda-title">活动主题与目标</h3>
        <p class="agenda-desc">传承·创新·聚力·连接</p>
    </div>
    
    <div class="agenda-item">
        <div class="num-badge"><p>2</p></div>
        <div class="icon-circle icon-clock">
            <p class="icon-emoji">⏱️</p>
        </div>
        <h3 class="agenda-title">活动流程总览</h3>
        <p class="agenda-desc">时间安排与项目导览</p>
    </div>
    
    <div class="agenda-item">
        <div class="num-badge"><p>3</p></div>
        <div class="icon-circle icon-game">
            <p class="icon-emoji">🎮</p>
        </div>
        <h3 class="agenda-title">精彩项目预览</h3>
        <p class="agenda-desc">线下体验 + 线上互动</p>
    </div>
    
    <div class="agenda-item">
        <div class="num-badge"><p>4</p></div>
        <div class="icon-circle icon-budget">
            <p class="icon-emoji">💰</p>
        </div>
        <h3 class="agenda-title">分工与预算说明</h3>
        <p class="agenda-desc">周密筹备，保障落地</p>
    </div>
</div>
</body>
</html>`;

    fs.writeFileSync(path.join(WORK_DIR, 'slide2.html'), slide2Html);
    console.log('Slide 2 HTML已创建');

    // ========== Slide 3: 活动主题与目标 ==========
    const slide3Html = `<!DOCTYPE html>
<html>
<head>
<style>
html { background: #ffffff; }
body {
    width: 720pt; height: 405pt; margin: 0; padding: 0;
    background: #F4F1DE;
    font-family: Arial, sans-serif;
    display: flex;
    flex-direction: column;
    position: relative;
    overflow: hidden;
}
.header-bar {
    background: #2D5A27;
    height: 55pt;
    display: flex;
    align-items: center;
    padding: 0 40pt;
}
.title {
    font-size: 24pt;
    font-weight: bold;
    color: #FFF8E7;
    margin: 0;
}
.content-grid {
    display: flex;
    flex: 1;
    padding: 20pt 30pt;
    gap: 15pt;
}
.theme-card {
    flex: 1;
    background: white;
    border-radius: 12pt;
    padding: 15pt 12pt;
    display: flex;
    flex-direction: column;
    align-items: center;
    box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
}
.card-accent {
    width: 100%;
    height: 5pt;
    border-radius: 3pt;
    margin-bottom: 10pt;
}
.accent-chuan { background: #2D5A27; }
.accent-chuang { background: #F5B841; }
.accent-ju { background: #E07A5F; }
.accent-lian { background: #2C5F7C; }
.keyword {
    font-size: 18pt;
    font-weight: bold;
    color: #1C3A5F;
    margin: 8pt 0 3pt 0;
}
.keyword-desc {
    font-size: 9pt;
    color: #888;
    margin: 0 0 8pt 0;
}
.card-icon {
    width: 45pt;
    height: 45pt;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 8pt;
}
.icon-chuan { background: #E8F5E9; }
.icon-chuang { background: #FFF8E7; }
.icon-ju { background: #FFF3E0; }
.icon-lian { background: #E3F2FD; }
.icon-emoji {
    font-size: 22pt;
}
.detail-text {
    font-size: 9pt;
    color: #555;
    text-align: center;
    line-height: 1.4;
    margin: 0;
    padding: 0 5pt;
}
</style>
</head>
<body>
<div class="header-bar">
    <h1 class="title">我们的主题：传承·创新·聚力·连接</h1>
</div>

<div class="content-grid">
    <div class="theme-card">
        <div class="card-accent accent-chuan"></div>
        <div class="card-icon icon-chuan">
            <p class="icon-emoji">🏮</p>
        </div>
        <h2 class="keyword">传 承</h2>
        <p class="keyword-desc">文化导向</p>
        <p class="detail-text">体验传统习俗，感受端午节文化内涵，增强文化自信</p>
    </div>
    
    <div class="theme-card">
        <div class="card-accent accent-chuang"></div>
        <div class="card-icon icon-chuang">
            <p class="icon-emoji">💡</p>
        </div>
        <h2 class="keyword">创 新</h2>
        <p class="keyword-desc">创新导向</p>
        <p class="detail-text">结合研发特色，打造新颖有趣的线上线下节日体验</p>
    </div>
    
    <div class="theme-card">
        <div class="card-accent accent-ju"></div>
        <div class="card-icon icon-ju">
            <p class="icon-emoji">🤝</p>
        </div>
        <h2 class="keyword">聚 力</h2>
        <p class="keyword-desc">团队导向</p>
        <p class="detail-text">放松身心，促进跨团队交流与协作，提升部门凝聚力</p>
    </div>
    
    <div class="theme-card">
        <div class="card-accent accent-lian"></div>
        <div class="card-icon icon-lian">
            <p class="icon-emoji">💝</p>
        </div>
        <h2 class="keyword">连 接</h2>
        <p class="keyword-desc">温暖导向</p>
        <p class="detail-text">通过线上祝福打破时空限制，加深同事间的温暖连接</p>
    </div>
</div>
</body>
</html>`;

    fs.writeFileSync(path.join(WORK_DIR, 'slide3.html'), slide3Html);
    console.log('Slide 3 HTML已创建');

    // ========== Slide 4: 活动流程总览 ==========
    const slide4Html = `<!DOCTYPE html>
<html>
<head>
<style>
html { background: #ffffff; }
body {
    width: 720pt; height: 405pt; margin: 0; padding: 0;
    background: #E8F5E9;
    font-family: Arial, sans-serif;
    display: flex;
    flex-direction: column;
    position: relative;
    overflow: hidden;
}
.header-bar {
    background: #2D5A27;
    height: 55pt;
    display: flex;
    align-items: center;
    padding: 0 40pt;
}
.title {
    font-size: 24pt;
    font-weight: bold;
    color: #FFF8E7;
    margin: 0;
}
.timeline-container {
    flex: 1;
    padding: 25pt 40pt;
}
.main-timeline {
    position: relative;
    height: 150pt;
    margin-bottom: 25pt;
}
.timeline-line {
    position: absolute;
    top: 45%;
    left: 5%;
    width: 90%;
    height: 4pt;
    background: #4A7C59;
    border-radius: 2pt;
}
.timeline-node {
    position: absolute;
    top: 45%;
    transform: translate(-50%, -50%);
    width: 45pt;
    height: 45pt;
    border-radius: 50%;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    box-shadow: 2px 2px 8px rgba(0,0,0,0.2);
}
.node-icon {
    font-size: 18pt;
}
.node-time {
    font-size: 7pt;
    font-weight: bold;
    color: white;
    margin-top: 1pt;
}
.node-start { background: #F5B841; }
.node-1 { background: #4A7C59; left: 20%; }
.node-2 { background: #E07A5F; left: 40%; }
.node-3 { background: #2C5F7C; left: 60%; }
.node-end { background: #F5B841; left: 85%; }

.node-label {
    position: absolute;
    top: 75%;
    left: 50%;
    transform: translateX(-50%);
    font-size: 9pt;
    color: #1C3A5F;
    font-weight: bold;
    white-space: nowrap;
}
.node-desc {
    position: absolute;
    top: 88%;
    left: 50%;
    transform: translateX(-50%);
    font-size: 7pt;
    color: #666;
    white-space: nowrap;
}

.online-timeline {
    position: relative;
    margin-top: 15pt;
    padding: 12pt 18pt;
    background: #FFF8E7;
    border-radius: 10pt;
    border: 2pt dashed #F5B841;
}
.online-label {
    position: absolute;
    top: -10pt;
    left: 25pt;
    background: #F5B841;
    padding: 2pt 12pt;
    border-radius: 10pt;
    font-size: 9pt;
    font-weight: bold;
    color: white;
}
.online-content {
    display: flex;
    align-items: center;
    gap: 15pt;
    margin-top: 5pt;
}
.online-icon {
    font-size: 26pt;
}
.online-text {
    font-size: 11pt;
    color: #1C3A5F;
    margin: 0;
}
.online-time {
    font-size: 9pt;
    color: #666;
    margin: 0;
    margin-left: auto;
}

.legend {
    position: absolute;
    bottom: 15pt;
    right: 40pt;
    display: flex;
    gap: 12pt;
}
.legend-item {
    display: flex;
    align-items: center;
    gap: 4pt;
}
.legend-dot {
    width: 10pt;
    height: 10pt;
    border-radius: 50%;
}
.legend-text {
    font-size: 8pt;
    color: #666;
}
</style>
</head>
<body>
<div class="header-bar">
    <h1 class="title">活动流程总览 | 6月18日 中午12:00-13:00</h1>
</div>

<div class="timeline-container">
    <div class="main-timeline">
        <div class="timeline-line"></div>
        
        <div class="timeline-node node-start">
            <span class="node-icon">🎬</span>
            <span class="node-time">12:00</span>
        </div>
        <p class="node-label" style="left: 5%;">活动开场</p>
        <p class="node-desc" style="left: 5%;">分组</p>
        
        <div class="timeline-node node-1">
            <span class="node-icon">🧧</span>
            <span class="node-time">12:10</span>
        </div>
        <p class="node-label" style="left: 20%;">主题体验区</p>
        <p class="node-desc" style="left: 20%;">包粽子·香囊</p>
        
        <div class="timeline-node node-2">
            <span class="node-icon">🎮</span>
            <span class="node-time">12:10</span>
        </div>
        <p class="node-label" style="left: 40%;">竞技游戏区</p>
        <p class="node-desc" style="left: 40%;">划龙舟·套圈</p>
        
        <div class="timeline-node node-3">
            <span class="node-icon">📸</span>
            <span class="node-time">12:50</span>
        </div>
        <p class="node-label" style="left: 60%;">合影留念</p>
        <p class="node-desc" style="left: 60%;">礼品发放</p>
        
        <div class="timeline-node node-end">
            <span class="node-icon">🏁</span>
            <span class="node-time">13:00</span>
        </div>
        <p class="node-label" style="left: 85%;">活动结束</p>
    </div>
    
    <div class="online-timeline">
        <span class="online-label">📩 好运粽动员</span>
        <div class="online-content">
            <span class="online-icon">💌</span>
            <div>
                <p class="online-text">线上祝福活动 · 幸福派送中</p>
                <p class="online-time">贯穿全天，随时传递祝福</p>
            </div>
        </div>
    </div>
    
    <div class="legend">
        <div class="legend-item">
            <div class="legend-dot" style="background: #4A7C59;"></div>
            <span class="legend-text">线下体验</span>
        </div>
        <div class="legend-item">
            <div class="legend-dot" style="background: #E07A5F;"></div>
            <span class="legend-text">线下竞技</span>
        </div>
        <div class="legend-item">
            <div class="legend-dot" style="background: #F5B841;"></div>
            <span class="legend-text">线上祝福</span>
        </div>
    </div>
</div>
</body>
</html>`;

    fs.writeFileSync(path.join(WORK_DIR, 'slide4.html'), slide4Html);
    console.log('Slide 4 HTML已创建');

    // ========== Slide 5: 传统技艺体验区 ==========
    const slide5Html = `<!DOCTYPE html>
<html>
<head>
<style>
html { background: #ffffff; }
body {
    width: 720pt; height: 405pt; margin: 0; padding: 0;
    background: #FFF8E7;
    font-family: Arial, sans-serif;
    display: flex;
    flex-direction: column;
    position: relative;
    overflow: hidden;
}
.header-bar {
    background: #2D5A27;
    height: 55pt;
    display: flex;
    align-items: center;
    padding: 0 40pt;
}
.title {
    font-size: 26pt;
    font-weight: bold;
    color: #FFF8E7;
    margin: 0;
}
.content-area {
    display: flex;
    flex: 1;
    padding: 20pt 35pt;
    gap: 25pt;
}
.text-section {
    flex: 1;
    display: flex;
    flex-direction: column;
    justify-content: center;
    gap: 20pt;
}
.activity-item {
    background: white;
    border-radius: 12pt;
    padding: 18pt 22pt;
    box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
    border-left: 5pt solid #4A7C59;
}
.activity-item2 {
    border-left-color: #F5B841;
}
.activity-title {
    font-size: 16pt;
    font-weight: bold;
    color: #2D5A27;
    margin: 0 0 6pt 0;
    display: flex;
    align-items: center;
    gap: 8pt;
}
.activity-title2 {
    color: #F5B841;
}
.activity-icon {
    font-size: 20pt;
}
.activity-desc {
    font-size: 11pt;
    color: #555;
    margin: 0;
    line-height: 1.4;
}
.image-section {
    flex: 1;
    display: flex;
    flex-direction: column;
    gap: 12pt;
}
.image-card {
    flex: 1;
    background: white;
    border-radius: 12pt;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
    position: relative;
    overflow: hidden;
}
.card-overlay {
    position: absolute;
    bottom: 0;
    left: 0;
    right: 0;
    background: #2D5A27;
    padding: 6pt 12pt;
}
.card-overlay2 {
    background: #F5B841;
}
.overlay-text {
    font-size: 10pt;
    color: white;
    margin: 0;
}
.emoji-large {
    font-size: 45pt;
    margin-bottom: 8pt;
}
.emoji-title {
    font-size: 13pt;
    font-weight: bold;
    margin: 0;
}
.title-green {
    color: #2D5A27;
}
.title-yellow {
    color: #F5B841;
}
</style>
</head>
<body>
<div class="header-bar">
    <h1 class="title">手作时光，粽情心意</h1>
</div>

<div class="content-area">
    <div class="text-section">
        <div class="activity-item">
            <h3 class="activity-title">
                <span class="activity-icon">🫔</span>
                包粽子体验
            </h3>
            <p class="activity-desc">专业师傅指导，甜咸粽料自选，体验亲手制作的乐趣</p>
        </div>
        
        <div class="activity-item activity-item2">
            <h3 class="activity-title activity-title2">
                <span class="activity-icon">💐</span>
                香囊DIY
            </h3>
            <p class="activity-desc">配备多种中药材和香囊包，制作专属辟邪祈福香囊</p>
        </div>
    </div>
    
    <div class="image-section">
        <div class="image-card">
            <p class="emoji-large">🫔</p>
            <p class="emoji-title title-green">端午·粽香</p>
            <div class="card-overlay">
                <p class="overlay-text">手工包粽 · 传承文化</p>
            </div>
        </div>
        
        <div class="image-card">
            <p class="emoji-large">💐</p>
            <div class="card-overlay card-overlay2">
                <p class="overlay-text">香囊祈福 · DIY香囊</p>
            </div>
        </div>
    </div>
</div>
</body>
</html>`;

    fs.writeFileSync(path.join(WORK_DIR, 'slide5.html'), slide5Html);
    console.log('Slide 5 HTML已创建');

    // ========== Slide 6: 创意竞技欢乐区 ==========
    const slide6Html = `<!DOCTYPE html>
<html>
<head>
<style>
html { background: #ffffff; }
body {
    width: 720pt; height: 405pt; margin: 0; padding: 0;
    background: #FFF3E0;
    font-family: Arial, sans-serif;
    display: flex;
    flex-direction: column;
    position: relative;
    overflow: hidden;
}
.header-bar {
    background: #E07A5F;
    height: 55pt;
    display: flex;
    align-items: center;
    padding: 0 40pt;
}
.title {
    font-size: 26pt;
    font-weight: bold;
    color: white;
    margin: 0;
}
.content-grid {
    display: flex;
    flex: 1;
    padding: 20pt 35pt;
    gap: 18pt;
}
.game-card {
    flex: 1;
    background: white;
    border-radius: 12pt;
    padding: 18pt 12pt;
    display: flex;
    flex-direction: column;
    align-items: center;
    box-shadow: 2px 2px 12px rgba(0,0,0,0.1);
    position: relative;
}
.game-icon-wrap {
    width: 65pt;
    height: 65pt;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 10pt;
}
.icon-boat { background: #E3F2FD; }
.icon-ring { background: #FFF8E1; }
.icon-photo { background: #FCE4EC; }
.game-icon {
    font-size: 32pt;
}
.game-title {
    font-size: 14pt;
    font-weight: bold;
    color: #1C3A5F;
    margin: 6pt 0 4pt 0;
    text-align: center;
}
.game-desc {
    font-size: 9pt;
    color: #666;
    margin: 0;
    text-align: center;
    line-height: 1.4;
}
.game-tag {
    position: absolute;
    top: -6pt;
    right: 12pt;
    background: #E07A5F;
    color: white;
    font-size: 7pt;
    padding: 2pt 8pt;
    border-radius: 8pt;
    font-weight: bold;
}
.tag-hot { background: #F5B841; }
.tag-fun { background: #4A7C59; }
.footer-text {
    text-align: center;
    padding: 8pt;
}
.footer-note {
    font-size: 9pt;
    color: #888;
    margin: 0;
}
</style>
</head>
<body>
<div class="header-bar">
    <h1 class="title">智勇双全，"码"上争锋</h1>
</div>

<div class="content-grid">
    <div class="game-card">
        <span class="game-tag">团队竞技</span>
        <div class="game-icon-wrap icon-boat">
            <span class="game-icon">🚣</span>
        </div>
        <h3 class="game-title">线上划龙舟比赛</h3>
        <p class="game-desc">通过手机/电脑小程序进行团队划龙舟竞速，科技感十足</p>
    </div>
    
    <div class="game-card">
        <span class="game-tag tag-hot">趣味投掷</span>
        <div class="game-icon-wrap icon-ring">
            <span class="game-icon">🎯</span>
        </div>
        <h3 class="game-title">一举高"粽"</h3>
        <p class="game-desc">用套圈投掷粽子玩偶，寓意"一举高中"，祈福项目顺利</p>
    </div>
    
    <div class="game-card">
        <span class="game-tag tag-fun">欢乐留影</span>
        <div class="game-icon-wrap icon-photo">
            <span class="game-icon">📸</span>
        </div>
        <h3 class="game-title">趣味留影区</h3>
        <p class="game-desc">提供端午主题相框和道具，即时打印照片，留下欢乐瞬间</p>
    </div>
</div>

<div class="footer-text">
    <p class="footer-note">🎮 精彩游戏，等你来战！</p>
</div>
</body>
</html>`;

    fs.writeFileSync(path.join(WORK_DIR, 'slide6.html'), slide6Html);
    console.log('Slide 6 HTML已创建');

    // ========== Slide 7: 线上温情互动区 ==========
    const slide7Html = `<!DOCTYPE html>
<html>
<head>
<style>
html { background: #ffffff; }
body {
    width: 720pt; height: 405pt; margin: 0; padding: 0;
    background: #FFE4B5;
    font-family: Arial, sans-serif;
    display: flex;
    flex-direction: column;
    position: relative;
    overflow: hidden;
}
.header-bar {
    background: #E07A5F;
    height: 55pt;
    display: flex;
    align-items: center;
    padding: 0 40pt;
}
.title {
    font-size: 26pt;
    font-weight: bold;
    color: white;
    margin: 0;
}
.content-area {
    display: flex;
    flex: 1;
    padding: 18pt 35pt;
    gap: 25pt;
}
.main-content {
    flex: 1.2;
    display: flex;
    flex-direction: column;
    gap: 15pt;
}
.card-display {
    flex: 1;
    background: white;
    border-radius: 12pt;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
    position: relative;
}
.card-inner {
    width: 180pt;
    height: 130pt;
    background: #FFF8E7;
    border-radius: 10pt;
    border: 3pt solid #F5B841;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
}
.card-emoji {
    font-size: 40pt;
    margin-bottom: 8pt;
}
.card-text {
    font-size: 16pt;
    font-weight: bold;
    color: #2D5A27;
    margin: 4pt 0;
}
.card-subtext {
    font-size: 10pt;
    color: #888;
    margin: 0;
}
.info-cards {
    display: flex;
    gap: 12pt;
}
.info-card {
    flex: 1;
    background: white;
    border-radius: 10pt;
    padding: 12pt;
    box-shadow: 2px 2px 8px rgba(0,0,0,0.08);
}
.info-icon {
    font-size: 20pt;
    margin-bottom: 6pt;
}
.info-title {
    font-size: 11pt;
    font-weight: bold;
    color: #1C3A5F;
    margin: 0 0 4pt 0;
}
.info-desc {
    font-size: 8pt;
    color: #666;
    margin: 0;
    line-height: 1.3;
}
.side-content {
    flex: 0.8;
    display: flex;
    flex-direction: column;
    gap: 10pt;
}
.step-card {
    background: white;
    border-radius: 10pt;
    padding: 12pt;
    display: flex;
    align-items: center;
    gap: 12pt;
    box-shadow: 2px 2px 8px rgba(0,0,0,0.08);
}
.step-num {
    width: 30pt;
    height: 30pt;
    border-radius: 50%;
    background: #E07A5F;
    color: white;
    font-size: 14pt;
    font-weight: bold;
    display: flex;
    align-items: center;
    justify-content: center;
}
.step-text {
    font-size: 10pt;
    color: #333;
    margin: 0;
}
.highlight-box {
    background: #2D5A27;
    border-radius: 10pt;
    padding: 12pt;
    text-align: center;
}
.highlight-text {
    font-size: 11pt;
    color: white;
    margin: 0;
    font-weight: bold;
}
</style>
</head>
<body>
<div class="header-bar">
    <h1 class="title">💌 好运"粽"动员，幸福派送中</h1>
</div>

<div class="content-area">
    <div class="main-content">
        <div class="card-display">
            <div class="card-inner">
                <p class="card-emoji">🎁</p>
                <p class="card-text">端午安康</p>
                <p class="card-subtext">粽叶飘香 · 幸福绵长</p>
            </div>
        </div>
        
        <div class="info-cards">
            <div class="info-card">
                <p class="info-icon">🔗</p>
                <h4 class="info-title">温情连接</h4>
                <p class="info-desc">打破场地限制，让因工作无法到场的同事也能参与互动</p>
            </div>
            <div class="info-card">
                <p class="info-icon">🎁</p>
                <h4 class="info-title">惊喜体验</h4>
                <p class="info-desc">收到祝福贺卡的惊喜感，增强员工归属感</p>
            </div>
            <div class="info-card">
                <p class="info-icon">🎨</p>
                <h4 class="info-title">文化传递</h4>
                <p class="info-desc">贺卡融入端午元素，传统与现代的美好结合</p>
            </div>
        </div>
    </div>
    
    <div class="side-content">
        <div class="step-card">
            <div class="step-num"><p>1</p></div>
            <p class="step-text">选择精美端午电子贺卡</p>
        </div>
        <div class="step-card">
            <div class="step-num"><p>2</p></div>
            <p class="step-text">撰写个性化祝福语</p>
        </div>
        <div class="step-card">
            <div class="step-num"><p>3</p></div>
            <p class="step-text">定向或随机送给同事</p>
        </div>
        <div class="step-card">
            <div class="step-num"><p>4</p></div>
            <p class="step-text">收集祝福，收获温暖</p>
        </div>
        
        <div class="highlight-box">
            <p class="highlight-text">📩 活动日前上线专属H5平台<br/>全天可参与，随时传递祝福</p>
        </div>
    </div>
</div>
</body>
</html>`;

    fs.writeFileSync(path.join(WORK_DIR, 'slide7.html'), slide7Html);
    console.log('Slide 7 HTML已创建');

    // ========== Slide 8: 分工与预算说明 ==========
    const slide8Html = `<!DOCTYPE html>
<html>
<head>
<style>
html { background: #ffffff; }
body {
    width: 720pt; height: 405pt; margin: 0; padding: 0;
    background: #F4F1DE;
    font-family: Arial, sans-serif;
    display: flex;
    flex-direction: column;
    position: relative;
    overflow: hidden;
}
.header-bar {
    background: #1C3A5F;
    height: 55pt;
    display: flex;
    align-items: center;
    padding: 0 40pt;
}
.title {
    font-size: 26pt;
    font-weight: bold;
    color: #FFF8E7;
    margin: 0;
}
.content-area {
    display: flex;
    flex: 1;
    padding: 18pt 30pt;
    gap: 20pt;
}
.section-left {
    flex: 1.2;
    display: flex;
    flex-direction: column;
    gap: 12pt;
}
.section-title {
    font-size: 14pt;
    font-weight: bold;
    color: #1C3A5F;
    margin: 0;
    display: flex;
    align-items: center;
    gap: 6pt;
}
.section-icon {
    font-size: 16pt;
}
.table-container {
    background: white;
    border-radius: 10pt;
    padding: 12pt;
    box-shadow: 2px 2px 8px rgba(0,0,0,0.08);
}
.table-row {
    display: flex;
    padding: 8pt 0;
    border-bottom: 1pt solid #eee;
}
.table-row:last-child {
    border-bottom: none;
}
.row-label {
    width: 85pt;
    font-size: 10pt;
    font-weight: bold;
    color: #2D5A27;
    display: flex;
    align-items: center;
    gap: 4pt;
}
.row-content {
    flex: 1;
    font-size: 9pt;
    color: #555;
}
.row-icon {
    font-size: 11pt;
}
.section-right {
    flex: 0.8;
    display: flex;
    flex-direction: column;
    gap: 12pt;
}
.budget-card {
    flex: 1;
    background: white;
    border-radius: 12pt;
    padding: 12pt;
    box-shadow: 2px 2px 8px rgba(0,0,0,0.08);
}
.budget-title {
    font-size: 13pt;
    font-weight: bold;
    color: #1C3A5F;
    margin: 0 0 8pt 0;
    text-align: center;
}
.pie-placeholder {
    height: 85pt;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
}
.pie-visual {
    width: 70pt;
    height: 70pt;
    border-radius: 50%;
    background: #2D5A27;
    position: relative;
}
.pie-center {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    width: 35pt;
    height: 35pt;
    background: white;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 9pt;
    font-weight: bold;
    color: #1C3A5F;
}
.budget-legend {
    display: flex;
    flex-wrap: wrap;
    gap: 6pt;
    margin-top: 8pt;
    justify-content: center;
}
.legend-item {
    display: flex;
    align-items: center;
    gap: 3pt;
    font-size: 7pt;
    color: #666;
}
.legend-dot {
    width: 9pt;
    height: 9pt;
    border-radius: 50%;
}
.footer-note {
    background: #E8F5E9;
    border-radius: 8pt;
    padding: 8pt 12pt;
}
.note-text {
    font-size: 8pt;
    color: #2D5A27;
    margin: 0;
}
</style>
</head>
<body>
<div class="header-bar">
    <h1 class="title">周密筹备，保障落地</h1>
</div>

<div class="content-area">
    <div class="section-left">
        <div>
            <h3 class="section-title">
                <span class="section-icon">👥</span>
                活动分工
            </h3>
            <div class="table-container">
                <div class="table-row">
                    <div class="row-label">
                        <span class="row-icon">🎯</span>
                        <p>总协调组</p>
                    </div>
                    <div class="row-content"><p>XXX - 负责整体统筹与进度把控</p></div>
                </div>
                <div class="table-row">
                    <div class="row-label">
                        <span class="row-icon">📦</span>
                        <p>物资组</p>
                    </div>
                    <div class="row-content"><p>XXX - 负责食材、物料采购与分发</p></div>
                </div>
                <div class="table-row">
                    <div class="row-label">
                        <span class="row-icon">📢</span>
                        <p>宣传组</p>
                    </div>
                    <div class="row-content"><p>XXX - 负责现场布置、摄影及线上祝福活动预热引导</p></div>
                </div>
                <div class="table-row">
                    <div class="row-label">
                        <span class="row-icon">💻</span>
                        <p>技术支持组</p>
                    </div>
                    <div class="row-content"><p>XXX - 保障线上比赛流畅及祝福平台稳定运行</p></div>
                </div>
            </div>
        </div>
        
        <div class="footer-note">
            <p class="note-text">💡 提示：各小组需在活动前3天完成准备工作，活动当天提前30分钟到场</p>
        </div>
    </div>
    
    <div class="section-right">
        <div class="budget-card">
            <h3 class="budget-title">💰 活动预算</h3>
            <div class="pie-placeholder">
                <div class="pie-visual">
                    <div class="pie-center"><p>总预算</p></div>
                </div>
            </div>
            <div class="budget-legend">
                <div class="legend-item">
                    <div class="legend-dot" style="background: #2D5A27;"></div>
                    <p>物料费 30%</p>
                </div>
                <div class="legend-item">
                    <div class="legend-dot" style="background: #F5B841;"></div>
                    <p>食材费 25%</p>
                </div>
                <div class="legend-item">
                    <div class="legend-dot" style="background: #E07A5F;"></div>
                    <p>奖品费 25%</p>
                </div>
                <div class="legend-item">
                    <div class="legend-dot" style="background: #2C5F7C;"></div>
                    <p>平台费 20%</p>
                </div>
            </div>
        </div>
    </div>
</div>
</body>
</html>`;

    fs.writeFileSync(path.join(WORK_DIR, 'slide8.html'), slide8Html);
    console.log('Slide 8 HTML已创建');

    // ========== Slide 9: 结束页 ==========
    const slide9Html = `<!DOCTYPE html>
<html>
<head>
<style>
html { background: #ffffff; }
body {
    width: 720pt; height: 405pt; margin: 0; padding: 0;
    background: #2D5A27;
    font-family: Arial, sans-serif;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    position: relative;
    overflow: hidden;
}
.main-content {
    text-align: center;
    z-index: 10;
    position: relative;
}
.main-invite {
    font-size: 40pt;
    font-weight: bold;
    color: #FFF8E7;
    margin: 0 0 12pt 0;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
}
.sub-invite {
    font-size: 15pt;
    color: #FFF8E7;
    margin: 0 0 25pt 0;
    opacity: 0.9;
}
.slogan-box {
    background: rgba(255,248,231,0.2);
    border: 2pt solid #F5B841;
    border-radius: 12pt;
    padding: 18pt 45pt;
    display: inline-block;
}
.slogan-main {
    font-size: 20pt;
    font-weight: bold;
    color: #F5B841;
    margin: 0 0 6pt 0;
}
.slogan-sub {
    font-size: 13pt;
    color: #FFF8E7;
    margin: 0;
}
.corner-deco {
    position: absolute;
    width: 80pt;
    height: 80pt;
}
.corner-tl {
    top: 20pt;
    left: 20pt;
    border-top: 3pt solid #F5B841;
    border-left: 3pt solid #F5B841;
}
.corner-tr {
    top: 20pt;
    right: 20pt;
    border-top: 3pt solid #F5B841;
    border-right: 3pt solid #F5B841;
}
.corner-bl {
    bottom: 20pt;
    left: 20pt;
    border-bottom: 3pt solid #F5B841;
    border-left: 3pt solid #F5B841;
}
.corner-br {
    bottom: 20pt;
    right: 20pt;
    border-bottom: 3pt solid #F5B841;
    border-right: 3pt solid #F5B841;
}
.float-emoji {
    position: absolute;
    font-size: 35pt;
    opacity: 0.3;
}
.emoji-1 { top: 50pt; left: 70pt; }
.emoji-2 { bottom: 80pt; right: 80pt; }
.emoji-3 { bottom: 80pt; left: 50pt; }
.emoji-4 { top: 60pt; right: 100pt; }
</style>
</head>
<body>
<div class="corner-deco corner-tl"></div>
<div class="corner-deco corner-tr"></div>
<div class="corner-deco corner-bl"></div>
<div class="corner-deco corner-br"></div>

<span class="float-emoji emoji-1">🫔</span>
<span class="float-emoji emoji-2">💐</span>
<span class="float-emoji emoji-3">🚣</span>
<span class="float-emoji emoji-4">💌</span>

<div class="main-content">
    <h1 class="main-invite">期待您的参与！</h1>
    <p class="sub-invite">线下欢聚，线上联动，让我们共度一个充满欢声笑语的端午佳节！</p>
    
    <div class="slogan-box">
        <p class="slogan-main">粽叶飘香，代码传情！</p>
        <p class="slogan-sub">祝大家端午安康！</p>
    </div>
</div>
</body>
</html>`;

    fs.writeFileSync(path.join(WORK_DIR, 'slide9.html'), slide9Html);
    console.log('Slide 9 HTML已创建');

    // 开始生成所有幻灯片
    console.log('\n开始转换HTML为PPT...');

    for (let i = 1; i <= 9; i++) {
        const htmlPath = path.join(WORK_DIR, `slide${i}.html`);
        console.log(`转换第 ${i} 张幻灯片...`);
        try {
            await html2pptx(htmlPath, pptx);
            console.log(`  ✓ 第 ${i} 张完成`);
        } catch (err) {
            console.error(`  ✗ 第 ${i} 张失败:`, err.message);
        }
    }

    // 保存PPT文件
    const outputPath = path.join(WORK_DIR, '端午节活动方案汇报.pptx');
    await pptx.writeFile({ fileName: outputPath });
    console.log(`\n✅ PPT已保存至: ${outputPath}`);

    return outputPath;
}

generateAllSlides().catch(console.error);