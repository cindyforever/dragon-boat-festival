const pptxgen = require('pptxgenjs');
const html2pptx = require('C:/Users/x00336913/AppData/Roaming/GTS-Cowork/sdk_config/.claude/skills/pptx/scripts/html2pptx');
const path = require('path');

const WORK_DIR = 'C:/Users/x00336913/Desktop/新建文件夹/dragonboat_ppt';

async function regenerateSlide5() {
    console.log('重新生成Slide 5...');

    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    // Read existing PPT
    const existingPath = path.join(WORK_DIR, '端午节活动方案汇报.pptx');
    
    try {
        await html2pptx(path.join(WORK_DIR, 'slide5.html'), pptx);
        console.log('Slide 5 生成成功');

        const outputPath = path.join(WORK_DIR, '端午节活动方案汇报_final.pptx');
        await pptx.writeFile({ fileName: outputPath });
        console.log(`PPT已保存至: ${outputPath}`);
    } catch (err) {
        console.error('Slide 5 生成失败:', err.message);
    }
}

regenerateSlide5().catch(console.error);