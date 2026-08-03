// html2pptx 构建脚本：将 slides/*.html 转换为 PPTX
const pptxgen = require('pptxgenjs');
const html2pptx = require('./html2pptx');

async function main() {
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'MelonAI';
    pptx.title = '今日三大AI要闻：打工人视角';

    for (let i = 1; i <= 10; i++) {
        const htmlFile = `slides/slide${i}.html`;
        await html2pptx(htmlFile, pptx);
        console.log(`slide ${i} converted`);
    }

    const out = '今日三大AI要闻_html2pptx.pptx';
    await pptx.writeFile(out);
    console.log('saved:', out);
}

main().catch((e) => { console.error(e); process.exit(1); });
