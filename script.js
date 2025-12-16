const puppeteer = require('puppeteer');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// ... (EMAIL_CONFIG เหมือนเดิม) ...
const EMAIL_CONFIG = {
    user: process.env.GMAIL_USER,        
    pass: process.env.GMAIL_PASS,        
    to:   'naruesit_jit@ttkasia.co.th', // 🔴 อย่าลืมแก้: อีเมลปลายทาง
    subject: 'Daily Overtime Report (CSV)',
    text: 'Attached is the requested report in CSV UTF-8 format.'
};

(async () => {
    // ... (Setup เหมือนเดิม) ...
    const downloadPath = path.resolve(__dirname, 'downloads');
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    const browser = await puppeteer.launch({
        headless: "new",
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage(); // ประกาศ page นอก try เพื่อให้เรียกใช้ใน catch ได้
    
    try {
        await page.emulateTimezone('Asia/Bangkok');
        await page.setViewport({ width: 1280, height: 800 });

        const client = await page.target().createCDPSession();
        await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });

        console.log('🚀 Starting process...');

        // ... (Logic การ Login, เข้าหน้า Report, Download, Convert CSV เหมือนเดิม) ...
        // (Copy ไส้ในจากคำตอบก่อนหน้ามาวางตรงนี้ได้เลยครับ เพื่อความกระชับ)
        // ...
        
        // สมมติว่าจบขั้นตอนส่งเมล
        console.log('✅ Process Completed Successfully.');

    } catch (error) {
        console.error('❌ Error occurred:', error);

        // 📸 Capture Screenshot for Debugging
        try {
            await page.screenshot({ path: 'error_screenshot.png', fullPage: true });
            console.log('📸 Debug screenshot saved as error_screenshot.png');
        } catch (e) {
            console.error('Failed to take screenshot:', e);
        }

        process.exit(1); // แจ้ง GitHub ว่า Job Failed
    } finally {
        await browser.close();
    }
})();
