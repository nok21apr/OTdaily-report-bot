require('dotenv').config();
const puppeteer = require('puppeteer');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// --- ตั้งค่าต่าง ๆ ---
const EMAIL_CONFIG = {
    user: process.env.GMAIL_USER,
    pass: process.env.GMAIL_PASS,
    to:   process.env.EMAIL_TO || 'naruesit_jit@ttkasia.co.th', // ใช้ค่าจาก .env หรือใส่ตรงนี้
    subject: 'Daily Overtime Report (CSV)',
    text: 'Attached is the requested report in CSV UTF-8 format.'
};

const WEB_CONFIG = {
    user: process.env.WEB_USER, 
    pass: process.env.WEB_PASS
};

(async () => {
    // เตรียมโฟลเดอร์
    const downloadPath = path.resolve(__dirname, 'downloads');
    if (!WEB_CONFIG.user || !WEB_CONFIG.pass) {
        console.error('❌ Error: Secrets incomplete. Check .env or GitHub Secrets.');
        process.exit(1);
    }
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    // เปิด Browser
    const browser = await puppeteer.launch({
        headless: "new",
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    let page = await browser.newPage();

    try {
        await page.emulateTimezone('Asia/Bangkok');
        // ตั้งขนาดจอให้ใหญ่หน่อย เพื่อให้เมนูไม่ถูกย่อ
        await page.setViewport({ width: 1366, height: 768 }); 

        const client = await page.target().createCDPSession();
        await client.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        console.log('🚀 Starting process...');

        // ---------------------------------------------------------
        // 1. Login Process
        // ---------------------------------------------------------
        console.log('🔑 Logging in...');
        await page.goto('https://leave.ttkasia.co.th/Login/Login.aspx', { waitUntil: 'networkidle2', timeout: 60000 });
        
        await page.waitForSelector('#txtUsername', { visible: true });
        await page.type('#txtUsername', WEB_CONFIG.user.trim(), { delay: 50 });
        await page.type('#txtPassword', WEB_CONFIG.pass.trim(), { delay: 50 });
        
        console.log('   Pressing Enter to login...');
        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 60000 }),
            page.keyboard.press('Enter')
        ]);
        console.log('✅ Login Success (Current URL: ' + page.url() + ')');

        // ---------------------------------------------------------
        // 2. Navigation (แก้ไขจุดที่ Error)
        // ---------------------------------------------------------
        console.log('📂 Navigating to Report...');

        // เช็คว่าต้องกดปุ่ม Leave ไอคอนใหญ่หรือไม่ (บางที Login แล้วข้ามหน้านี้ไปเลย)
        const imgLeaveSelector = '#ctl00_ContentPlaceHolder1_imgLeave';
        const isImgLeaveVisible = await page.$(imgLeaveSelector);

        if (isImgLeaveVisible) {
            console.log('   Clicking Leave Icon...');
            await Promise.all([
                page.waitForNavigation({ waitUntil: 'networkidle2' }),
                page.click(imgLeaveSelector)
            ]);
        } else {
            console.log('   Leave Icon not found (Skipping...)');
        }

        // --- ส่วนเลือกเมนู รายงาน ---
        console.log('   Selecting Menu Report...');
        const parentMenuSelector = '#ctl00_Report_Menu > a';
        
        // รอจนกว่าเมนู "รายงาน" จะโผล่มา
        await page.waitForSelector(parentMenuSelector, { visible: true, timeout: 30000 });
        
        // กดเมนูหลัก (รายงาน)
        await page.click(parentMenuSelector);
        
        // 🟡 สำคัญ: รอ 1 วินาที ให้เมนูค่อยๆ เลื่อนลงมา (Animation)
        await new Promise(r => setTimeout(r, 1000));

        // คลิกเมนูย่อย
        const subMenuSelector = '#ctl00_Report_Menu > ul a'; 
        console.log('   Clicking Sub-Menu...');
        await page.waitForSelector(subMenuSelector, { visible: true });
        
        // กดและรอหน้าเปลี่ยน
        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2' }),
            page.click(subMenuSelector)
        ]);

        console.log('✅ Arrived at Report Page.');

        // ---------------------------------------------------------
        // 3. Fill Form & Date Logic
        // ---------------------------------------------------------
        console.log('📝 Filling form...');
        // รอให้ Dropdown โผล่มาก่อน
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_ddlDoctype', { visible: true, timeout: 30000 });
        
        // เลือกประเภทเอกสาร = 1
        await page.select('#ctl00_ContentPlaceHolder1_ddlDoctype', '1');
        await new Promise(r => setTimeout(r, 500)); // พักนิดนึง
        
        // เลือกประเภท OT = 14
        await page.select('#ctl00_ContentPlaceHolder1_ddlOt', '14');

        // คำนวณวันที่
        const now = new Date();
        const thaiYear = now.getFullYear() + 543;
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const firstDayValue = `01/${month}/${thaiYear}`;
        
        console.log(`📅 Setting date: ${firstDayValue}`);
        await page.click('#ctl00_ContentPlaceHolder1_txtFromDate', { clickCount: 3 });
        await page.type('#ctl00_ContentPlaceHolder1_txtFromDate', firstDayValue);

        // ---------------------------------------------------------
        // 4. Generate Report
        // ---------------------------------------------------------
        console.log('⏳ Generating Report...');
        
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');
        
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        page = reportPage; 
        await page.bringToFront();
        await page.setViewport({ width: 1280, height: 800 });

        const reportClient = await page.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // 5. Crystal Report Export
        // ---------------------------------------------------------
        console.log('💾 Handling Crystal Report Export...');

        const exportBtnSelector = 'a[title="Export this report"], img[alt="Export this report"]';
        await page.waitForSelector(exportBtnSelector, { visible: true, timeout: 60000 });
        
        console.log('   Clicking Export Icon...');
        await page.click(exportBtnSelector);

        // รอ Dropdown
        await page.waitForSelector('select', { visible: true });
        
        const selectId = await page.evaluate(() => {
            const options = Array.from(document.querySelectorAll('option'));
            const target = options.find(o => o.text.includes('Microsoft Excel Workbook Data-only'));
            return target ? target.parentElement.id : null;
        });
        
        if (selectId) {
            await page.select(`#${selectId}`, 'Microsoft Excel Workbook Data-only');
        } else {
            console.warn('⚠️ Warning: Could not find option by text, trying arrow keys...');
            await page.keyboard.press('ArrowDown');
        }
        await new Promise(r => setTimeout(r, 1000));

        // กดปุ่ม Export สุดท้าย
        const finalSubmitSelector = 'a[id$="_dialog_submitBtn"]';
        await page.waitForSelector(finalSubmitSelector, { visible: true });
        await page.click(finalSubmitSelector);

        // ---------------------------------------------------------
        // 6. Wait for Download
        // ---------------------------------------------------------
        console.log('⬇️ Waiting for file...');
        let downloadedFile;
        for (let i = 0; i < 60; i++) {
            await new Promise(r => setTimeout(r, 1000));
            if (fs.existsSync(downloadPath)) {
                const files = fs.readdirSync(downloadPath);
                downloadedFile = files.find(file => !file.endsWith('.crdownload') && (file.endsWith('.xls') || file.endsWith('.xlsx')));
                if (downloadedFile) break;
            }
        }

        if (!downloadedFile) throw new Error('Download failed or timed out');
        const originalFilePath = path.join(downloadPath, downloadedFile);
        console.log(`✅ Excel downloaded: ${originalFilePath}`);
        
        await browser.close(); 

        // ---------------------------------------------------------
        // 7. Convert & Email
        // ---------------------------------------------------------
        console.log('🔄 Converting to CSV UTF-8...');
        const workbook = xlsx.readFile(originalFilePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const csvContent = xlsx.utils.sheet_to_csv(worksheet);
        
        const csvFileName = downloadedFile.replace(/\.[^/.]+$/, "") + ".csv";
        const csvFilePath = path.join(downloadPath, csvFileName);

        fs.writeFileSync(csvFilePath, '\uFEFF' + csvContent, { encoding: 'utf8' });
        console.log(`✅ Converted to: ${csvFilePath}`);

        fs.unlinkSync(originalFilePath); 

        console.log('📧 Sending email...');
        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: {
                user: EMAIL_CONFIG.user,
                pass: EMAIL_CONFIG.pass
            }
        });

        await transporter.sendMail({
            from: `"Auto Reporter" <${EMAIL_CONFIG.user}>`,
            to: EMAIL_CONFIG.to,
            subject: EMAIL_CONFIG.subject,
            text: EMAIL_CONFIG.text,
            attachments: [{ filename: csvFileName, path: csvFilePath }]
        });

        console.log('✅ Email sent successfully!');

    } catch (error) {
        console.error('❌ Error occurred:', error);
        try {
            if (page && !page.isClosed()) {
                await page.screenshot({ path: 'error_screenshot.png', fullPage: true });
                console.log('📸 Debug screenshot saved: error_screenshot.png');
            }
        } catch (e) {}

        if (browser) await browser.close();
        process.exit(1);
    }
})();
