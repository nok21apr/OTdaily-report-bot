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
    to:   process.env.EMAIL_TO || 'naruesit_jit@ttkasia.co.th',
    subject: 'Daily Overtime Report (CSV)',
    text: 'Attached is the requested report in CSV UTF-8 format.'
};

const WEB_CONFIG = {
    user: process.env.WEB_USER, 
    pass: process.env.WEB_PASS
};

(async () => {
    const downloadPath = path.resolve(__dirname, 'downloads');
    if (!WEB_CONFIG.user || !WEB_CONFIG.pass) {
        console.error('❌ Error: Secrets incomplete.');
        process.exit(1);
    }
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    const browser = await puppeteer.launch({
        headless: "new",
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    let page = await browser.newPage();

    try {
        await page.emulateTimezone('Asia/Bangkok');
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
        console.log('✅ Login Success');

        // ---------------------------------------------------------
        // 2. Navigation
        // ---------------------------------------------------------
        console.log('📂 Navigating to Report...');

        const imgLeaveSelector = '#ctl00_ContentPlaceHolder1_imgLeave';
        if (await page.$(imgLeaveSelector)) {
            await Promise.all([
                page.waitForNavigation({ waitUntil: 'networkidle2' }),
                page.click(imgLeaveSelector)
            ]);
        }

        // เลือกเมนู รายงาน
        const parentMenuSelector = '#ctl00_Report_Menu > a';
        await page.waitForSelector(parentMenuSelector, { visible: true, timeout: 30000 });
        await page.click(parentMenuSelector);
        
        // รอเมนูเลื่อนลงมา
        await new Promise(r => setTimeout(r, 1000));

        // คลิกเมนูย่อย
        const subMenuSelector = '#ctl00_Report_Menu > ul a'; 
        await page.waitForSelector(subMenuSelector, { visible: true });
        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2' }),
            page.click(subMenuSelector)
        ]);

        console.log('✅ Arrived at Report Page.');

        // ---------------------------------------------------------
        // 3. Fill Form & Date Logic (แก้ไขจุดที่มีปัญหา)
        // ---------------------------------------------------------
        console.log('📝 Filling form...');
        
        // 3.1 เลือกประเภทเอกสาร
        const docTypeSelector = '#ctl00_ContentPlaceHolder1_ddlDoctype';
        await page.waitForSelector(docTypeSelector, { visible: true });
        await page.select(docTypeSelector, '1');
        
        // 🛑 (จุดสำคัญ) รอให้หน้าเว็บโหลดใหม่ (PostBack) หลังจากเลือกค่าแรก
        console.log('   Waiting for page update...');
        await new Promise(r => setTimeout(r, 3000)); // รอ 3 วินาทีให้ชัวร์

        // 3.2 เลือกประเภท OT
        const otTypeSelector = '#ctl00_ContentPlaceHolder1_ddlOt';
        // เช็คก่อนว่ามี Dropdown นี้ไหม (บางทีเลือกเอกสารอื่นแล้วอันนี้หายไป)
        if (await page.$(otTypeSelector) !== null) {
            await page.select(otTypeSelector, '14');
            await new Promise(r => setTimeout(r, 2000)); // รออีกรอบ
        }

        // 3.3 ใส่วันที่
        const now = new Date();
        const thaiYear = now.getFullYear() + 543;
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const firstDayValue = `01/${month}/${thaiYear}`;
        
        console.log(`📅 Setting date to: ${firstDayValue}`);
        
        const dateInputSelector = '#ctl00_ContentPlaceHolder1_txtFromDate';
        await page.waitForSelector(dateInputSelector);
        
        // 🟢 ใช้วิธีล้างค่าด้วย JS (แม่นยำกว่าการคลิก)
        await page.$eval(dateInputSelector, el => el.value = '');
        
        // พิมพ์ค่าใหม่ลงไป
        await page.type(dateInputSelector, firstDayValue, { delay: 100 });
        
        // กด Tab เพื่อให้เว็บรู้ว่าเราพิมพ์เสร็จแล้ว (สำคัญมาก!)
        await page.keyboard.press('Tab');

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
        await page.click(exportBtnSelector);

        await page.waitForSelector('select', { visible: true });
        
        const selectId = await page.evaluate(() => {
            const options = Array.from(document.querySelectorAll('option'));
            const target = options.find(o => o.text.includes('Microsoft Excel Workbook Data-only'));
            return target ? target.parentElement.id : null;
        });
        
        if (selectId) {
            await page.select(`#${selectId}`, 'Microsoft Excel Workbook Data-only');
        } else {
            await page.keyboard.press('ArrowDown');
        }
        await new Promise(r => setTimeout(r, 1000));

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
            }
        } catch (e) {}

        if (browser) await browser.close();
        process.exit(1);
    }
})();
