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

// 🛠️ ฟังก์ชันพิเศษ: วนลูปหา Element ในทุก Frame (ทั้งหน้าหลักและ Iframe ย่อย)
async function findElementInFrames(page, selector) {
    // 1. ลองหาในหน้าหลักก่อน
    try {
        if (await page.$(selector)) return page;
    } catch (e) {}

    // 2. ถ้าไม่เจอ ให้วนหาในทุก Frame
    for (const frame of page.frames()) {
        try {
            if (await frame.$(selector)) return frame;
        } catch (e) {}
    }
    return null;
}

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

        const parentMenuSelector = '#ctl00_Report_Menu > a';
        await page.waitForSelector(parentMenuSelector, { visible: true, timeout: 30000 });
        await page.click(parentMenuSelector);
        
        await new Promise(r => setTimeout(r, 1000));

        const subMenuSelector = '#ctl00_Report_Menu > ul a'; 
        await page.waitForSelector(subMenuSelector, { visible: true });
        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2' }),
            page.click(subMenuSelector)
        ]);

        console.log('✅ Arrived at Report Page.');

        // ---------------------------------------------------------
        // 3. Fill Form & Date Logic
        // ---------------------------------------------------------
        console.log('📝 Filling form...');
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_ddlDoctype', { visible: true });
        await page.select('#ctl00_ContentPlaceHolder1_ddlDoctype', '1');
        
        console.log('   Waiting for page update...');
        await new Promise(r => setTimeout(r, 3000));

        const otTypeSelector = '#ctl00_ContentPlaceHolder1_ddlOt';
        if (await page.$(otTypeSelector) !== null) {
            await page.select(otTypeSelector, '14');
            await new Promise(r => setTimeout(r, 2000));
        }

        const now = new Date();
        const thaiYear = now.getFullYear() + 543;
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const firstDayValue = `01/${month}/${thaiYear}`;
        
        console.log(`📅 Setting date to: ${firstDayValue}`);
        const dateInputSelector = '#ctl00_ContentPlaceHolder1_txtFromDate';
        await page.waitForSelector(dateInputSelector);
        await page.$eval(dateInputSelector, el => el.value = '');
        await page.type(dateInputSelector, firstDayValue, { delay: 100 });
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
        
        // รอให้หน้า Report โหลดเสร็จจริง ๆ
        await new Promise(r => setTimeout(r, 5000));

        // ตั้งค่า Download ให้หน้าใหม่
        const reportClient = await page.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // 5. Crystal Report Export (แก้ปัญหา Iframe)
        // ---------------------------------------------------------
        console.log('💾 Handling Crystal Report Export...');

        // 5.1 หาปุ่ม Export Icon (อาจอยู่ใน Iframe)
        const exportIconSelector = 'a[title="Export this report"], img[alt="Export this report"]';
        let targetFrame = null;
        
        // วนรอจนกว่าจะเจอ Frame ที่มีปุ่ม Export (รอสูงสุด 60 วิ)
        console.log('   Searching for Export button inside frames...');
        for (let i = 0; i < 30; i++) {
            targetFrame = await findElementInFrames(page, exportIconSelector);
            if (targetFrame) break;
            await new Promise(r => setTimeout(r, 2000)); // รอ 2 วิ แล้วหาใหม่
        }

        if (!targetFrame) throw new Error("Could not find Export button in any frame!");
        
        console.log('   Found Export button! Clicking...');
        await targetFrame.click(exportIconSelector);

        // 5.2 รอ Popup Dialog เด้งขึ้นมา
        console.log('   Waiting for Export Dialog...');
        await new Promise(r => setTimeout(r, 2000));

        // 5.3 เลือก Microsoft Excel (Data-only)
        // หา Dropdown ใน Frame เดิม (เพราะ Dialog มักอยู่ใน Frame เดียวกัน)
        await targetFrame.waitForSelector('select', { timeout: 10000 });
        
        const selectId = await targetFrame.evaluate(() => {
            const options = Array.from(document.querySelectorAll('option'));
            // หา Option ที่มีคำว่า Data-only
            const target = options.find(o => o.text.includes('Data-only') || o.text.includes('Excel'));
            return target ? target.parentElement.id : null;
        });

        if (selectId) {
            console.log(`   Selecting Excel Format (ID: ${selectId})...`);
            await targetFrame.select(`#${selectId}`, 'Microsoft Excel Workbook Data-only');
        } else {
            console.warn('⚠️ Could not find Excel option by text. Trying arrow keys...');
            await page.keyboard.press('ArrowDown');
        }
        
        await new Promise(r => setTimeout(r, 1000));

        // 5.4 กดปุ่ม Export (ปุ่มยืนยันสุดท้าย)
        // ใช้ Selector แบบลงท้ายด้วย _dialog_submitBtn (ตามรูปที่คุณส่งมา)
        const submitBtnSelector = 'a[id$="_dialog_submitBtn"]';
        await targetFrame.waitForSelector(submitBtnSelector, { visible: true });
        
        console.log('   Clicking Final Export Button...');
        await targetFrame.click(submitBtnSelector);

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
