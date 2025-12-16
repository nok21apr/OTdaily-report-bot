const puppeteer = require('puppeteer');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// --- 1. ตั้งค่า Email (ดึงจาก GitHub Secrets) ---
const EMAIL_CONFIG = {
    user: process.env.GMAIL_USER,
    pass: process.env.GMAIL_PASS,
    to:   'naruesit_jit@ttkasia.co.th', // 🔴 อย่าลืมตรวจสอบอีเมลปลายทาง
    subject: 'Daily Overtime Report (CSV)',
    text: 'Attached is the requested report in CSV UTF-8 format.'
};

// --- 2. ตั้งค่า Web Login (ดึงจาก GitHub Secrets) ---
// 🟢 ส่วนที่แก้ไข: ดึง User/Pass เว็บจาก Secret
const WEB_CONFIG = {
    user: process.env.WEB_USER, 
    pass: process.env.WEB_PASS
};

(async () => {
    const downloadPath = path.resolve(__dirname, 'downloads');
    
    // ตรวจสอบว่ามีรหัสผ่านครบไหมก่อนเริ่ม
    if (!WEB_CONFIG.user || !WEB_CONFIG.pass) {
        console.error('❌ Error: WEB_USER or WEB_PASS secrets are missing!');
        process.exit(1);
    }

    if (fs.existsSync(downloadPath)) {
        fs.rmSync(downloadPath, { recursive: true, force: true });
    }
    fs.mkdirSync(downloadPath);

    const browser = await puppeteer.launch({
        headless: "new",
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    let page = await browser.newPage();

    try {
        await page.emulateTimezone('Asia/Bangkok');
        await page.setViewport({ width: 1280, height: 800 });

        const client = await page.target().createCDPSession();
        await client.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        console.log('🚀 Starting process...');

        // ---------------------------------------------------------
        // 3. Login Process (ใช้ค่าจาก Secret)
        // ---------------------------------------------------------
        console.log('🔑 Logging in...');
        await page.goto('https://leave.ttkasia.co.th/Login/Login.aspx', { waitUntil: 'networkidle0' });
        
        // 🟢 ใช้ตัวแปรแทนการพิมพ์เลขตรงๆ
        await page.type('#txtUsername', WEB_CONFIG.user);
        await page.type('#txtPassword', WEB_CONFIG.pass);
        
        await Promise.all([
            page.waitForNavigation(),
            page.keyboard.press('Enter')
        ]);

        // ---------------------------------------------------------
        // 4. Navigation to Report
        // ---------------------------------------------------------
        console.log('📂 Navigating to Report...');
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_imgLeave');
        await page.click('#ctl00_ContentPlaceHolder1_imgLeave');
        
        await page.waitForSelector('#ctl00_Report_Menu > a');
        await page.click('#ctl00_Report_Menu > a');
        
        const subMenuSelector = '#ctl00_Report_Menu > ul a';
        await page.waitForSelector(subMenuSelector);
        await Promise.all([
            page.waitForNavigation(),
            page.click(subMenuSelector)
        ]);

        // ---------------------------------------------------------
        // 5. Fill Form & Date Logic
        // ---------------------------------------------------------
        console.log('📝 Filling form...');
        await page.select('#ctl00_ContentPlaceHolder1_ddlDoctype', '1');
        await page.select('#ctl00_ContentPlaceHolder1_ddlOt', '14');

        const now = new Date();
        const thaiYear = now.getFullYear() + 543;
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const firstDayValue = `01/${month}/${thaiYear}`;
        
        console.log(`📅 Setting date: ${firstDayValue}`);
        await page.click('#ctl00_ContentPlaceHolder1_txtFromDate', { clickCount: 3 });
        await page.type('#ctl00_ContentPlaceHolder1_txtFromDate', firstDayValue);

        // ---------------------------------------------------------
        // 6. Generate Report & Handle New Tab
        // ---------------------------------------------------------
        console.log('⏳ Generating Report...');
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');
        
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        page = reportPage; // สลับตัวแปร page มาคุมหน้าใหม่
        await page.bringToFront();
        
        const reportClient = await page.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // Click Export
        const exportBtnSelector = 'a[title="Export this report"], img[alt="Export this report"]';
        await page.waitForSelector(exportBtnSelector, { timeout: 60000 });
        console.log('💾 Clicking Export...');
        await page.click(exportBtnSelector);

        // Select Excel
        await page.waitForSelector('select');
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

        // Final Submit
        const finalSubmitSelector = 'a[id$="_dialog_submitBtn"]';
        await page.waitForSelector(finalSubmitSelector);
        await page.click(finalSubmitSelector);

        // ---------------------------------------------------------
        // 7. Wait for Download
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
        // 8. Convert & Email
        // ---------------------------------------------------------
        console.log('🔄 Converting to CSV UTF-8...');
        const workbook = xlsx.readFile(originalFilePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const csvContent = xlsx.utils.sheet_to_csv(worksheet);
        
        const csvFileName = downloadedFile.replace(/\.[^/.]+$/, "") + ".csv";
        const csvFilePath = path.join(downloadPath, csvFileName);

        fs.writeFileSync(csvFilePath, '\uFEFF' + csvContent, { encoding: 'utf8' });
        console.log(`✅ Converted to: ${csvFilePath}`);

        fs.unlinkSync(originalFilePath); // ลบไฟล์ Excel เดิม

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
            if (page) await page.screenshot({ path: 'error_screenshot.png', fullPage: true });
        } catch (e) {}
        await browser.close();
        process.exit(1);
    }
})();
