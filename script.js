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
    to:   'recipient_email@example.com', // 🔴 แก้ไขอีเมลปลายทางที่นี่
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
        // ตั้งขนาดจอตามไฟล์ Recorder ของคุณ (771x791) หรือใหญ่กว่าเพื่อให้เห็นครบ
        await page.setViewport({ width: 1280, height: 800 }); 

        const client = await page.target().createCDPSession();
        await client.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        console.log('🚀 Starting process...');

        // ---------------------------------------------------------
        // 1. Login Process (แก้ตาม Recorder: ใช้ Enter แทนการคลิก)
        // ---------------------------------------------------------
        console.log('🔑 Logging in...');
        await page.goto('https://leave.ttkasia.co.th/Login/Login.aspx', { waitUntil: 'networkidle2', timeout: 60000 });
        
        // รอช่อง User และพิมพ์
        await page.waitForSelector('#txtUsername', { visible: true });
        await page.type('#txtUsername', WEB_CONFIG.user.trim(), { delay: 100 });

        // รอช่อง Password และพิมพ์
        await page.type('#txtPassword', WEB_CONFIG.pass.trim(), { delay: 100 });
        
        console.log('   Pressing Enter to login...');
        // 🟢 แก้ไข: ใช้การกด Enter แบบใน Recorder แทนการหาปุ่มคลิก
        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 60000 }),
            page.keyboard.press('Enter')
        ]);
        console.log('✅ Login Success');

        // ---------------------------------------------------------
        // 2. Navigation (ตาม Flow เดิมและ Recorder)
        // ---------------------------------------------------------
        console.log('📂 Navigating to Report...');
        
        // คลิกไอคอน Leave
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_imgLeave', { visible: true });
        await page.click('#ctl00_ContentPlaceHolder1_imgLeave');
        
        // คลิกเมนูรายงาน (Main Menu)
        await page.waitForSelector('#ctl00_Report_Menu > a', { visible: true });
        await page.click('#ctl00_Report_Menu > a');
        
        // คลิกเมนูย่อย (Sub Menu)
        // ใช้ logic รอ navigation แบบใน Recorder
        const subMenuSelector = '#ctl00_Report_Menu > ul a'; 
        await page.waitForSelector(subMenuSelector, { visible: true });
        
        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2' }),
            page.click(subMenuSelector)
        ]);

        // ---------------------------------------------------------
        // 3. Fill Form & Date Logic
        // ---------------------------------------------------------
        console.log('📝 Filling form...');
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_ddlDoctype');
        
        // เลือกประเภทเอกสาร = 1
        await page.select('#ctl00_ContentPlaceHolder1_ddlDoctype', '1');
        // รอสักนิดเผื่อเว็บโหลด ajax
        await new Promise(r => setTimeout(r, 500));
        
        // เลือกประเภท OT = 14
        await page.select('#ctl00_ContentPlaceHolder1_ddlOt', '14');

        // คำนวณวันที่ 1 ของเดือนปัจจุบัน (พ.ศ.)
        const now = new Date();
        const thaiYear = now.getFullYear() + 543;
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const firstDayValue = `01/${month}/${thaiYear}`;
        
        console.log(`📅 Setting date: ${firstDayValue}`);
        // Recorder ใช้การคลิก แต่เราต้องพิมพ์ค่าใหม่ลงไป
        await page.click('#ctl00_ContentPlaceHolder1_txtFromDate', { clickCount: 3 });
        await page.type('#ctl00_ContentPlaceHolder1_txtFromDate', firstDayValue);

        // ---------------------------------------------------------
        // 4. Generate Report & Handle New Tab
        // ---------------------------------------------------------
        console.log('⏳ Generating Report...');
        
        // เตรียมจับ Event หน้าต่างใหม่
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        
        // กดปุ่มแสดงรายงาน
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');
        
        // รอรับหน้าต่างใหม่
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        // สลับตัวแปร page มาคุมหน้าใหม่
        page = reportPage; 
        await page.bringToFront();
        await page.setViewport({ width: 1280, height: 800 });

        // ตั้งค่า Download ให้หน้าใหม่ด้วย
        const reportClient = await page.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // 5. Crystal Report Export (แก้ Dynamic ID)
        // ---------------------------------------------------------
        console.log('💾 Handling Crystal Report Export...');

        // รอจนปุ่ม Export โผล่
        // ใช้ Selector แบบ Attribute เพื่อหนี Dynamic ID (bobjid_...)
        const exportBtnSelector = 'a[title="Export this report"], img[alt="Export this report"]';
        await page.waitForSelector(exportBtnSelector, { visible: true, timeout: 60000 });
        
        console.log('   Clicking Export Icon...');
        await page.click(exportBtnSelector);

        // รอ Dropdown เลือก Format
        await page.waitForSelector('select', { visible: true });
        
        // เลือก Excel Data-only โดยการหา Text ใน Option (เพราะ ID เปลี่ยนตลอด)
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

        // รอ 1 วินาทีให้ UI อัปเดต
        await new Promise(r => setTimeout(r, 1000));

        // กดปุ่ม Export สุดท้าย
        // Recorder ใช้ ID: bobjid_..._dialog_submitBtn
        // เราใช้ CSS Selector ที่ลงท้ายด้วย _dialog_submitBtn เพื่อรองรับ ID ที่เปลี่ยนไป
        const finalSubmitSelector = 'a[id$="_dialog_submitBtn"]';
        await page.waitForSelector(finalSubmitSelector, { visible: true });
        await page.click(finalSubmitSelector);

        // ---------------------------------------------------------
        // 6. Wait for Download
        // ---------------------------------------------------------
        console.log('⬇️ Waiting for file...');
        let downloadedFile;
        // วนลูปรอ 60 วินาที
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

        // ใส่ BOM (\uFEFF) เพื่อให้ Excel อ่านไทยออก
        fs.writeFileSync(csvFilePath, '\uFEFF' + csvContent, { encoding: 'utf8' });
        console.log(`✅ Converted to: ${csvFilePath}`);

        // ลบไฟล์ต้นฉบับ
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
        
        // ถ่ายรูป Error (ถ้า Browser ยังเปิดอยู่)
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
