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
    to:   'naruesit_jit@ttkasia.co.th',
    subject: 'Daily Overtime Report (CSV)',
    text: 'Attached is the requested report in CSV UTF-8 format.'
};

const WEB_CONFIG = {
    user: process.env.WEB_USER, 
    pass: process.env.WEB_PASS
};

(async () => {
    const downloadPath = path.resolve(__dirname, 'downloads');
    
    // ตรวจสอบค่า Config
    if (!WEB_CONFIG.user || !WEB_CONFIG.pass) {
        console.error('❌ Error: Secrets incomplete. Please check .env file.');
        process.exit(1);
    }
    
    // เคลียร์ไฟล์เก่า
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

        console.log('✅ Arrived at Report Form Page.');

        // ---------------------------------------------------------
        // 3. Fill Form & Date Logic
        // ---------------------------------------------------------
        console.log('📝 Filling form...');
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_ddlDoctype', { visible: true });
        await page.select('#ctl00_ContentPlaceHolder1_ddlDoctype', '1');
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
        // 4. Generate Report & 5. Export (Updated from 1.txt)
        // ---------------------------------------------------------
        console.log('⏳ Generating Report...');
        
        // เตรียม Promise สำหรับหน้าต่างใหม่
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        
        // กดปุ่มแสดงรายงาน
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');

        // --- เริ่มต้นส่วนโค้ดจากไฟล์ 1.txt ---
        
        // 12. Select Window (Switch to new tab)
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!"); // เพิ่มเช็คกันเหนียว
        
        await reportPage.bringToFront();
        await reportPage.setViewport({ width: 1366, height: 768 }); // เพิ่มการตั้งขนาดจอให้ชัวร์

        // Setup Download สำหรับหน้าใหม่ (สำคัญมาก)
        const reportClient = await reportPage.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // รอให้หน้า Report โหลดเสร็จ (สำคัญมาก เพราะ Crystal Report โหลดนาน)
        // ใช้ try-catch เผื่อกรณี timeout แต่หน้าโหลดเสร็จแล้ว
        try {
            await reportPage.waitForNavigation({ waitUntil: 'networkidle0', timeout: 60000 }); // เพิ่ม timeout เป็น 60s
        } catch (e) {
            console.log('Navigation wait finished (or timed out), proceeding...');
        }
        
        console.log('Interacting with Report Viewer...');
        
        // 13. Click Export Icon
        // Selector ต้นฉบับ: id=IconImg_รายงานรายละเอียดขออนุมัติใบโอทีของพนักงาน_toptoolbar_export
        // เนื่องจากชื่อ ID ยาวและมีภาษาไทย ควรใช้ Attribute selector หรือ ID เต็มถ้ามั่นใจว่าไม่เปลี่ยน
        console.log('13. Clicking Export Icon...');
        const exportIconSelector = '[id*="toptoolbar_export"]'; // เลือกที่มีคำว่า toptoolbar_export
        await reportPage.waitForSelector(exportIconSelector, { visible: true, timeout: 60000 });
        await reportPage.click(exportIconSelector);

        // 14. Click Format Dropdown
        // ต้นฉบับ: id=IconImg_Txt_iconMenu_icon_bobjid_1765873459217_dialog_combo
        // ปัญหา: ID มีตัวเลข timestamp (1765873459217) ซึ่งจะเปลี่ยนตลอด
        // แก้ไข: ใช้ Selector ที่จับ Pattern แทน
        console.log('14. Opening Export Format Dialog...');
        const formatDropdownSelector = 'div[id*="IconImg_Txt_iconMenu_icon"][id*="dialog_combo"]';
        await reportPage.waitForSelector(formatDropdownSelector, { visible: true });
        await reportPage.click(formatDropdownSelector);

        // 15. Click Specific Format (Item 14)
        // ต้นฉบับ: id=iconMenu_menu_bobjid_..._it_14
        console.log('15. Selecting Export Format (it_14)...');
        // สมมติว่าต้องการเลือก item ลำดับที่น่าจะเป็น Excel หรือ PDF (ในสคริปต์คือ it_14)
        // พยายามหา Element ที่เป็น menu item และลงท้ายด้วย it_14
        const formatItemSelector = 'span[id*="iconMenu_menu"][id*="it_14"]';
        await reportPage.waitForSelector(formatItemSelector, { visible: true });
        await reportPage.click(formatItemSelector);

        // 16. Click Export Button (LinkText = Export)
        console.log('16. Clicking Final Export...');
        const exportBtnXpath = '//a[text()="Export"]';
        
        // ใช้ waitForSelector กับ xpath (Puppeteer เวอร์ชันใหม่ๆ รองรับ xpath นำหน้าด้วย // หรือใช้ waitForXPath ถ้าเวอร์ชันเก่า)
        try {
            await reportPage.waitForSelector(`xpath/${exportBtnXpath}`, { visible: true });
            const exportBtns = await reportPage.$$(`xpath/${exportBtnXpath}`);
            if (exportBtns.length > 0) {
                await exportBtns[0].click();
            } else {
                throw new Error("Export button not found via XPath selector");
            }
        } catch (e) {
            // Fallback: ถ้าใช้ puppeteer รุ่นเก่าที่ต้องใช้ waitForXPath
            try {
                await reportPage.waitForXPath(exportBtnXpath, { visible: true });
                const exportBtn = (await reportPage.$x(exportBtnXpath))[0];
                await exportBtn.click();
            } catch (ex) {
                console.warn('Fallback click failed, trying Enter key...');
                await reportPage.keyboard.press('Enter');
            }
        }

        console.log('Done! Waiting a bit before closing...');
        await new Promise(r => setTimeout(r, 5000));

        // --- จบส่วนโค้ดจากไฟล์ 1.txt ---

        // ---------------------------------------------------------
        // 6. Wait for Download
        // ---------------------------------------------------------
        console.log('⬇️ Waiting for file...');
        let downloadedFile;
        // รอ 5 นาที (300 วินาที) เผื่อไฟล์ใหญ่
        for (let i = 0; i < 300; i++) { 
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

        console.log('📧 Sending email to: ' + EMAIL_CONFIG.to);
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
            if (page && !page.isClosed()) await page.screenshot({ path: 'error_main.png', fullPage: true });
            const pages = await browser.pages();
            if (pages.length > 1) await pages[pages.length-1].screenshot({ path: 'error_report.png', fullPage: true });
        } catch(e){}

        if (browser) await browser.close();
        process.exit(1);
    }
})();
