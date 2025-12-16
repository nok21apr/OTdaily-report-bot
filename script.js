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

// 🛠️ ฟังก์ชันพิเศษ: วนลูปหา Element ในทุก Frame (ใช้ XPath ได้ด้วย)
async function findElementInFrames(page, selector, isXPath = false) {
    const frames = [page, ...page.frames()];
    for (const frame of frames) {
        try {
            const element = isXPath 
                ? await frame.$x(selector) 
                : await frame.$(selector);
            
            if (element && (isXPath ? element.length > 0 : element)) {
                return { frame, element: isXPath ? element[0] : element };
            }
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
        
        await new Promise(r => setTimeout(r, 5000)); // รอโหลดหน้า

        const reportClient = await page.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // 5. Crystal Report Export (แก้ใหม่: ใช้ Partial ID Matcher)
        // ---------------------------------------------------------
        console.log('💾 Handling Crystal Report Export...');

        // 5.1 หาปุ่ม Export Icon (ไอคอนรูปเครื่องพิมพ์/ส่งออก)
        // ใช้ Selector จับ ID ที่ลงท้ายด้วย _toptoolbar_export (ตัดส่วนภาษาไทยทิ้ง)
        const exportIconSelector = '[id$="_toptoolbar_export"]'; 
        let found = null;
        
        console.log('   Searching for Export button...');
        for (let i = 0; i < 30; i++) {
            found = await findElementInFrames(page, exportIconSelector);
            if (found) break;
            await new Promise(r => setTimeout(r, 1000));
        }

        if (!found) throw new Error("Could not find Export button (timeout)!");
        
        console.log('   Clicking Export Icon...');
        await found.element.click();
        const activeFrame = found.frame; // จำ Frame ที่เจอไว้ใช้ต่อ

        // 5.2 รอ Popup และหาปุ่มเลือก Format
        console.log('   Waiting for Dialog...');
        await new Promise(r => setTimeout(r, 2000));

        // พยายามหาปุ่ม Dropdown Arrow (Combo box)
        // Selector: จับ ID ที่ลงท้ายด้วย _dialog_combo
        const comboSelector = '[id$="_dialog_combo"]'; 
        
        // ลองคลิกที่ Combo box เพื่อเปิดเมนู (สำคัญ! ถ้าไม่คลิก ตัวเลือกอาจไม่โผล่)
        try {
            const comboBtn = await activeFrame.$(comboSelector);
            if (comboBtn) {
                console.log('   Clicking Dropdown Arrow...');
                await comboBtn.click();
                await new Promise(r => setTimeout(r, 1000));
            }
        } catch(e) { console.log('   (Combo button not found, trying direct search...)'); }

        // 5.3 เลือก "Microsoft Excel Workbook Data-only"
        // ใช้ XPath หา text โดยตรง (แม่นยำที่สุดสำหรับ Custom Menu)
        console.log('   Selecting Excel Data-only...');
        const excelOptionXPath = "//*[contains(text(), 'Microsoft Excel Workbook Data-only')]";
        const excelOption = await activeFrame.$x(excelOptionXPath);

        if (excelOption.length > 0) {
            await excelOption[0].click();
        } else {
            console.warn('⚠️ Text option not found, trying Fallback (Keyboard)...');
            // Fallback: กดลูกศรลงเรื่อยๆ
            await page.keyboard.press('ArrowDown');
            await page.keyboard.press('ArrowDown');
            await page.keyboard.press('Enter');
        }
        
        await new Promise(r => setTimeout(r, 1000));

        // 5.4 กดปุ่ม Export (ปุ่มยืนยันสุดท้าย)
        // Selector: จับ ID ที่ลงท้ายด้วย _dialog_submitBtn
        const submitBtnSelector = '[id$="_dialog_submitBtn"]';
        const submitBtn = await activeFrame.$(submitBtnSelector);
        
        if (submitBtn) {
            console.log('   Clicking Final Export Button...');
            await submitBtn.click();
        } else {
             throw new Error("Could not find Final Submit button!");
        }

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
            if (page && !page.isClosed()) {
                await page.screenshot({ path: 'error_screenshot.png', fullPage: true });
            }
        } catch (e) {}

        if (browser) await browser.close();
        process.exit(1);
    }
})();
