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

    // [CONFIG] เพิ่ม Timeout ให้นานขึ้น
    const browser = await puppeteer.launch({
        headless: "new",
        protocolTimeout: 300000, 
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    let page = await browser.newPage();

    // ฟังก์ชั่นช่วยถ่ายรูป Debug
    const takeSnap = async (name) => {
        try {
            if (page && !page.isClosed()) await page.screenshot({ path: name, fullPage: true });
        } catch(e) {}
    };

    try {
        await page.emulateTimezone('Asia/Bangkok');
        // ปรับ Viewport ตาม Code ที่ท่านบันทึกมา (928x791)
        await page.setViewport({ width: 928, height: 791 }); 

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
        // 2. Navigation (Original Method)
        // ---------------------------------------------------------
        console.log('📂 Navigating to Report (Original Method)...');
        
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
        await takeSnap('03_arrived_report.png');

        // ---------------------------------------------------------
        // 3. Fill Form (UPDATED: Step 3.1 & 3.2 with New Offsets)
        // ---------------------------------------------------------
        console.log('📝 Filling form using updated logic...');
        
        // --- 3.1 Select Doctype = 1 ---
        const ddlDoctype = '#ctl00_ContentPlaceHolder1_ddlDoctype';
        await page.waitForSelector(ddlDoctype);
        // เลียนแบบการคลิกที่พิกัด x: 526, y: 16.49 (ตามโค้ดใหม่)
        await page.click(ddlDoctype, { offset: { x: 526, y: 16.49 } });
        await new Promise(r => setTimeout(r, 500));
        await page.select(ddlDoctype, '1');

        // --- 3.2 Select OT = 14 ---
        const ddlOt = '#ctl00_ContentPlaceHolder1_ddlOt';
        if (await page.$(ddlOt)) {
            // เลียนแบบการคลิกที่พิกัด x: 396, y: 15.52 (ตามโค้ดใหม่)
            await page.click(ddlOt, { offset: { x: 396, y: 15.52 } });
            await new Promise(r => setTimeout(r, 500));
            await page.select(ddlOt, '14');
        }

        // --- 3.3 FROM Date Input Sequence ---
        console.log('   Handling "From Date" Input...');
        const fromDateSelector = '#ctl00_ContentPlaceHolder1_txtFromDate';
        await page.waitForSelector(fromDateSelector);

        // Click Offset x: 128, y: 18
        await page.click(fromDateSelector, { offset: { x: 128, y: 18 } });
        await new Promise(r => setTimeout(r, 500));

        // Backspace loop (15 times)
        console.log('      Clearing From Date (Backspace x15)...');
        for(let i=0; i<15; i++) {
            await page.keyboard.press('Backspace');
        }
        await new Promise(r => setTimeout(r, 300));

        // Type '01'
        console.log("      Typing '01'...");
        await page.keyboard.type('01', { delay: 100 });(พิมพ์ข้อความ 01)
        await new Promise(r => setTimeout(r, 300));

        // Press Tab
        console.log("      Pressing Tab...");
        await page.keyboard.press('Tab');
        await new Promise(r => setTimeout(r, 1000));


        // --- 3.4 TO Date Input Sequence ---
        console.log('   Handling "To Date" Input...');
        const toDateSelector = '#ctl00_ContentPlaceHolder1_txtToDate';
        // Click Offset x: 185, y: 14
        await page.click(toDateSelector, { offset: { x: 185, y: 14 } });
        await new Promise(r => setTimeout(r, 500));

        // Backspace loop (4 times)
        console.log('      Clearing To Date (Backspace x4)...');
        for(let i=0; i<4; i++) {
            await page.keyboard.press('Backspace');
        }
        await new Promise(r => setTimeout(r, 300));

        // Press Tab
        console.log("      Pressing Tab...");
        await page.keyboard.press('Tab');
        await new Promise(r => setTimeout(r, 1000));
        
        await takeSnap('04_form_filled_updated.png');

        // ---------------------------------------------------------
        // 4. Generate Report (Click Show Report)
        // ---------------------------------------------------------
        console.log('⏳ Generating Report...');
        
        // เตรียมจับ Tab ใหม่
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        
        // คลิกปุ่มแสดงรายงาน (Offset x: 65, y: 16)
        const showReportBtn = '#ctl00_ContentPlaceHolder1_lnkShowReport';
        await page.waitForSelector(showReportBtn);
        await page.click(showReportBtn, { offset: { x: 65, y: 16 } });
        
        console.log('   Click command sent. Waiting for tab...');
        
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        console.log('✅ Switched to New Report Tab');
        await reportPage.bringToFront();
        await reportPage.setViewport({ width: 928, height: 791 }); 
        
        // รอโหลด Crystal Report
        console.log('   Waiting 20s for Crystal Report Iframe...');
        await new Promise(r => setTimeout(r, 20000));
        await takeSnap('05_report_loaded.png'); 

        const reportClient = await reportPage.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // 5. Crystal Report Export (New Recorder Sequence)
        // ---------------------------------------------------------
        console.log('💾 Handling Crystal Report Export (New Recorder Sequence)...');

        // 1. คลิกไอคอน Export (รูปเครื่องพิมพ์/ส่งออก)
        console.log('   1. Clicking Export Icon...');
        const exportIconSelector = '[id$="_toptoolbar_export"]'; // Selector แบบยืดหยุ่น (Ends with)
        try {
            await reportPage.waitForSelector(exportIconSelector, { timeout: 10000 });
            await reportPage.click(exportIconSelector, { offset: { x: 5, y: 7 } });
        } catch (e) {
            console.log('      ⚠️ Standard selector failed, trying aria-label...');
            await reportPage.click('#IconImg_รายงานรายละเอียดขออนุมัติใบโอทีของพนักงาน_toptoolbar_export');
        }
        await new Promise(r => setTimeout(r, 2000)); // รอ Dialog เด้ง

        // 2. Keyboard Navigation (Tab -> Enter -> Tab x3 -> Enter)
        console.log('   2. Navigating Export Options (Tab/Enter Sequence)...');
        
        // Tab -> Enter (เปิด List Format?)
        await reportPage.keyboard.press('Tab'); await new Promise(r => setTimeout(r, 500));
        await reportPage.keyboard.press('Enter'); await new Promise(r => setTimeout(r, 2000));

        // Tab x3 (เลือก Excel Data Only?)
        for(let i=0; i<3; i++) {
            await reportPage.keyboard.press('Tab');
            await new Promise(r => setTimeout(r, 500));
        }

        // Enter (ยืนยัน Format)
        await reportPage.keyboard.press('Enter');
        await new Promise(r => setTimeout(r, 2000));

        // 3. คลิกปุ่ม Export Final Submit
        console.log('   3. Clicking Final Export Button...');
        const submitBtnSelector = '[id$="_dialog_submitBtn"]'; 
        
        try {
            await reportPage.waitForSelector(submitBtnSelector, { timeout: 10000 });
            await reportPage.click(submitBtnSelector, { offset: { x: 12, y: 1 } });
        } catch (e) {
             console.log('      ⚠️ Submit button selector failed, trying fallback...');
             await reportPage.keyboard.press('Enter');
        }
        
        console.log('      Export sequence completed. Waiting for download...');

        // ---------------------------------------------------------
        // 6. Wait for Download
        // ---------------------------------------------------------
        console.log('⬇️ Waiting for file...');
        let downloadedFile;
        // รอสูงสุด 5 นาที
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
            if (page && !page.isClosed()) await page.screenshot({ path: '99_error_final.png', fullPage: true });
        } catch(e){}
        if (browser) await browser.close();
        process.exit(1);
    }
})();
