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

        console.log('✅ Arrived at Report Page.');

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
        // 4. Generate Report & SWITCH TAB
        // ---------------------------------------------------------
        console.log('⏳ Generating Report...');
        
        // เตรียมจับ Tab ใหม่
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');
        
        // รอรับ Tab ใหม่
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        console.log('✅ Switched to New Report Tab');
        await reportPage.bringToFront();
        await reportPage.setViewport({ width: 1366, height: 768 });
        
        // รอโหลด (Hard Wait) 20 วินาที เพื่อให้แน่ใจว่า Crystal Report พร้อม
        console.log('   Waiting 20s for Crystal Report Iframe...');
        await new Promise(r => setTimeout(r, 20000));

        // ตั้งค่า Download
        const reportClient = await reportPage.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // 5. Crystal Report Export (Strict Keyboard Sequence) - ADJUSTED
        // ---------------------------------------------------------
        console.log('💾 Handling Crystal Report Export via Keyboard...');
        
        // คลิกที่ว่างๆ 1 ที เพื่อ Focus หน้าเว็บ
        try { await reportPage.click('body'); } catch(e) {}
        await new Promise(r => setTimeout(r, 2000)); // รอ 2 วินาทีให้ Focus นิ่ง

        // --- Sequence 1: เปิด Dialog ---
        console.log('   1. Opening Dialog (Tab x2 -> Enter)...');
        await reportPage.keyboard.press('Tab'); await new Promise(r => setTimeout(r, 1000)); // รอ 1 วิ
        await reportPage.keyboard.press('Tab'); await new Promise(r => setTimeout(r, 1000)); // รอ 1 วิ
        await reportPage.keyboard.press('Enter');
        
        // [จุดสำคัญ] หน้าต่าง Dialog มักจะโหลดนาน ให้รอ 10 วินาที เพื่อความชัวร์
        console.log('      (Waiting 10s for Dialog to fully appear...)');
        await new Promise(r => setTimeout(r, 10000)); 

        // --- Sequence 2: เข้าเมนูเลือกไฟล์ ---
        console.log('   2. Entering Format Menu (Tab x1 -> Enter)...');
        await reportPage.keyboard.press('Tab'); await new Promise(r => setTimeout(r, 1000));
        await reportPage.keyboard.press('Enter');
        
        // รอเมนู Dropdown ไหลลงมา
        console.log('      (Waiting 5s for Menu options...)');
        await new Promise(r => setTimeout(r, 5000));

        // --- Sequence 3: เลือก Excel Data-only ---
        console.log('   3. Selecting "Excel Data-only" (Tab x4 -> Enter)...');
        for (let i = 0; i < 4; i++) {
            await reportPage.keyboard.press('Tab');
            await new Promise(r => setTimeout(r, 800)); // รอระหว่างกด Tab นานขึ้นนิดนึง
        }
        await reportPage.keyboard.press('Enter');
        
        // รอให้ระบบเลือกค่าเสร็จ (บางทีเลือกแล้วมันจะกระพริบโหลด)
        console.log('      (Waiting 5s for selection confirmation...)');
        await new Promise(r => setTimeout(r, 5000));

        // --- Sequence 4: กดปุ่ม Export สุดท้าย ---
        console.log('   4. Clicking Export Button (Tab x3 -> Enter)...');
        for (let i = 0; i < 3; i++) { // **ลองตรวจสอบว่าต้องกด Tab กี่ครั้งแน่ (2 หรือ 3)**
            await reportPage.keyboard.press('Tab');
            await new Promise(r => setTimeout(r, 800));
        }
        await reportPage.keyboard.press('Enter');

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
            if (page && !page.isClosed()) await page.screenshot({ path: 'error_main.png', fullPage: true });
            const pages = await browser.pages();
            if (pages.length > 1) await pages[pages.length-1].screenshot({ path: 'error_report.png', fullPage: true });
        } catch(e){}

        if (browser) await browser.close();
        process.exit(1);
    }
})();
