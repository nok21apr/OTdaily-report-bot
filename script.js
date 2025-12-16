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

    // เปิด Browser (แก้ headless: false ถ้าอยากดูในเครื่องตัวเอง)
    const browser = await puppeteer.launch({
        headless: "new", 
        protocolTimeout: 300000, // เพิ่มความอึดเป็น 5 นาที
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    let page = await browser.newPage();

    // ฟังก์ชั่นช่วยถ่ายรูป (ตั้งชื่อตามลำดับเลข)
    const takeSnap = async (p, name) => {
        try {
            console.log(`📸 Snap: ${name}`);
            await p.screenshot({ path: name, fullPage: true });
        } catch(e) { console.log('⚠️ Snap failed: ' + name); }
    };

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
        
        await takeSnap(page, '01_login_page.png'); // 📸 1

        await page.waitForSelector('#txtUsername', { visible: true });
        await page.type('#txtUsername', WEB_CONFIG.user.trim(), { delay: 50 });
        await page.type('#txtPassword', WEB_CONFIG.pass.trim(), { delay: 50 });
        
        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 60000 }),
            page.keyboard.press('Enter')
        ]);
        console.log('✅ Login Success');
        await takeSnap(page, '02_logged_in.png'); // 📸 2

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
        await takeSnap(page, '03_arrived_report.png'); // 📸 3

        // ---------------------------------------------------------
        // 3. Fill Form & Date Logic (Fixed)
        // ---------------------------------------------------------
        console.log('📝 Filling form...');
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_ddlDoctype', { visible: true });
        await page.select('#ctl00_ContentPlaceHolder1_ddlDoctype', '1');
        await new Promise(r => setTimeout(r, 2000));

        const otTypeSelector = '#ctl00_ContentPlaceHolder1_ddlOt';
        if (await page.$(otTypeSelector) !== null) {
            await page.select(otTypeSelector, '14');
            await new Promise(r => setTimeout(r, 2000));
        }

        // วันที่
        const now = new Date();
        const thaiYear = now.getFullYear() + 543;
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const firstDayValue = `01/${month}/${thaiYear}`;
        
        console.log(`📅 Setting date to: ${firstDayValue}`);
        const dateInputSelector = '#ctl00_ContentPlaceHolder1_txtFromDate';
        await page.waitForSelector(dateInputSelector);
        
        // ลบค่าเก่าแบบชัวร์ๆ
        await page.click(dateInputSelector);
        await new Promise(r => setTimeout(r, 500));
        await page.keyboard.down('Control'); await page.keyboard.press('A'); await page.keyboard.up('Control');
        await page.keyboard.press('Backspace');
        
        // พิมพ์ค่าใหม่
        await page.type(dateInputSelector, firstDayValue, { delay: 100 });
        await page.keyboard.press('Tab'); // ปิดปฏิทิน
        await page.click('body'); // ย้ำปิดปฏิทิน
        
        await takeSnap(page, '04_form_filled.png'); // 📸 4 เช็ควันที่ตรงนี้

        // ---------------------------------------------------------
        // 4. Generate Report & SWITCH TAB
        // ---------------------------------------------------------
        console.log('⏳ Generating Report...');
        
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');
        
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        console.log('✅ Switched to New Report Tab');
        await reportPage.bringToFront();
        await reportPage.setViewport({ width: 1366, height: 768 });
        
        console.log('   Waiting 20s for Crystal Report Iframe...');
        await new Promise(r => setTimeout(r, 20000));
        await takeSnap(reportPage, '05_report_loaded.png'); // 📸 5 เช็คหน้า Report

        const reportClient = await reportPage.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // 5. Crystal Report Export (Slow Sequence)
        // ---------------------------------------------------------
        console.log('💾 Handling Crystal Report Export...');
        try { await reportPage.click('body'); } catch(e) {}
        await new Promise(r => setTimeout(r, 2000));

        // --- 1. Open Dialog ---
        await reportPage.keyboard.press('Tab'); await new Promise(r => setTimeout(r, 800));
        await reportPage.keyboard.press('Tab'); await new Promise(r => setTimeout(r, 800));
        await reportPage.keyboard.press('Enter');
        console.log('   Waiting 10s for Dialog...');
        await new Promise(r => setTimeout(r, 10000));
        await takeSnap(reportPage, '06_dialog_opened.png'); // 📸 6 เช็ค Dialog

        // --- 2. Menu ---
        await reportPage.keyboard.press('Tab'); await new Promise(r => setTimeout(r, 800));
        await reportPage.keyboard.press('Enter');
        await new Promise(r => setTimeout(r, 5000));
        await takeSnap(reportPage, '07_menu_list.png'); // 📸 7 เช็คเมนู

        // --- 3. Select Excel ---
        for (let i = 0; i < 4; i++) {
            await reportPage.keyboard.press('Tab');
            await new Promise(r => setTimeout(r, 600));
        }
        await reportPage.keyboard.press('Enter');
        console.log('   Waiting 5s for format selection...');
        await new Promise(r => setTimeout(r, 5000));
        await takeSnap(reportPage, '08_format_selected.png'); // 📸 8 เช็คว่าเลือก Excel ถูกไหม

        // --- 4. Final Export ---
        console.log('   Ready for Final Export...');
        // Tab ครั้งที่ 1
        await reportPage.keyboard.press('Tab'); await new Promise(r => setTimeout(r, 1500));
        // Tab ครั้งที่ 2 (ควรจะถึงปุ่ม Export)
        await reportPage.keyboard.press('Tab'); await new Promise(r => setTimeout(r, 2000));
        
        await takeSnap(reportPage, '09_final_focus.png'); // 📸 9 สำคัญมาก! ดูว่าโฟกัสอยู่ที่ปุ่มไหน

        await reportPage.keyboard.press('Enter');

        // ---------------------------------------------------------
        // 6. Wait for Download
        // ---------------------------------------------------------
        console.log('⬇️ Waiting for file...');
        let downloadedFile;
        for (let i = 0; i < 300; i++) { 
            await new Promise(r => setTimeout(r, 1000));
            if (fs.existsSync(downloadPath)) {
                const files = fs.readdirSync(downloadPath);
                // Log ดูไฟล์ทุกไฟล์ที่เจอ
                if (i % 10 === 0) console.log(`   Scanned files: ${JSON.stringify(files)}`); 
                
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
        // ถ่ายรูปสุดท้ายตอน Error เผื่อไว้
        if (page && !page.isClosed()) await page.screenshot({ path: '99_error_final.png', fullPage: true });
        
        if (browser) await browser.close();
        process.exit(1);
    }
})();

    
