require('dotenv').config();
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');

// รับค่าจาก .env หรือ GitHub Secrets
const WEB_USER = process.env.WEB_USER;
const WEB_PASS = process.env.WEB_PASS;
const EMAIL_USER = process.env.GMAIL_USER;
const EMAIL_PASS = process.env.GMAIL_PASS;
const EMAIL_TO = process.env.EMAIL_TO;

(async () => {
    console.log('🚀 Starting Bot (Business Plus Crystal Report)...');

    // ตรวจสอบตัวแปรสำคัญ
    if (!WEB_USER || !WEB_PASS || !EMAIL_USER || !EMAIL_PASS) {
        console.error('❌ Error: Secrets incomplete (Check .env or GitHub Secrets).');
        process.exit(1);
    }

    const downloadPath = path.join(__dirname, 'downloads');
    // เคลียร์โฟลเดอร์ดาวน์โหลดก่อนเริ่ม
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    let browser = null;
    let page = null;

    try {
        console.log('🖥️ Launching Browser...');
        browser = await puppeteer.launch({
            headless: "new", // รันแบบไม่มีหน้าจอ
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--window-size=1920,1080',
                '--lang=th-TH,th'
            ]
        });

        page = await browser.newPage();
        
        // Timeout นานหน่อยเผื่อเว็บช้า
        page.setDefaultNavigationTimeout(300000);
        page.setDefaultTimeout(300000);

        await page.emulateTimezone('Asia/Bangkok');
        const client = await page.target().createCDPSession();
        await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });

        // ---------------------------------------------------------
        // Step 1: Login
        // ---------------------------------------------------------
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://leave.ttkasia.co.th/Login/Login.aspx', { waitUntil: 'networkidle2' });
        
        // รอช่อง User โผล่
        await page.waitForSelector('#txtUsername', { visible: true });
        
        // พิมพ์ User/Pass (ใช้ trim กันเหนียว)
        await page.type('#txtUsername', WEB_USER.trim());
        await page.type('#txtPassword', WEB_PASS.trim());
        
        console.log('   Clicking Login...');
        await Promise.all([
            page.waitForNavigation(),
            page.keyboard.press('Enter')
        ]);
        console.log('✅ Login Success');

        // ---------------------------------------------------------
        // Step 2: Navigate to Report
        // ---------------------------------------------------------
        console.log('2️⃣ Step 2: Go to Report Page...');
        // คลิกเมนูหลัก
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_imgLeave', { visible: true });
        await page.click('#ctl00_ContentPlaceHolder1_imgLeave');
        
        // คลิกเมนูรายงาน
        await page.waitForSelector('#ctl00_Report_Menu > a');
        await page.click('#ctl00_Report_Menu > a');
        
        // คลิกรายงานย่อย
        const subMenuSelector = '#ctl00_Report_Menu > ul a';
        await page.waitForSelector(subMenuSelector, { visible: true });
        await Promise.all([
            page.waitForNavigation(),
            page.click(subMenuSelector)
        ]);

        // ---------------------------------------------------------
        // Step 3: Fill Form & Date Logic
        // ---------------------------------------------------------
        console.log('3️⃣ Step 3: Fill Form...');
        
        // เลือกประเภทเอกสาร และ OT
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_ddlDoctype');
        await page.select('#ctl00_ContentPlaceHolder1_ddlDoctype', '1');
        await page.select('#ctl00_ContentPlaceHolder1_ddlOt', '14');

        // คำนวณวันที่ 1 ของเดือนปัจจุบัน (พ.ศ.)
        const now = new Date();
        const thaiYear = now.getFullYear() + 543;
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const firstDayValue = `01/${month}/${thaiYear}`;
        
        console.log(`   Setting date to: ${firstDayValue}`);
        // คลิก 3 ครั้งเพื่อเลือกข้อความเก่าทั้งหมดแล้วพิมพ์ทับ
        await page.click('#ctl00_ContentPlaceHolder1_txtFromDate', { clickCount: 3 });
        await page.type('#ctl00_ContentPlaceHolder1_txtFromDate', firstDayValue);

        // ---------------------------------------------------------
        // Step 4: Generate Report & Handle New Tab
        // ---------------------------------------------------------
        console.log('4️⃣ Step 4: Generating Report (Waiting for New Tab)...');
        
        // เตรียมจับ Event หน้าต่างใหม่ (สำคัญมากสำหรับ Crystal Report)
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        
        // กดปุ่มแสดงรายงาน
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');
        
        // รอรับหน้าต่างใหม่ที่เด้งขึ้นมา
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        // สลับไปคุมหน้าใหม่
        page = reportPage; 
        await page.bringToFront();
        await page.setViewport({ width: 1920, height: 1080 });
        
        // Setup Download behavior ให้หน้าใหม่ด้วย (ต้องทำซ้ำสำหรับ Tab ใหม่)
        const reportClient = await page.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // Step 5: Export Handling
        // ---------------------------------------------------------
        console.log('5️⃣ Step 5: Handling Crystal Report Export...');
        
        // รอจนปุ่ม Export โผล่ (Timeout 60 วิ เผื่อโหลดนาน)
        const exportBtnSelector = 'a[title="Export this report"], img[alt="Export this report"]';
        await page.waitForSelector(exportBtnSelector, { visible: true, timeout: 60000 });
        
        console.log('   Clicking Export Icon...');
        await page.click(exportBtnSelector);

        // รอ Dropdown เลือก Format
        await page.waitForSelector('select', { visible: true });
        
        // เลือก Microsoft Excel Workbook Data-only
        // (ใช้ Logic หาจาก Text เพราะ ID เปลี่ยนได้)
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

        // กดปุ่ม Export สุดท้าย (ใช้ Selector ที่ลงท้ายด้วย _dialog_submitBtn)
        const finalSubmitSelector = 'a[id$="_dialog_submitBtn"]';
        await page.waitForSelector(finalSubmitSelector);
        await page.click(finalSubmitSelector);

        // ---------------------------------------------------------
        // Step 6: Wait for Download
        // ---------------------------------------------------------
        console.log('6️⃣ Step 6: Waiting for file...');
        let downloadedFile = null;
        
        // วนลูปรอเหมือนโค้ด DTC (รอสูงสุด 300 วินาที)
        for (let i = 0; i < 300; i++) {
            await new Promise(r => setTimeout(r, 1000));
            const files = fs.readdirSync(downloadPath);
            // หาไฟล์ .xls/.xlsx ที่ไม่ใช่ .crdownload
            const target = files.find(f => (f.endsWith('.xlsx') || f.endsWith('.xls')) && !f.endsWith('.crdownload'));
            
            if (target) {
                downloadedFile = target;
                break;
            }
            if (i > 0 && i % 10 === 0) console.log(`   ...still waiting (${i}s)`);
        }

        if (!downloadedFile) throw new Error('❌ Download Timeout: File never arrived.');
        const originalFilePath = path.join(downloadPath, downloadedFile);
        console.log(`✅ File Downloaded: ${originalFilePath}`);

        // ปิด Browser ได้เลยเมื่อโหลดเสร็จ
        await browser.close();
        browser = null; // reset variable

        // ---------------------------------------------------------
        // Step 7: Convert to CSV UTF-8
        // ---------------------------------------------------------
        console.log('🔄 Step 7: Converting to CSV UTF-8...');
        
        // อ่านไฟล์ Excel
        const workbook = xlsx.readFile(originalFilePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const csvContent = xlsx.utils.sheet_to_csv(worksheet);
        
        // ตั้งชื่อไฟล์ใหม่ (.csv)
        const csvFileName = downloadedFile.replace(/\.[^/.]+$/, "") + ".csv";
        const csvFilePath = path.join(downloadPath, csvFileName);

        // เขียนไฟล์ CSV โดยเติม \uFEFF (BOM) ข้างหน้า เพื่อให้อ่านภาษาไทยออก
        fs.writeFileSync(csvFilePath, '\uFEFF' + csvContent, { encoding: 'utf8' });
        console.log(`✅ Converted to: ${csvFilePath}`);

        // ลบไฟล์ Excel ต้นฉบับทิ้ง
        fs.unlinkSync(originalFilePath);
        console.log('🗑️ Deleted original Excel file');

        // ---------------------------------------------------------
        // Step 8: Email
        // ---------------------------------------------------------
        console.log('📧 Step 8: Sending Email...');
        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: { user: EMAIL_USER, pass: EMAIL_PASS }
        });

        await transporter.sendMail({
            from: `"Auto Report Bot" <${EMAIL_USER}>`,
            to: EMAIL_TO,
            subject: `รายงาน Business Plus - ${new Date().toLocaleDateString()}`,
            text: `ดาวน์โหลดและแปลงไฟล์สำเร็จ\nไฟล์: ${csvFileName}`,
            attachments: [{ filename: csvFileName, path: csvFilePath }]
        });
        
        console.log('🎉 Mission Complete!');

    } catch (error) {
        console.error('❌ FATAL ERROR:', error);
        
        // ถ่ายรูปตอน Error เก็บไว้ดู
        if (page && !page.isClosed()) {
            try { 
                await page.screenshot({ 
                    path: path.join(downloadPath, 'fatal_error.png'),
                    fullPage: true 
                });
                console.log('📸 Screenshot saved: fatal_error.png');
            } catch(e){}
        }
        
        if (browser) await browser.close();
        process.exit(1);
    }
})();
