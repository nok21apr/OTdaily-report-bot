const puppeteer = require('puppeteer');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// --- ตั้งค่า Email (ดึงจาก GitHub Secrets) ---
const EMAIL_CONFIG = {
    user: process.env.GMAIL_USER,
    pass: process.env.GMAIL_PASS,
    to:   'recipient_email@example.com', // 🔴 อย่าลืมแก้: อีเมลปลายทาง
    subject: 'Daily Overtime Report (CSV)',
    text: 'Attached is the requested report in CSV UTF-8 format.'
};

(async () => {
    // กำหนดโฟลเดอร์สำหรับดาวน์โหลดไฟล์
    const downloadPath = path.resolve(__dirname, 'downloads');
    
    // เคลียร์โฟลเดอร์เก่าทิ้ง (ถ้ามี) เพื่อไม่ให้ไฟล์ปนกัน
    if (fs.existsSync(downloadPath)) {
        fs.rmSync(downloadPath, { recursive: true, force: true });
    }
    fs.mkdirSync(downloadPath);

    // 1. Setup Browser
    const browser = await puppeteer.launch({
        headless: "new", // โหมด Headless สำหรับ GitHub Actions
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    // ประกาศตัวแปร page ไว้ข้างนอก try เพื่อให้เรียกใช้ใน catch ได้ (ตอนถ่ายรูป Error)
    let page = await browser.newPage();

    try {
        // ตั้งค่า Timezone และขนาดหน้าจอ
        await page.emulateTimezone('Asia/Bangkok');
        await page.setViewport({ width: 1280, height: 800 });

        // ตั้งค่าให้ Download ลงโฟลเดอร์ที่เตรียมไว้
        const client = await page.target().createCDPSession();
        await client.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        console.log('🚀 Starting process...');

        // ---------------------------------------------------------
        // 2. Login Process
        // ---------------------------------------------------------
        console.log('🔑 Logging in...');
        await page.goto('https://leave.ttkasia.co.th/Login/Login.aspx', { waitUntil: 'networkidle0' });
        await page.type('#txtUsername', '200068');
        await page.type('#txtPassword', 'QIMLhLwh');
        
        await Promise.all([
            page.waitForNavigation(),
            page.keyboard.press('Enter')
        ]);

        // ---------------------------------------------------------
        // 3. Navigation to Report
        // ---------------------------------------------------------
        console.log('📂 Navigating to Report...');
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_imgLeave');
        await page.click('#ctl00_ContentPlaceHolder1_imgLeave');
        
        // รอเมนูโหลดแล้วคลิก
        await page.waitForSelector('#ctl00_Report_Menu > a');
        await page.click('#ctl00_Report_Menu > a');
        
        // คลิกเมนูย่อย
        const subMenuSelector = '#ctl00_Report_Menu > ul a';
        await page.waitForSelector(subMenuSelector);
        await Promise.all([
            page.waitForNavigation(),
            page.click(subMenuSelector)
        ]);

        // ---------------------------------------------------------
        // 4. Fill Form & Date Logic
        // ---------------------------------------------------------
        console.log('📝 Filling form...');
        await page.select('#ctl00_ContentPlaceHolder1_ddlDoctype', '1');
        await page.select('#ctl00_ContentPlaceHolder1_ddlOt', '14');

        // คำนวณวันที่ 1 ของเดือนปัจจุบัน (พ.ศ.)
        const now = new Date();
        const thaiYear = now.getFullYear() + 543;
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const firstDayValue = `01/${month}/${thaiYear}`;
        
        console.log(`📅 Setting date: ${firstDayValue}`);
        // คลิก 3 ทีเพื่อเลือกข้อความเก่าทั้งหมด แล้วพิมพ์ทับ
        await page.click('#ctl00_ContentPlaceHolder1_txtFromDate', { clickCount: 3 });
        await page.type('#ctl00_ContentPlaceHolder1_txtFromDate', firstDayValue);

        // ---------------------------------------------------------
        // 5. Generate Report & Handle New Tab
        // ---------------------------------------------------------
        console.log('⏳ Generating Report...');
        
        // เตรียมจับ Event หน้าต่างใหม่ (สำคัญมากสำหรับ Crystal Report)
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');
        
        // รอรับหน้าต่างใหม่ที่เด้งขึ้นมา
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        // สลับไปคุมหน้าใหม่ (อัปเดตตัวแปร page ให้เป็นหน้า report เพื่อใช้ถ่ายรูปถ้า error)
        page = reportPage; 
        await page.bringToFront();
        
        // Setup Download behavior ให้หน้าใหม่ด้วย (ต้องทำซ้ำสำหรับ Tab ใหม่)
        const reportClient = await page.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // รอจนปุ่ม Export โผล่ (Timeout 60 วิ เผื่อโหลดนาน)
        const exportBtnSelector = 'a[title="Export this report"], img[alt="Export this report"]';
        await page.waitForSelector(exportBtnSelector, { timeout: 60000 });
        
        console.log('💾 Clicking Export...');
        await page.click(exportBtnSelector);

        // รอ Dropdown เลือก Format
        await page.waitForSelector('select');
        
        // ค้นหา Option ที่เป็น Excel Data-only โดยการอ่าน Text ใน Dropdown
        const selectId = await page.evaluate(() => {
            const options = Array.from(document.querySelectorAll('option'));
            const target = options.find(o => o.text.includes('Microsoft Excel Workbook Data-only'));
            return target ? target.parentElement.id : null;
        });
        
        if (selectId) {
            await page.select(`#${selectId}`, 'Microsoft Excel Workbook Data-only');
        } else {
            // Fallback: ถ้าหา ID ไม่เจอ ให้กดลูกศรลง (ไม่แนะนำแต่มักใช้แก้ขัดได้)
            await page.keyboard.press('ArrowDown');
        }

        // กดปุ่มยืนยัน Export (Step สุดท้าย)
        // ใช้ Selector ที่ลงท้ายด้วย _dialog_submitBtn เพื่อหนี Dynamic ID
        const finalSubmitSelector = 'a[id$="_dialog_submitBtn"]';
        await page.waitForSelector(finalSubmitSelector);
        await page.click(finalSubmitSelector);

        // ---------------------------------------------------------
        // 6. Wait for Download
        // ---------------------------------------------------------
        console.log('⬇️ Waiting for file...');
        let downloadedFile;
        // วนลูปรอจนกว่าไฟล์จะมา (สูงสุด 60 วิ) และต้องไม่ใช่ .crdownload
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
        
        // ปิด Browser ได้เลยเมื่อโหลดเสร็จ
        await browser.close(); 

        // ---------------------------------------------------------
        // 7. Convert Excel to CSV UTF-8 & Cleanup
        // ---------------------------------------------------------
        console.log('🔄 Converting to CSV UTF-8...');
        
        // อ่านไฟล์ Excel
        const workbook = xlsx.readFile(originalFilePath);
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // แปลงเป็น CSV Text
        const csvContent = xlsx.utils.sheet_to_csv(worksheet);
        
        // ตั้งชื่อไฟล์ใหม่ (.csv)
        const csvFileName = downloadedFile.replace(/\.[^/.]+$/, "") + ".csv";
        const csvFilePath = path.join(downloadPath, csvFileName);

        // เขียนไฟล์ CSV โดยเติม \uFEFF (BOM) ข้างหน้า เพื่อให้อ่านไทยออก
        fs.writeFileSync(csvFilePath, '\uFEFF' + csvContent, { encoding: 'utf8' });
        console.log(`✅ Converted to: ${csvFilePath}`);

        // ลบไฟล์ Excel ต้นฉบับทิ้ง
        fs.unlinkSync(originalFilePath);
        console.log('🗑️ Deleted original Excel file');

        // ---------------------------------------------------------
        // 8. Send Email
        // ---------------------------------------------------------
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
        
        // ลบไฟล์ CSV หลังส่งเสร็จ (Optional)
        // fs.unlinkSync(csvFilePath);

    } catch (error) {
        console.error('❌ Error occurred:', error);

        // ---------------------------------------------------------
        // 📸 Debug Screenshot (ทำงานเมื่อ Error เท่านั้น)
        // ---------------------------------------------------------
        try {
            if (page) {
                await page.screenshot({ 
                    path: 'error_screenshot.png', 
                    fullPage: true 
                });
                console.log('📸 Debug screenshot saved: error_screenshot.png');
            }
        } catch (snapshotError) {
            console.error('Could not take screenshot:', snapshotError);
        }

        // ปิด Browser ในกรณี Error
        await browser.close();
        
        // ส่ง exit code 1 เพื่อแจ้ง GitHub ว่า Job Failed
        process.exit(1);
    }
})();
