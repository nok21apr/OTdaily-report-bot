const puppeteer = require('puppeteer');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx'); // 🟢 เพิ่ม Library จัดการ Excel

// --- ตั้งค่า Email (ดึงจาก GitHub Secrets) ---
const EMAIL_CONFIG = {
    user: process.env.GMAIL_USER,        
    pass: process.env.GMAIL_PASS,        
    to:   'naruesit_jit@ttkasia.co.th',  // 🔴 อย่าลืมแก้: อีเมลปลายทาง
    subject: 'Daily Overtime Report (CSV)',
    text: 'Attached is the requested report in CSV UTF-8 format.'
};

(async () => {
    // กำหนดโฟลเดอร์สำหรับดาวน์โหลดไฟล์
    const downloadPath = path.resolve(__dirname, 'downloads');
    
    // เคลียร์โฟลเดอร์เก่าทิ้ง
    if (fs.existsSync(downloadPath)) {
        fs.rmSync(downloadPath, { recursive: true, force: true });
    }
    fs.mkdirSync(downloadPath);

    // 1. Setup Browser
    const browser = await puppeteer.launch({
        headless: "new",
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage();
    await page.emulateTimezone('Asia/Bangkok');
    await page.setViewport({ width: 1280, height: 800 });

    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', {
        behavior: 'allow',
        downloadPath: downloadPath
    });

    try {
        console.log('🚀 Starting process at 08:00 AM schedule...');

        // 2. Login
        console.log('🔑 Logging in...');
        await page.goto('https://leave.ttkasia.co.th/Login/Login.aspx', { waitUntil: 'networkidle0' });
        await page.type('#txtUsername', '200068');
        await page.type('#txtPassword', 'QIMLhLwh');
        await Promise.all([
            page.waitForNavigation(),
            page.keyboard.press('Enter')
        ]);

        // 3. Navigate to Report
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

        // 4. Fill Form & Date Logic
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

        // 5. Generate Report & Handle New Tab
        console.log('⏳ Generating Report...');
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');
        
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        await reportPage.bringToFront();
        
        // Setup Download on New Tab
        const reportClient = await reportPage.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // Click Export
        const exportBtnSelector = 'a[title="Export this report"], img[alt="Export this report"]';
        await reportPage.waitForSelector(exportBtnSelector, { timeout: 60000 });
        
        console.log('💾 Clicking Export...');
        await reportPage.click(exportBtnSelector);

        // Select Excel
        await reportPage.waitForSelector('select');
        const selectId = await reportPage.evaluate(() => {
            const options = Array.from(document.querySelectorAll('option'));
            const target = options.find(o => o.text.includes('Microsoft Excel Workbook Data-only'));
            return target ? target.parentElement.id : null;
        });
        
        if (selectId) {
            await reportPage.select(`#${selectId}`, 'Microsoft Excel Workbook Data-only');
        } else {
            await reportPage.keyboard.press('ArrowDown');
        }

        const finalSubmitSelector = 'a[id$="_dialog_submitBtn"]';
        await reportPage.waitForSelector(finalSubmitSelector);
        await reportPage.click(finalSubmitSelector);

        // 6. Wait for Download
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

        await browser.close(); // ปิด Browser ได้เลย

        // -------------------------------------------------------------
        // 🔄 7. Convert Excel to CSV UTF-8 & Cleanup
        // -------------------------------------------------------------
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

        // เขียนไฟล์ CSV โดยเติม \uFEFF (BOM) ข้างหน้า เพื่อให้ Excel อ่านไทยออก
        fs.writeFileSync(csvFilePath, '\uFEFF' + csvContent, { encoding: 'utf8' });
        console.log(`✅ Converted to: ${csvFilePath}`);

        // ลบไฟล์ Excel ต้นฉบับทิ้ง
        fs.unlinkSync(originalFilePath);
        console.log('🗑️ Deleted original Excel file');

        // -------------------------------------------------------------
        // 📧 8. Send Email (Send CSV file)
        // -------------------------------------------------------------
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
            attachments: [{ filename: csvFileName, path: csvFilePath }] // แนบไฟล์ CSV
        });

        console.log('✅ Email sent!');

        // 9. Cleanup Final File (Optional)
        // fs.unlinkSync(csvFilePath); 
        
    } catch (error) {
        console.error('❌ Error:', error);
        process.exit(1);
    }
})();
