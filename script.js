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

// 🛠️ ฟังก์ชันพิเศษ: พยายามหา Element ในทุก Frame โดยอัตโนมัติ (ฉลาดขึ้น)
async function findAndClickInFrames(page, selectors) {
    const startTime = Date.now();
    const timeout = 60000; // ให้เวลาหา 60 วินาที

    console.log('   🕵️‍♂️ Searching for button in all frames...');

    while (Date.now() - startTime < timeout) {
        // 1. ดึงรายชื่อ Frame ทั้งหมดใหม่ทุกรอบ (เผื่อมีการโหลด Iframe ใหม่)
        const frames = page.frames();
        
        // 2. วนลูปหาในทุก Frame
        for (const frame of frames) {
            for (const selector of selectors) {
                try {
                    // ลองหา Element
                    const element = await frame.$(selector);
                    if (element) {
                        // เช็คว่าปุ่มมองเห็นจริงไหม (Visible)
                        const isVisible = await element.evaluate(el => {
                            const style = window.getComputedStyle(el);
                            return style && style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
                        });

                        if (isVisible) {
                            console.log(`   ✅ Found button! (Frame: ${frame.name() || 'unnamed'}, Selector: ${selector})`);
                            
                            // เลื่อนเมาส์ไปชี้ก่อน แล้วค่อยคลิก (ช่วยเรื่องปุ่ม Hover)
                            await element.hover(); 
                            await new Promise(r => setTimeout(r, 500));
                            await element.click();
                            return frame; // ส่งคืน Frame ที่เจอ เพื่อใช้ต่อ
                        }
                    }
                } catch (e) {
                    // Ignore errors (frame detached etc.)
                }
            }
        }
        // ถ้าไม่เจอ ให้รอ 2 วินาทีแล้วหาใหม่
        await new Promise(r => setTimeout(r, 2000));
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
        await page.setViewport({ width: 1366, height: 768 }); // จอใหญ่หน่อย
        
        // รอแบบ Hard Wait นานๆ เลย เพื่อให้ Crystal Report โหลด Iframe ครบ
        console.log('   Waiting 10s for Crystal Report to load...');
        await new Promise(r => setTimeout(r, 10000));

        // Setup Download
        const reportClient = await page.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // 5. Crystal Report Export (The Hunt Begins!)
        // ---------------------------------------------------------
        console.log('💾 Handling Crystal Report Export...');

        // รายชื่อตัวจับปุ่ม Export (ลองหลายๆ แบบ)
        const exportSelectors = [
            '[id$="_toptoolbar_export"]',          // จับ ID ส่วนท้าย (มาตรฐาน)
            'a[title="Export this report"]',       // จับ Title ภาษาอังกฤษ
            'img[alt="Export this report"]',       // จับ Alt ภาพ
            'a[title*="Export"]',                  // จับ Title มีคำว่า Export
            '[id*="IconImg"][id*="export"]'        // จับ ID ที่มีคำว่า IconImg และ export
        ];

        // เรียกฟังก์ชันค้นหาและคลิก (จะได้ Frame ที่เจอ กลับมาใช้งานต่อ)
        const activeFrame = await findAndClickInFrames(page, exportSelectors);

        if (!activeFrame) {
            throw new Error("❌ FATAL: Could not find Export button in ANY frame after 60s.");
        }

        // 5.2 รอ Popup Dialog เด้งขึ้นมา
        console.log('   Waiting for Export Dialog...');
        await new Promise(r => setTimeout(r, 3000));

        // 5.3 เลือก "Microsoft Excel Workbook Data-only"
        // หา Dropdown ใน Frame เดิมที่เจอปุ่ม Export
        const dropdownSelectors = [
            '[id$="_dialog_combo"]', // ปุ่มเปิด Dropdown
            'select'                 // หรือถ้าเป็น select ธรรมดา
        ];

        // พยายามเปิด Dropdown ก่อน
        try {
            const dropdownBtn = await activeFrame.$(dropdownSelectors[0]);
            if (dropdownBtn) {
                console.log('   Clicking Dropdown Arrow...');
                await dropdownBtn.click();
                await new Promise(r => setTimeout(r, 1000));
            }
        } catch(e) {}

        // เลือกเมนู Excel (ใช้ XPath หา Text เอา ชัวร์สุด)
        console.log('   Selecting Excel Data-only...');
        const excelOptionXPath = "//*[contains(text(), 'Microsoft Excel Workbook Data-only')] | //*[contains(text(), 'Excel')]";
        const excelOptions = await activeFrame.$x(excelOptionXPath);

        if (excelOptions.length > 0) {
            await excelOptions[0].click();
        } else {
            console.warn('⚠️ Text option not found, trying Fallback (Keyboard ArrowDown)...');
            // ถ้าหาไม่เจอจริงๆ กดลูกศรลง 2 ทีแล้ว Enter (สูตรกันตาย)
            await page.keyboard.press('ArrowDown');
            await page.keyboard.press('ArrowDown');
            await page.keyboard.press('Enter');
        }
        
        await new Promise(r => setTimeout(r, 2000));

        // 5.4 กดปุ่ม Export (ปุ่มยืนยันสุดท้าย)
        const submitBtnSelector = '[id$="_dialog_submitBtn"]';
        const submitBtn = await activeFrame.$(submitBtnSelector);
        
        if (submitBtn) {
            console.log('   Clicking Final Export Button...');
            await submitBtn.click();
        } else {
            // ถ้าหาปุ่มไม่เจอ ลองกด Enter อีกทีเผื่อ Popup Active อยู่
            console.warn('⚠️ Submit button not found, pressing Enter...');
            await page.keyboard.press('Enter');
        }

        // ---------------------------------------------------------
        // 6. Wait for Download
        // ---------------------------------------------------------
        console.log('⬇️ Waiting for file...');
        let downloadedFile;
        for (let i = 0; i < 90; i++) { // รอ 90 วินาที
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
