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

// 🛠️ ฟังก์ชันพิเศษ: ค้นหา Element ในทุก Frame (รองรับ Dynamic ID)
async function findElementRecursive(page, selectors, timeout = 60000) {
    const start = Date.now();
    console.log(`   🕵️‍♂️ Searching for: ${selectors[0]}...`);
    
    while (Date.now() - start < timeout) {
        const frames = page.frames();
        for (const frame of frames) {
            for (const selector of selectors) {
                try {
                    let el;
                    if (selector.startsWith('//')) {
                        const els = await frame.$x(selector);
                        if (els.length > 0) el = els[0];
                    } else {
                        el = await frame.$(selector);
                    }

                    if (el) {
                        const isVisible = await el.evaluate(e => {
                            const style = window.getComputedStyle(e);
                            return style && style.display !== 'none' && style.visibility !== 'hidden';
                        });
                        
                        if (isVisible) {
                            console.log(`   ✅ Found element in frame: "${frame.name() || 'unnamed'}"`);
                            return { frame, element: el };
                        }
                    }
                } catch (e) { }
            }
        }
        await new Promise(r => setTimeout(r, 1000));
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
        // 3. Fill Form & Date Logic (UI.Vision Lines 2-9)
        // ---------------------------------------------------------
        console.log('📝 Filling form...');
        // เลือกประเภทเอกสาร
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_ddlDoctype', { visible: true });
        await page.select('#ctl00_ContentPlaceHolder1_ddlDoctype', '1');
        await new Promise(r => setTimeout(r, 3000));

        // เลือก OT
        const otTypeSelector = '#ctl00_ContentPlaceHolder1_ddlOt';
        if (await page.$(otTypeSelector) !== null) {
            await page.select(otTypeSelector, '14');
            await new Promise(r => setTimeout(r, 2000));
        }

        // ใส่วันที่ (ใช้วันที่ 1 ของเดือนตามที่คุณต้องการ)
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
        // 4. Generate Report & SWITCH TAB (UI.Vision Line 13-14)
        // ---------------------------------------------------------
        console.log('⏳ Generating Report...');
        
        // 🟢 เตรียมดักจับหน้าต่างใหม่ (New Tab) - หัวใจสำคัญ!
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        
        // กดปุ่มแสดงรายงาน
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');
        
        // 🟢 รอและสลับไปหน้าใหม่
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        console.log('✅ Switched to New Report Tab (ReportView.aspx)');
        await reportPage.bringToFront();
        await reportPage.setViewport({ width: 1366, height: 768 });
        
        // รอโหลดหน้าใหม่
        console.log('   Waiting 15s for Crystal Report Iframe...');
        await new Promise(r => setTimeout(r, 15000));

        // Setup Download บนหน้าใหม่
        const reportClient = await reportPage.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // 5. Crystal Report Export (อ้างอิง UI.Vision Lines 15-18)
        // ---------------------------------------------------------
        console.log('💾 Handling Crystal Report Export...');

        // 5.1 กดปุ่ม Toolbar (UI.Vision Line 15)
        // ใช้ ID ภาษาไทยจากไฟล์ text หรือ Partial Match
        const iconSelectors = [
            '[id*="IconImg_รายงานรายละเอียดขออนุมัติใบโอทีของพนักงาน_toptoolbar_export"]', 
            '[id$="_toptoolbar_export"]' // สำรอง
        ];
        
        console.log('   1. Searching for Toolbar Export Icon...');
        // **ใช้ reportPage ในการค้นหา**
        const iconFound = await findElementRecursive(reportPage, iconSelectors, 60000);
        
        if (!iconFound) throw new Error("Could not find Toolbar Export Icon!");
        
        await iconFound.element.evaluate(el => el.click());
        const activeFrame = iconFound.frame; 

        // 5.2 รอ Dialog และกด Dropdown (UI.Vision Line 16)
        console.log('   Waiting for Dialog...');
        await new Promise(r => setTimeout(r, 3000));

        console.log('   2. Clicking Dropdown Arrow...');
        const dropdownSelectors = ['[id$="_dialog_combo"]']; // จาก UI.Vision
        const dropdownFound = await findElementRecursive(reportPage, dropdownSelectors, 10000);
        
        if (dropdownFound) {
            await dropdownFound.element.evaluate(el => el.click());
            await new Promise(r => setTimeout(r, 1000));
        }

        // 5.3 เลือก Excel Data-only (UI.Vision Line 17)
        // ใช้ ID ที่ลงท้ายด้วย _it_14 (ตามไฟล์ Text เป๊ะๆ)
        console.log('   3. Selecting Excel Data-only...');
        
        const excelOptionID = await activeFrame.$('[id$="_dialog_combo_it_14"]');
        
        if (excelOptionID) {
            console.log('   ✅ Found option by ID suffix (_it_14)');
            await excelOptionID.click();
        } else {
            // สำรอง: หาด้วย Text
            console.log('   ⚠️ ID not found, trying text match...');
            const excelOptionText = await activeFrame.$x("//*[contains(text(), 'Microsoft Excel Workbook Data-only')]");
            if (excelOptionText.length > 0) {
                await excelOptionText[0].click();
            } else {
                console.warn('⚠️ Option not found, using ArrowDown Fallback...');
                await reportPage.keyboard.press('ArrowDown');
                await reportPage.keyboard.press('ArrowDown');
                await reportPage.keyboard.press('Enter');
            }
        }
        await new Promise(r => setTimeout(r, 2000));

        // 5.4 กดปุ่ม Export สุดท้าย (UI.Vision Line 18)
        console.log('   4. Clicking Final Export Button...');
        const finalBtnSelectors = [
            '//a[text()="Export"]',        // ตาม linkText ใน UI.Vision
            'a[id$="_dialog_submitBtn"]',  // ตาม ID ใน UI.Vision
            '[id$="_dialog_submitBtn"]'
        ];
        
        const finalBtnFound = await findElementRecursive(reportPage, finalBtnSelectors);
        
        if (finalBtnFound) {
            console.log('   ✅ Found Final Button! Clicking...');
            await finalBtnFound.element.evaluate(el => el.click());
        } else {
            console.warn('⚠️ Final Button not found. Pressing Enter...');
            await reportPage.keyboard.press('Enter');
        }

        // ---------------------------------------------------------
        // 6. Wait for Download
        // ---------------------------------------------------------
        console.log('⬇️ Waiting for file...');
        let downloadedFile;
        for (let i = 0; i < 120; i++) { 
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
