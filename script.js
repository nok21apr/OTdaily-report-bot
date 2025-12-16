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
    console.log('🚀 Starting Bot (Business Plus - Fix Login Mode)...');

    // ตรวจสอบตัวแปรสำคัญ
    if (!WEB_USER || !WEB_PASS || !EMAIL_USER || !EMAIL_PASS) {
        console.error('❌ Error: Secrets incomplete.');
        process.exit(1);
    }

    const downloadPath = path.join(__dirname, 'downloads');
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    let browser = null;
    let page = null;

    try {
        console.log('🖥️ Launching Browser...');
        browser = await puppeteer.launch({
            headless: "new",
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--window-size=1920,1080',
                '--lang=th-TH,th'
            ]
        });

        page = await browser.newPage();
        
        // Timeout 5 นาที
        page.setDefaultNavigationTimeout(300000);
        page.setDefaultTimeout(300000);

        await page.emulateTimezone('Asia/Bangkok');
        const client = await page.target().createCDPSession();
        await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });

        // ---------------------------------------------------------
        // Step 1: Login (แก้ไขใหม่)
        // ---------------------------------------------------------
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://leave.ttkasia.co.th/Login/Login.aspx', { waitUntil: 'networkidle2' });
        
        // รอช่อง User
        await page.waitForSelector('#txtUsername', { visible: true });
        
        // พิมพ์ User/Pass
        await page.type('#txtUsername', WEB_USER.trim(), { delay: 50 });
        await page.type('#txtPassword', WEB_PASS.trim(), { delay: 50 });
        
        console.log('   Clicking Login Button...');
        
        // 🛠️ FIX: คลิกปุ่ม Login แทนการกด Enter (หาปุ่มที่มีคำว่า submit หรือ id ที่น่าจะเป็นปุ่ม)
        // ลองหา Selector ของปุ่ม Login (ส่วนใหญ่ Business Plus ใช้ #btnLogin หรือ input[type=submit])
        const loginBtnSelector = '#btnLogin, input[type="submit"], button[type="submit"], a.wizbutton';
        await page.waitForSelector(loginBtnSelector, { visible: true, timeout: 5000 }).catch(() => console.log("   Warning: Could not find explicit login button, trying Enter..."));

        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 60000 }), // รอโหลดหน้าใหม่
            (async () => {
                // พยายามคลิกปุ่ม ถ้าหาไม่เจอให้กด Enter
                const btn = await page.$(loginBtnSelector);
                if (btn) {
                    await btn.click();
                } else {
                    await page.keyboard.press('Enter');
                }
            })()
        ]);

        // 🔍 CHECK: ตรวจสอบว่า Login ผ่านจริงไหม?
        const currentUrl = page.url();
        console.log(`   Current URL: ${currentUrl}`);
        
        // ถ้า URL ยังเป็น Login.aspx แสดงว่าเข้าไม่ได้
        if (currentUrl.includes('Login.aspx')) {
            console.error('❌ Login Failed: ยังอยู่ที่หน้า Login');
            await page.screenshot({ path: path.join(downloadPath, 'login_failed.png'), fullPage: true });
            throw new Error('Login Failed - Still on login page');
        }

        console.log('✅ Login Success (Dashboard reached)');

        // ---------------------------------------------------------
        // Step 2: Navigate to Report
        // ---------------------------------------------------------
        console.log('2️⃣ Step 2: Go to Report Page...');
        
        // คลิกเมนูหลัก (รอ selector นานหน่อยเผื่อหน้า Dashboard โหลดช้า)
        const mainMenuSelector = '#ctl00_ContentPlaceHolder1_imgLeave';
        try {
            await page.waitForSelector(mainMenuSelector, { visible: true, timeout: 30000 });
        } catch (e) {
            throw new Error(`หาเมนูหลักไม่เจอ (${mainMenuSelector}) - อาจจะ Login ไม่สมบูรณ์`);
        }
        await page.click(mainMenuSelector);
        
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
        
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_ddlDoctype');
        await page.select('#ctl00_ContentPlaceHolder1_ddlDoctype', '1');
        await page.select('#ctl00_ContentPlaceHolder1_ddlOt', '14');

        const now = new Date();
        const thaiYear = now.getFullYear() + 543;
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const firstDayValue = `01/${month}/${thaiYear}`;
        
        console.log(`   Setting date to: ${firstDayValue}`);
        await page.click('#ctl00_ContentPlaceHolder1_txtFromDate', { clickCount: 3 });
        await page.type('#ctl00_ContentPlaceHolder1_txtFromDate', firstDayValue);

        // ---------------------------------------------------------
        // Step 4: Generate Report
        // ---------------------------------------------------------
        console.log('4️⃣ Step 4: Generating Report...');
        
        const newPagePromise = new Promise(x => browser.once('targetcreated', target => x(target.page())));
        await page.click('#ctl00_ContentPlaceHolder1_lnkShowReport');
        
        const reportPage = await newPagePromise;
        if (!reportPage) throw new Error("Report tab did not open!");
        
        page = reportPage; 
        await page.bringToFront();
        await page.setViewport({ width: 1920, height: 1080 });
        
        const reportClient = await page.target().createCDPSession();
        await reportClient.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        // ---------------------------------------------------------
        // Step 5: Export Handling
        // ---------------------------------------------------------
        console.log('5️⃣ Step 5: Handling Export...');
        
        const exportBtnSelector = 'a[title="Export this report"], img[alt="Export this report"]';
        await page.waitForSelector(exportBtnSelector, { visible: true, timeout: 60000 });
        await page.click(exportBtnSelector);

        await page.waitForSelector('select', { visible: true });
        
        const selectId = await page.evaluate(() => {
            const options = Array.from(document.querySelectorAll('option'));
            const target = options.find(o => o.text.includes('Microsoft Excel Workbook Data-only'));
            return target ? target.parentElement.id : null;
        });
        
        if (selectId) {
            await page.select(`#${selectId}`, 'Microsoft Excel Workbook Data-only');
        } else {
            await page.keyboard.press('ArrowDown');
        }

        const finalSubmitSelector = 'a[id$="_dialog_submitBtn"]';
        await page.waitForSelector(finalSubmitSelector);
        await page.click(finalSubmitSelector);

        // ---------------------------------------------------------
        // Step 6: Wait for Download
        // ---------------------------------------------------------
        console.log('6️⃣ Step 6: Waiting for file...');
        let downloadedFile = null;
        
        for (let i = 0; i < 300; i++) {
            await new Promise(r => setTimeout(r, 1000));
            const files = fs.readdirSync(downloadPath);
            const target = files.find(f => (f.endsWith('.xlsx') || f.endsWith('.xls')) && !f.endsWith('.crdownload'));
            if (target) {
                downloadedFile = target;
                break;
            }
            if (i > 0 && i % 10 === 0) console.log(`   ...still waiting (${i}s)`);
        }

        if (!downloadedFile) throw new Error('❌ Download Timeout.');
        const originalFilePath = path.join(downloadPath, downloadedFile);
        console.log(`✅ File Downloaded: ${originalFilePath}`);

        await browser.close();
        browser = null; 

        // ---------------------------------------------------------
        // Step 7: Convert & Email
        // ---------------------------------------------------------
        console.log('🔄 Step 7: Converting & Emailing...');
        
        const workbook = xlsx.readFile(originalFilePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const csvContent = xlsx.utils.sheet_to_csv(worksheet);
        
        const csvFileName = downloadedFile.replace(/\.[^/.]+$/, "") + ".csv";
        const csvFilePath = path.join(downloadPath, csvFileName);

        fs.writeFileSync(csvFilePath, '\uFEFF' + csvContent, { encoding: 'utf8' });

        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: { user: EMAIL_USER, pass: EMAIL_PASS }
        });

        await transporter.sendMail({
            from: `"Auto Report Bot" <${EMAIL_USER}>`,
            to: EMAIL_TO,
            subject: `รายงาน Business Plus - ${new Date().toLocaleDateString()}`,
            text: `ดาวน์โหลดสำเร็จ\nไฟล์: ${csvFileName}`,
            attachments: [{ filename: csvFileName, path: csvFilePath }]
        });
        
        console.log('🎉 Mission Complete!');

    } catch (error) {
        console.error('❌ FATAL ERROR:', error);
        if (page && !page.isClosed()) {
            try { 
                await page.screenshot({ path: path.join(downloadPath, 'fatal_error.png'), fullPage: true });
                console.log('📸 Screenshot saved: fatal_error.png');
            } catch(e){}
        }
        if (browser) await browser.close();
        process.exit(1);
    }
})();
