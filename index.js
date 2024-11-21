// استخدام require بدلاً من import للتوافق مع CommonJS
const express = require('express');  // استيراد مكتبة Express
const TelegramBot = require('node-telegram-bot-api');
const XLSX = require('xlsx');
const fs = require('fs');
const { Dropbox } = require('dropbox');
const fetch = require('isomorphic-fetch');  // لاستخدام fetch مع Dropbox

// التوكن الخاص بالبوت من BotFather
const TELEGRAM_TOKEN = '8026253210:AAEedpGTkUA8GevbVOQhkysAIWz5v5U9ovg';
// توكن Dropbox
const DROPBOX_ACCESS_TOKEN = 'sl.CBITRY3JwaiwBuKYFLL0vrKK9Vamj9rIw9Mee3q4Q02rWMgJHIvNoax9rawbTWpGbtwVf5ZLE4Q3f_sAhlvCyUCGAZYNpdQlIX0vNc1F5SqRD8Vruf-2zM-tQlgqB0NPdY5eerWmgY6G';

// إعداد البوت
const bot = new TelegramBot(TELEGRAM_TOKEN, { polling: true });

// إعداد اتصال Dropbox
const dbx = new Dropbox({ accessToken: DROPBOX_ACCESS_TOKEN, fetch: fetch });

// إنشاء خادم Express
const app = express();

// إعداد البورت
const PORT = process.env.PORT || 3000;  // استخدام المنفذ من متغير البيئة أو المنفذ الافتراضي 3000

// دالة لتحميل الملف إلى Dropbox
async function uploadToDropbox(filePath, dropboxPath) {
    const fileContents = fs.readFileSync(filePath);
    try {
        const response = await dbx.filesUpload({
            path: dropboxPath, // المسار في Dropbox
            contents: fileContents,
            mode: { '.tag': 'overwrite' } // إذا كان الملف موجودًا، يتم استبداله
        });
        return response;
    } catch (error) {
        console.error('Error uploading to Dropbox:', error);
        return null;
    }
}

// دالة للبحث في ملف Excel
function searchExcel(filePath, searchTerm) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // اختيار الورقة الأولى
    const sheet = workbook.Sheets[sheetName];

    // تحويل الورقة إلى JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    return jsonData.filter(row => {
        return Object.values(row).some(value => {
            return String(value).toLowerCase().includes(searchTerm.toLowerCase());
        });
    });
}

// معالجة الرسائل
bot.onText(/\/start/, (msg) => {
    bot.sendMessage(msg.chat.id, 'مرحبًا! يمكنك إرسال ملف Excel لأقوم برفعه إلى Dropbox أو البحث فيه.');
});

// معالجة رفع الملفات
bot.on('document', async (msg) => {
    const chatId = msg.chat.id;
    const file = msg.document;
    const fileId = file.file_id;
    const fileName = file.file_name;

    // تحميل الملف من تيليجرام
    const fileUrl = await bot.getFileLink(fileId);
    const filePath = `./${fileName}`;
    
    const fileStream = fs.createWriteStream(filePath);
    const response = await fetch(fileUrl);
    response.body.pipe(fileStream);

    fileStream.on('finish', async () => {
        // رفع الملف إلى Dropbox
        const dropboxPath = `/Apps/gazatest/${fileName}`;  // تعديل المسار هنا لرفع الملف إلى /Apps/gazatest
        const dropboxResponse = await uploadToDropbox(filePath, dropboxPath);

        if (dropboxResponse) {
            bot.sendMessage(chatId, `تم رفع الملف بنجاح إلى Dropbox: ${dropboxPath}`);
        } else {
            bot.sendMessage(chatId, 'حدث خطأ أثناء رفع الملف إلى Dropbox.');
        }

        // تنظيف الملفات المحلية
        fs.unlinkSync(filePath);
    });
});

// معالجة البحث في الملف
bot.onText(/\/search (.+)/, (msg, match) => {
    const chatId = msg.chat.id;
    const searchTerm = match[1];  // الكلمة التي سيتم البحث عنها

    const filePath = './example.xlsx';  // افترض أن الملف موجود في الخادم لديك أو Dropbox
    
    const results = searchExcel(filePath, searchTerm);
    if (results.length > 0) {
        let resultText = 'النتائج:\n';
        results.forEach((row, index) => {
            resultText += `نتيجة ${index + 1}: ${JSON.stringify(row)}\n`;
        });
        bot.sendMessage(chatId, resultText);
    } else {
        bot.sendMessage(chatId, 'لم يتم العثور على نتائج.');
    }
});

// استعراض الأخطاء
bot.on('polling_error', (error) => {
    console.log(error);
});

// تشغيل خادم Express
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
