const TelegramBot = require('node-telegram-bot-api');
const XLSX = require('xlsx');
const fs = require('fs');
const Dropbox = require('dropbox').Dropbox;
const fetch = require('isomorphic-fetch');

// التوكن الخاص بالبوت من BotFather
const TELEGRAM_TOKEN = 'YOUR_BOT_TOKEN';
// توكن Dropbox
const DROPBOX_ACCESS_TOKEN = '8026253210:AAEedpGTkUA8GevbVOQhkysAIWz5v5U9ovg';

// إعداد البوت
const bot = new TelegramBot(TELEGRAM_TOKEN, { polling: true });

// إعداد اتصال Dropbox
const dbx = new Dropbox({ accessToken: DROPBOX_ACCESS_TOKEN, fetch: fetch });

// دالة لتحميل الملف إلى Dropbox
async function uploadToDropbox(filePath, dropboxPath) {
    const fileContents = fs.readFileSync(filePath);
    try {
        const response = await dbx.filesUpload({ path: dropboxPath, contents: fileContents, mode: { '.tag': 'overwrite' } });
        return response;
    } catch (error) {
        console.error('Error uploading to Dropbox:', error);
        return null;
    }
}

// دالة للبحث في ملف Excel
function searchExcel(filePath, searchTerm) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]]; // افترض أن البيانات في الورقة الأولى
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
        const dropboxPath = `/your_folder/${fileName}`;
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
    const searchTerm = match[1]; // الكلمة التي سيتم البحث عنها

    const filePath = './example.xlsx'; // افترض أن الملف موجود في الخادم لديك أو Dropbox
    
    // هنا يمكنك تحميل الملف من Dropbox إذا كان في المجلد
    // أو يمكنك تحميله محلياً مثل المثال هنا

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
