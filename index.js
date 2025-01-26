require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');
const fs = require('fs');

const token = process.env.BOT_TOKEN;
const bot = new TelegramBot(token, { polling: true });
const usersData = {};
let currentQuestion = 0;

const questions = [
  'Как вас зовут?',
  'Уточните Ваш рабочий номер телефона.',
  'Вы собственник парка? Диспетчер? Водитель на своей машине? Транспортная компания? Если “Другое”, напишите подробнее',
  'Тип ТС (Реф/не Реф, Габариты кузова “Д*Ш*В”, Объём, max кол-во паллет)?',
  'Кол-во ТС в собственности?',
  'Какой маршрут интересен? (СПб ⇄ СПБ / СПБ ⇄ МСК / МСК ⇄ МСК / СПб ⇄ Регионы)',
  'Вы ИП, ООО или Самозанятый?',
  'Работает с НДС или без НДС?',
  'Ваше предложение по ставке.',
];

const startMessage = `Добрый день, благодарим Вас за внимание. 
Наш бот-помощник создан для сбора информации о будующих партнерах.
Огромная просьба, ответьте на все вопросы и наш логист свяжется с вами в ближайшее время.
Опрос займет не более минуты.
Спасибо!
`;

const finishMessage = `Спасибо за обратную связь и ваше уделённое время, наш логист свяжется в ближайшее время.
Хорошего дня!`;

// Замените на ID администратора(ов) вашего бота
const ADMIN_IDS = [+process.env.YOUR_ADMIN_ID];

bot.onText(/\/start/, async (msg) => {
  const chatId = msg.chat.id;
  usersData[chatId] = {};
  currentQuestion = 0;
  await bot.sendMessage(chatId, startMessage);
  await bot.sendMessage(chatId, questions[currentQuestion]);
});

bot.on('message', (msg) => {
  const chatId = msg.chat.id;
  const text = msg.text;
  if (text === '/start') return;
  if (!usersData[chatId]) {
    bot.sendMessage(chatId, 'Пожалуйста, введите /start, чтобы начать.');
    return;
  }

  if (text === '/get_excel') {
    return handleGetExcel(msg);
  }

  usersData[chatId][`question_${currentQuestion}`] = text;
  currentQuestion++;

  if (currentQuestion < questions.length) {
    bot.sendMessage(chatId, questions[currentQuestion]);
  } else {
    bot.sendMessage(chatId, finishMessage);
    saveToExcel(usersData[chatId]);
    delete usersData[chatId];
  }
});

async function saveToExcel(userData) {
  const workbook = new ExcelJS.Workbook();
  let worksheet;
  let nextRow;
  try {
    await workbook.xlsx.readFile('partners_data.xlsx');
    worksheet = workbook.getWorksheet('Partners');
    nextRow = worksheet.rowCount + 1;
  } catch (error) {
    worksheet = workbook.addWorksheet('Partners');
    worksheet.columns = [
      { header: 'Имя', key: 'name', width: 30 },
      { header: 'Телефон', key: 'phone', width: 20 },
      { header: 'Инфо о партнере', key: 'infopartner', width: 30 },
      { header: 'Тип ТС', key: 'TS', width: 40 },
      { header: 'кол-во ТС', key: 'kolvoTS', width: 40 },
      { header: 'маршрут', key: 'marshrutTS', width: 40 },
      { header: 'ЮР форма', key: 'formaTS', width: 40 },
      { header: 'НДС/без НДС', key: 'wwwTS', width: 40 },
      { header: 'Ставка', key: 'stavkaTS', width: 40 },
    ];
    nextRow = 2;
  }

  const row = {
    name: userData.question_0,
    phone: userData.question_1,
    infopartner: userData.question_2,
    TS: userData.question_3,
    kolvoTS: userData.question_4,
    marshrutTS: userData.question_5,
    formaTS: userData.question_6,
    wwwTS: userData.question_7,
    stavkaTS: userData.question_8,
  };

  worksheet.addRow(row);
  worksheet.getRow(nextRow).values = Object.values(row);

  try {
    await workbook.xlsx.writeFile('partners_data.xlsx');
    console.log('Data saved to partners_data.xlsx');
  } catch (error) {
    console.error('Error saving to Excel file:', error);
  }
}

async function handleGetExcel(msg) {
  const chatId = msg.chat.id;
  if (!ADMIN_IDS.includes(msg.from.id)) {
    return bot.sendMessage(
      chatId,
      'У вас нет прав на выполнение этой команды.'
    );
  }

  try {
    // Проверяем существует ли файл
    fs.accessSync('partners_data.xlsx', fs.constants.F_OK);
    await bot.sendDocument(chatId, 'partners_data.xlsx', {
      caption: 'Файл с данными партнеров',
    });
  } catch (e) {
    console.log(e);
    return bot.sendMessage(chatId, 'Файл с данными не найден.');
  }
}
