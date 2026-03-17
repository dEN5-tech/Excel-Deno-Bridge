/// <reference path="./excel.d.ts" />

// Наша главная функция
globalThis.onExcelEvent = async (name, data) => {
    // Реагируем только на изменения ячеек
    if (name !== 'cell_change' || !data) return;

    Excel.StatusBar = `Обработка: ${data}...`;

    // Пример 1: Автоматическая раскраска (Умное форматирование)
    if (data.toLowerCase() === "ошибка") {
        Excel.ActiveCell.Interior.Color = 0x0000FF; // Красный фон
        Excel.ActiveCell.Font.Color = 0xFFFFFF;     // Белый текст
    } 
    else if (data.toLowerCase() === "ок") {
        Excel.ActiveCell.Interior.Color = 0x00FF00; // Зеленый фон
        Excel.ActiveCell.Font.Color = 0x000000;     // Черный текст
    }

    if (data.toLowerCase().startsWith("курс")) {
        const coin = (data.split(" ")[1] || "BTC").toUpperCase(); 
        const symbol = `${coin}USDT`;

        // Пишем статус прямо в ячейку, пока ждем
        Excel.ActiveCell.Value = `Загрузка ${coin}...`;

        try {
            const res = await fetch(`https://api.binance.com/api/v3/ticker/price?symbol=${symbol}`);
            const json = await res.json();
            
            if (json.price) {
                const price = parseFloat(json.price).toFixed(2);
                
                // ЗАМЕНЯЕМ текст в текущей ячейке на результат
                Excel.ActiveCell.Value = `💰 Курс ${coin}: $${price}`;
                
                // Делаем ячейку красивой (жирный текст, желтый фон)
                Excel.ActiveCell.Interior.Color = 0x00FFFF; // Желтый в формате BGR
                Excel.ActiveCell.Font.Bold = true;
                
                Excel.StatusBar = `Курс загружен успешно!`;
            } else {
                throw new Error("Монета не найдена");
            }
        } catch (e) {
            Excel.StatusBar = "Ошибка: " + e.message;
            Excel.ActiveCell.Value = "❌ Ошибка загрузки";
            Excel.ActiveCell.Interior.Color = 0x0000FF; // Красный
        }
    }
};

// Включаем монитор событий
startExcelEvents();