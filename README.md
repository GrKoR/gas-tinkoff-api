# GAS Tinkoff Trades
![GAS Tinkoff Trades main image](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/main-image.jpg "GAS Tinkoff Trades main image")

Данный [Google Apps Script](https://developers.google.com/apps-script) предназначен для импорта сделок, портфеля и текущих цен из Тинькофф Инвестиций прямо в Google таблицы для последующего анализа. 

Я сделал этот скрипт для автоматизации ручного вбивания данных из приложения тинькофф, и надеюсь он окажется полезен кому-нибудь ещё :)

## Установка

* Создать или открыть документ Google Spreadsheets http://drive.google.com
* В меню "Tools" выбрать "Script Editor"
* Дать проекту имя, например `TinkoffTrades`
* Скопировать код из [Code.gs](https://raw.githubusercontent.com/ErhoSen/gas-tinkoff-trades/master/Code.gs) и заменить им дефолтный текст скрипта.
* Создать второй файл скрипта: в панели слева "Files" -> "+" -> "Script". Назвать его `Tinkoff-API.gs`.
* В этот файл скопировать код из [Tinkoff-API.gs](https://raw.githubusercontent.com/GrKoR/gas-tinkoff-api/master/Tinkoff-API.gs).
* Получить [OpenApi-токен тинькофф](https://tinkoffcreditsystems.github.io/invest-openapi/auth/)
* Добавить свойство `OPENAPI_TOKEN` в разделе `Project properties -> Script properties` равным токену, полученному выше.
* Сохранить скрипт 💾
* В документе Google Spreadsheets выбрать любую ненужную ячейку и присвоить ей имя `UPDATE_DATE` с помощью меню `Data`->`Named ranges`->`Add named range`. В эту ячейку по команде меню TI->Обновить вставляется текущая дата. Данная ячейка может использоваться в качестве необязательного параметра `dummy` любой функции для [принудительного обновления формул](https://stackoverflow.com/a/27656313).

На этом всё. Теперь при работе с этим документом на всех листах будут доступны функции:
* API v1: `getPriceByTicker()`, `getTrades()`, `getPortfolio()`, `getMaxBidByTicker()`, `getMinAskByTicker()` и `getBidAskSpread()`
* API v2: `TI_GetAccounts()`, `TI_GetInstrumentsID()`, `TI_GetLastPrice()`, `TI_GetPortfolio()`, `TI_GetOperations()`, `TI_GetBidAskSpread()`

## Функции API v2

* `=TI_GetOperations(accountId, from, to)` - выводит операции по заданному номеру счета (можно получить через функцию `TI_GetAccounts()`) и, опционально, фильтрацию по времени. Параметры `from` и `to` являются ссылками на ячейки с типом Дата.

* `=TI_GetPortfolio(accountId)` - выводит портфель по заданному номеру счета (можно получить через функцию `TI_GetAccounts()`.

## Функции API v1 (перестанет работать в 2023 году)

* `=getPriceByTicker(ticker, dummy)` - требует на вход [тикер](https://ru.wikipedia.org/wiki/%D0%A2%D0%B8%D0%BA%D0%B5%D1%80), и опциональный параметр `dummy`. Для автоматичекого обновления необходимо указать в качестве `dummy` ячейку `Z1` (которая будет обновлятся через меню TI).

* `=getTrades(ticker, from, to)` - требует на вход [тикер](https://ru.wikipedia.org/wiki/%D0%A2%D0%B8%D0%BA%D0%B5%D1%80), и опционально фильтрацию по времени. Параметры `from` и `to` являются строками и должны быть в [ISO 8601 формате](https://ru.wikipedia.org/wiki/ISO_8601)

* `=getPortfolio()` - выводит портфель.

## Особенности

* Среди настроек скрипта есть `TRADING_START_AT` - дефолтная дата, начиная с которой фильтруются операции `getTrades`. По умолчанию это `Apr 01, 2020 10:00:00`, но данную константу можно в любой момент поменять в исходном коде.

## Пример использования работы функций API v1

```
=getPriceByTicker("V", UPDATE_DATE)  # Возвращает текущую цену акции Visa
=getPriceByTicker("FXMM", UPDATE_DATE)  # Возвращает текущую цену фонда казначейских облигаций США

=getTrades("V") 
# Вернёт все операции с акцией Visa, которые произошли начиная с TRADING_START_AT и по текущий момент.
=getTrades("V", "2020-05-01T00:00:00.000Z") 
# Вернёт все операции с акцией Visa, которые произошли начиная с 1 мая и по текущий моментs.
=getTrades("V", "2020-05-01T00:00:00.000Z", "2020-05-05T23:59:59.999Z") 
# Вернёт все операции с акцией Visa, которые произошли в период с 1 и по 5 мая.
```

## Пример работы функций API v1

#### `=getTrades()`
![getTrades in action](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/get-trades-in-action.gif "getTrades in Action")

#### `=getPriceByTicker()`
![Get price by ticker in action](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/get-price-by-ticker.gif)
