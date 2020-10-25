處理檔案命名為 source.xls 輸出檔案為 out.csv

執行指令為 npm start

檔案內容設定說明

SHEETNAME 為表名(現在為「資料」)

RANGE 為資料處理範圍 `const RANGE = 'A2:M992'` A2, M992 為範圍內的左上&右下角

現況排除資料的檢查為 `checkUnit` 此 function 如需增設條件可至該 function 內增補

預設第一欄呈現為「尺碼」(變數為 `const SIZE=‘尺碼’`

FAQ

如需增補丈量尺寸，該怎做?

Ans. 請修正 `RANGE='範圍'` 參數即可 eg. A2M10 -> A2N100 A2, M10 為範圍內的左上&右下角
